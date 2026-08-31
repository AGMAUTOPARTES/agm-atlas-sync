"""
AGM Auto Partes - Bling API v3 - Autenticacao OAuth2
======================================================
Responsavel por:
  - Fluxo de autorizacao inicial (abre o navegador, recebe o "code" via
    servidor HTTPS local em https://localhost:8765/callback, troca por
    access_token + refresh_token).
  - Renovar o access_token automaticamente usando o refresh_token quando
    expira (sem precisar logar de novo).
  - Persistir os tokens em bling_tokens.json (NUNCA versionar esse arquivo).

Uso:
    python bling_auth.py            -> roda a autorizacao inicial (1a vez)
    from bling_auth import get_valid_access_token   -> uso programatico

Pre-requisitos (arquivo .env na mesma pasta):
    BLING_CLIENT_ID=...
    BLING_CLIENT_SECRET=...
    BLING_REDIRECT_URI=https://localhost:8765/callback
    BLING_BASE_URL=https://api.bling.com.br/Api/v3
    BLING_TOKEN_FILE=bling_tokens.json
"""

import base64
import datetime as dt
import http.server
import json
import os
import secrets
import ssl
import threading
import time
import urllib.parse
import webbrowser
from pathlib import Path

import requests

import token_store

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

BASE_DIR = Path(__file__).resolve().parent

CLIENT_ID = os.environ.get("BLING_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("BLING_CLIENT_SECRET", "")
REDIRECT_URI = os.environ.get("BLING_REDIRECT_URI", "https://localhost:8765/callback")
BASE_URL = os.environ.get("BLING_BASE_URL", "https://api.bling.com.br/Api/v3")
TOKEN_FILE = BASE_DIR / os.environ.get("BLING_TOKEN_FILE", "bling_tokens.json")
CERT_FILE = BASE_DIR / "_localhost_cert.pem"
KEY_FILE = BASE_DIR / "_localhost_key.pem"

TOKEN_URL = "https://www.bling.com.br/Api/v3/oauth/token"
AUTHORIZE_URL = "https://www.bling.com.br/Api/v3/oauth/authorize"

_parsed = urllib.parse.urlparse(REDIRECT_URI)
CALLBACK_HOST = _parsed.hostname or "localhost"
CALLBACK_PORT = _parsed.port or 8765
CALLBACK_PATH = _parsed.path or "/callback"


def _check_config():
    if not CLIENT_ID or not CLIENT_SECRET:
        raise RuntimeError(
            "BLING_CLIENT_ID / BLING_CLIENT_SECRET nao encontrados. "
            "Confira o arquivo .env nesta pasta (veja .env.example)."
        )


def _basic_auth_header():
    raw = f"{CLIENT_ID}:{CLIENT_SECRET}".encode("utf-8")
    return "Basic " + base64.b64encode(raw).decode("ascii")


def _ensure_self_signed_cert():
    """Gera um certificado autoassinado para o servidor local HTTPS,
    caso ainda nao exista. Necessario porque o Bling exige redirect_uri
    em https mesmo para localhost."""
    if CERT_FILE.exists() and KEY_FILE.exists():
        return
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")])
    cert = (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(dt.datetime.utcnow() - dt.timedelta(days=1))
        .not_valid_after(dt.datetime.utcnow() + dt.timedelta(days=3650))
        .add_extension(
            x509.SubjectAlternativeName([x509.DNSName("localhost")]), critical=False
        )
        .sign(key, hashes.SHA256())
    )
    KEY_FILE.write_bytes(
        key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.TraditionalOpenSSL,
            serialization.NoEncryption(),
        )
    )
    CERT_FILE.write_bytes(cert.public_bytes(serialization.Encoding.PEM))


class _CallbackHandler(http.server.BaseHTTPRequestHandler):
    result = {}

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != CALLBACK_PATH:
            self.send_response(404)
            self.end_headers()
            return
        qs = urllib.parse.parse_qs(parsed.query)
        _CallbackHandler.result["code"] = qs.get("code", [None])[0]
        _CallbackHandler.result["state"] = qs.get("state", [None])[0]
        _CallbackHandler.result["error"] = qs.get("error_description", [None])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        if _CallbackHandler.result.get("code"):
            html = "<h2>AGM - Bling conectado com sucesso!</h2><p>Pode fechar esta aba e voltar ao terminal.</p>"
        else:
            html = f"<h2>Falha na autorizacao</h2><p>{_CallbackHandler.result.get('error')}</p>"
        self.wfile.write(html.encode("utf-8"))

    def log_message(self, format, *args):  # silencia o log padrao
        pass


def _run_local_server_and_wait(expected_state, timeout=600):
    _ensure_self_signed_cert()
    _CallbackHandler.result = {}
    # Forca IPv4 (127.0.0.1) explicitamente em vez de deixar o SO resolver
    # "localhost" -- em alguns Macs "localhost" resolve para ::1 (IPv6)
    # primeiro, e o navegador pode tentar essa rota antes da IPv4, dando
    # "conexao recusada" mesmo com o servidor de pe (ele so escuta em IPv4).
    bind_host = "127.0.0.1" if CALLBACK_HOST in ("localhost", "127.0.0.1") else CALLBACK_HOST
    server = http.server.HTTPServer((bind_host, CALLBACK_PORT), _CallbackHandler)
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(certfile=str(CERT_FILE), keyfile=str(KEY_FILE))
    server.socket = ctx.wrap_socket(server.socket, server_side=True)
    # server.timeout limita quanto handle_request() espera por UMA conexao
    # antes de devolver o controle -- assim conseguimos tentar de novo caso
    # a primeira conexao recebida nao seja o callback real (ex: uma conexao
    # perdida/antiga, ou uma tentativa espuriosa do navegador).
    server.timeout = 5

    def _serve_until_done():
        deadline = time.time() + timeout
        while time.time() < deadline:
            server.handle_request()
            if _CallbackHandler.result.get("code") or _CallbackHandler.result.get("error"):
                return
        return

    thread = threading.Thread(target=_serve_until_done, daemon=True)
    thread.start()
    thread.join(timeout=timeout + 5)
    server.server_close()

    result = _CallbackHandler.result
    if not result.get("code"):
        raise RuntimeError(f"Nao recebi o codigo de autorizacao a tempo. Detalhe: {result}")
    if expected_state and result.get("state") != expected_state:
        raise RuntimeError("State retornado nao confere (possivel CSRF). Aborte e tente novamente.")
    return result["code"]


def _exchange_code_for_tokens(code):
    resp = requests.post(
        TOKEN_URL,
        headers={
            "Authorization": _basic_auth_header(),
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "1.0",
        },
        data={
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": REDIRECT_URI,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def _refresh_tokens(refresh_token):
    resp = requests.post(
        TOKEN_URL,
        headers={
            "Authorization": _basic_auth_header(),
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "1.0",
        },
        data={
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def _save_tokens(data, previous_refresh_token=None):
    payload = {
        "access_token": data["access_token"],
        # O Bling nem sempre devolve um novo refresh_token a cada renovacao;
        # quando nao devolve, mantem o anterior em vez de gravar "None" por
        # cima (o que travaria a proxima renovacao).
        "refresh_token": data.get("refresh_token") or previous_refresh_token,
        "token_type": data.get("token_type", "Bearer"),
        "expires_in": data.get("expires_in", 21600),
        "obtained_at": dt.datetime.utcnow().isoformat() + "Z",
        "scope": data.get("scope"),
    }
    return token_store.save_tokens(TOKEN_FILE, payload)


def _load_tokens():
    return token_store.load_tokens(TOKEN_FILE)


def _is_expired(tokens, buffer_seconds=120):
    obtained_at = dt.datetime.fromisoformat(tokens["obtained_at"].replace("Z", ""))
    expires_at = obtained_at + dt.timedelta(seconds=tokens["expires_in"] - buffer_seconds)
    return dt.datetime.utcnow() >= expires_at


def run_first_time_authorization():
    """Abre o navegador, recebe o code via servidor local e salva os tokens.
    So funciona numa maquina com navegador e rede local disponiveis (Mac,
    Windows) -- nao roda dentro de um runner do GitHub Actions."""
    if token_store.MODE == "cloud":
        raise RuntimeError(
            "Sem token valido no armazenamento em nuvem (D1) e este e' um "
            "runner sem navegador -- nao da pra abrir o fluxo de autorizacao "
            "aqui. Rode 'python seed_bling_token_to_cloud.py' numa maquina "
            "que ja tem um bling_tokens.json valido pra enviar o token pro "
            "D1, ou refaca a autorizacao localmente (python bling_auth.py) "
            "e rode o seed de novo."
        )
    _check_config()
    state = secrets.token_urlsafe(16)
    params = {
        "response_type": "code",
        "client_id": CLIENT_ID,
        "state": state,
    }
    url = f"{AUTHORIZE_URL}?{urllib.parse.urlencode(params)}"
    print("Abrindo o navegador para voce autorizar o app AGM no Bling...")
    print("Se nao abrir sozinho, acesse manualmente:")
    print(url)
    webbrowser.open(url)
    code = _run_local_server_and_wait(state)
    data = _exchange_code_for_tokens(code)
    tokens = _save_tokens(data)
    print(f"Conectado! Tokens salvos em {TOKEN_FILE.name} (expira em {tokens['expires_in']}s).")
    return tokens


def get_valid_access_token():
    """Retorna um access_token valido, renovando via refresh_token se preciso.
    Se nao houver token salvo ainda, dispara o fluxo de autorizacao inicial."""
    _check_config()
    tokens = _load_tokens()
    if tokens is None:
        tokens = run_first_time_authorization()
        return tokens["access_token"]

    if _is_expired(tokens):
        try:
            data = _refresh_tokens(tokens["refresh_token"])
            tokens = _save_tokens(data, previous_refresh_token=tokens["refresh_token"])
        except requests.HTTPError:
            print("Refresh token invalido/expirado. Refazendo autorizacao inicial...")
            tokens = run_first_time_authorization()
    return tokens["access_token"]


if __name__ == "__main__":
    run_first_time_authorization()
