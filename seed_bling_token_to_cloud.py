"""
Envia o token OAuth do Bling que ja existe localmente (bling_tokens.json)
para o armazenamento em nuvem (D1, via o Worker do Atlas), UMA UNICA VEZ.

Depois desse envio, o GitHub Actions consegue rodar a sincronizacao sem
precisar de navegador nem de arquivo local -- ele le e renova o token
direto na nuvem (ver token_store.py).

Rode isto na mesma maquina onde a autorizacao do Bling ja foi feita
(o Mac, hoje), de dentro da pasta Bling_Sync, com o ambiente virtual
ativado:

    source .venv/bin/activate
    python seed_bling_token_to_cloud.py

Nenhum valor de credencial e' digitado aqui nem aparece impresso -- o
script so le os arquivos .env e bling_tokens.json que ja existem nesta
pasta e os envia direto pro Worker, por HTTPS.
"""
import json
import os
import sys
from pathlib import Path

import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

BASE_DIR = Path(__file__).resolve().parent
TOKEN_FILE = BASE_DIR / os.environ.get("BLING_TOKEN_FILE", "bling_tokens.json")
SITE_URL = os.environ.get(
    "AGM_SITE_URL", "https://atlas-agm.comprasagmautopartes.workers.dev"
).rstrip("/")
SYNC_KEY = os.environ.get("AGM_SITE_SYNC_KEY", "").strip()
CF_ACCESS_CLIENT_ID = os.environ.get("CF_ACCESS_CLIENT_ID", "").strip()
CF_ACCESS_CLIENT_SECRET = os.environ.get("CF_ACCESS_CLIENT_SECRET", "").strip()


def main():
    if not TOKEN_FILE.exists():
        sys.exit(
            f"Nao encontrei {TOKEN_FILE.name} nesta pasta. Rode primeiro "
            f"'python bling_auth.py' pra autorizar o app no Bling."
        )
    if len(SYNC_KEY) < 32:
        sys.exit("AGM_SITE_SYNC_KEY ausente/curta no .env desta pasta.")
    if not CF_ACCESS_CLIENT_ID or not CF_ACCESS_CLIENT_SECRET:
        sys.exit("Credenciais do Cloudflare Access ausentes no .env desta pasta.")

    tokens = json.loads(TOKEN_FILE.read_text(encoding="utf-8"))
    headers = {
        "x-agm-sync-key": SYNC_KEY,
        "CF-Access-Client-Id": CF_ACCESS_CLIENT_ID,
        "CF-Access-Client-Secret": CF_ACCESS_CLIENT_SECRET,
        "Content-Type": "application/json",
    }
    resp = requests.post(SITE_URL + "/api/bling-token", headers=headers, json=tokens, timeout=20)
    resp.raise_for_status()
    print("OK - token enviado pro armazenamento em nuvem do Atlas.")
    print("A partir de agora, o GitHub Actions consegue renovar esse token sozinho.")


if __name__ == "__main__":
    main()
