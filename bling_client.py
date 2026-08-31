"""
AGM Auto Partes - Bling API v3 - Cliente HTTP
================================================
Encapsula:
  - Autenticacao (Bearer token, renovado automaticamente via bling_auth).
  - Limite de requisicoes do Bling: 3 req/s e 120.000 req/dia, POR CONTA
    (nao por app) - se outro integrador (ex.: a sync antiga com Google
    Sheets) estiver rodando ao mesmo tempo, o limite e' compartilhado.
    Por isso o intervalo entre chamadas AQUI E' ADAPTATIVO: sobe quando
    leva 429 (throttle externo tambem conta) e desce devagar quando as
    coisas voltam a fluir.
  - Paginacao automatica (parametros "pagina" e "limite").
  - Retentativas com backoff exponencial em erros 429 / 5xx / rede.
"""

import time
import threading

import requests

import bling_auth

MIN_INTERVAL = 0.45   # piso ~2.2 req/s, com folga sob o limite de 3 req/s do Bling
MAX_INTERVAL = 8.0    # teto do intervalo adaptativo
MAX_RETRIES = 6
PAGE_SIZE = 100


class BlingApiError(Exception):
    def __init__(self, status_code, payload):
        self.status_code = status_code
        self.payload = payload
        msg = "Bling API erro " + str(status_code) + ": " + str(payload)
        super().__init__(msg)


class _AdaptiveRateLimiter:
    """Garante um intervalo minimo entre chamadas. O intervalo cresce
    automaticamente quando o Bling responde 429 (sinal de que o limite da
    CONTA foi atingido - por este script ou por qualquer outro integrador
    rodando ao mesmo tempo) e diminui aos poucos quando as chamadas voltam
    a ter sucesso, ate o piso MIN_INTERVAL."""

    def __init__(self, min_interval=MIN_INTERVAL, max_interval=MAX_INTERVAL):
        self.min_interval = min_interval
        self.max_interval = max_interval
        self.current = min_interval
        self._lock = threading.Lock()
        self._last_call = 0.0
        self._successes_since_bump = 0

    def wait(self):
        with self._lock:
            now = time.monotonic()
            elapsed = now - self._last_call
            if elapsed < self.current:
                time.sleep(self.current - elapsed)
            self._last_call = time.monotonic()

    def penalize(self):
        with self._lock:
            self.current = min(self.max_interval, max(self.current * 1.8, 1.0))
            self._successes_since_bump = 0

    def reward(self):
        with self._lock:
            self._successes_since_bump += 1
            if self._successes_since_bump >= 8 and self.current > self.min_interval:
                self.current = max(self.min_interval, self.current * 0.85)
                self._successes_since_bump = 0


class BlingClient:
    def __init__(self, base_url=None, min_interval=MIN_INTERVAL):
        self.base_url = (base_url or bling_auth.BASE_URL).rstrip("/")
        self.session = requests.Session()
        self._limiter = _AdaptiveRateLimiter(min_interval)
        self.total_requests = 0
        self.total_429 = 0

    def _headers(self):
        token = bling_auth.get_valid_access_token()
        return {
            "Authorization": "Bearer " + token,
            "Accept": "application/json",
        }

    def request(self, method, path, params=None, json_body=None):
        url = path if path.startswith("http") else self.base_url + path
        last_exc = None
        for attempt in range(1, MAX_RETRIES + 1):
            self._limiter.wait()
            try:
                resp = self.session.request(
                    method, url, headers=self._headers(), params=params,
                    json=json_body, timeout=30,
                )
                self.total_requests += 1
            except requests.RequestException as exc:
                last_exc = exc
                self._limiter.penalize()
                wait = min(2 ** attempt, 30)
                print("  [rede] erro (" + str(exc) + "); tentativa " + str(attempt) + "/" + str(MAX_RETRIES) + ", aguardando " + str(wait) + "s")
                time.sleep(wait)
                continue

            if resp.status_code == 429:
                self.total_429 += 1
                self._limiter.penalize()
                retry_after = resp.headers.get("Retry-After")
                wait = float(retry_after) if retry_after else min(2 ** attempt, 30)
                if self.total_429 % 20 == 0:
                    print("  [429] " + str(self.total_429) + "x ate agora - ritmo atual: " +
                          format(self._limiter.current, ".2f") + "s entre chamadas " +
                          "(pode ser outro integrador usando a mesma conta Bling)")
                time.sleep(wait)
                continue

            if resp.status_code in (500, 502, 503, 504):
                self._limiter.penalize()
                wait = min(2 ** attempt, 30)
                print("  [" + str(resp.status_code) + "] erro temporario do servidor; tentativa " + str(attempt) + "/" + str(MAX_RETRIES) + ", aguardando " + str(wait) + "s")
                time.sleep(wait)
                continue

            if resp.status_code >= 400:
                try:
                    payload = resp.json()
                except ValueError:
                    payload = resp.text
                raise BlingApiError(resp.status_code, payload)

            self._limiter.reward()
            if resp.status_code == 204 or not resp.content:
                return {}
            return resp.json()

        raise RuntimeError("Falha apos " + str(MAX_RETRIES) + " tentativas em " + url + ": " + str(last_exc))

    def get(self, path, params=None):
        return self.request("GET", path, params=params)

    def post(self, path, json_body=None, params=None):
        return self.request("POST", path, params=params, json_body=json_body)

    def get_all(self, path, params=None, page_size=PAGE_SIZE, max_pages=None):
        items = []
        page = 1
        while True:
            p = {"pagina": page, "limite": page_size}
            if params:
                p.update(params)
            # Nunca transforme erro de uma pagina em "carga completa". O
            # chamador precisa falhar para manter o ultimo CSV valido.
            data = self.get(path, params=p)
            page_items = (data or {}).get("data", [])
            if not page_items:
                break
            items.extend(page_items)
            if len(page_items) < page_size:
                break
            page += 1
            if max_pages and page > max_pages:
                break
        return items

    def get_one(self, path):
        # Um detalhe ausente no meio da carga torna o modulo incompleto.
        # Propagamos o erro em vez de pular silenciosamente o registro.
        data = self.get(path)
        return (data or {}).get("data", {})
