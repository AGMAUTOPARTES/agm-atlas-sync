"""
Armazenamento do token OAuth do Bling.

- Modo local (padrao, TOKEN_STORE ausente ou != "cloud"): le/grava um
  arquivo JSON (bling_tokens.json) na propria pasta, exatamente como o
  script sempre fez no Mac/Windows. Nada muda pra quem roda localmente.

- Modo nuvem (TOKEN_STORE=cloud, usado pelo GitHub Actions): le/grava o
  token via HTTP num endpoint protegido do proprio Worker do Atlas
  (/api/bling-token), que guarda o token no D1. Necessario porque um
  runner do GitHub Actions e descartado ao final de cada execucao -- nao
  tem disco persistente entre uma rodada e a proxima, entao o token
  renovado precisa morar em algum lugar que sobrevive entre execucoes.

Nenhuma credencial e' impressa ou logada por este modulo.
"""
import json
import os
from pathlib import Path

import requests

MODE = os.environ.get("TOKEN_STORE", "local").strip().lower()


def _cloud_headers():
    headers = {
        "x-agm-sync-key": os.environ.get("AGM_SITE_SYNC_KEY", ""),
        "Content-Type": "application/json",
    }
    cf_id = os.environ.get("CF_ACCESS_CLIENT_ID", "")
    cf_secret = os.environ.get("CF_ACCESS_CLIENT_SECRET", "")
    if cf_id and cf_secret:
        headers["CF-Access-Client-Id"] = cf_id
        headers["CF-Access-Client-Secret"] = cf_secret
    return headers


def _cloud_url():
    base = os.environ.get(
        "AGM_SITE_URL", "https://atlas-agm.comprasagmautopartes.workers.dev"
    ).rstrip("/")
    return base + "/api/bling-token"


def load_tokens(token_file: Path):
    if MODE == "cloud":
        resp = requests.get(_cloud_url(), headers=_cloud_headers(), timeout=20)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        data = resp.json()
        return data if data and data.get("access_token") else None
    if not token_file.exists():
        return None
    return json.loads(token_file.read_text(encoding="utf-8"))


def save_tokens(token_file: Path, payload: dict):
    if MODE == "cloud":
        resp = requests.post(_cloud_url(), headers=_cloud_headers(), json=payload, timeout=20)
        resp.raise_for_status()
        return payload
    token_file.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return payload
