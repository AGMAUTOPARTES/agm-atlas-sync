"""
Orquestra a sincronizacao completa (Bling -> inteligencia de compras ->
publicacao no Atlas) dentro de um runner do GitHub Actions, reportando
progresso pro MESMO endpoint que o agente local (atlas_sync_agent.ps1) ja
usa no Mac/Windows -- por isso a barra de progresso e o resto da tela do
Atlas continuam funcionando identico, sem precisar mudar nada no frontend.

Variaveis de ambiente esperadas (definidas como Secrets do repositorio no
GitHub -- ver README_CONFIGURACAO.md):
  BLING_CLIENT_ID, BLING_CLIENT_SECRET, BLING_BASE_URL
  AGM_SITE_SYNC_KEY, CF_ACCESS_CLIENT_ID, CF_ACCESS_CLIENT_SECRET
  AGM_SITE_URL          (opcional, ja tem um padrao)

Definidas pelo proprio workflow (nao sao segredo):
  TOKEN_STORE=cloud     (faz o bling_auth.py usar o D1 em vez de arquivo local)
  JOB_ID                (id do job criado pelo Atlas; vazio = teste manual,
                          roda a sincronizacao mas nao reporta heartbeat)
"""
import os
import subprocess
import sys

import requests

SITE_URL = os.environ.get(
    "AGM_SITE_URL", "https://atlas-agm.comprasagmautopartes.workers.dev"
).rstrip("/")
SYNC_KEY = os.environ.get("AGM_SITE_SYNC_KEY", "").strip()
CF_ID = os.environ.get("CF_ACCESS_CLIENT_ID", "").strip()
CF_SECRET = os.environ.get("CF_ACCESS_CLIENT_SECRET", "").strip()
JOB_ID = os.environ.get("JOB_ID", "").strip()

# Mesma lista de modulos "rapidos" que o refresh_bling.ps1 usa no clique
# padrao do botao -- cadastros auxiliares e financeiro ficam de fora pra
# nao gastar requisicoes sem necessidade.
FAST_MODULES = "categorias,situacoes,produtos,estoques_depositos,contatos,pedidos_venda,pedidos_compra,notas_entrada"

HEADERS = {
    "x-agm-sync-key": SYNC_KEY,
    "CF-Access-Client-Id": CF_ID,
    "CF-Access-Client-Secret": CF_SECRET,
    "Content-Type": "application/json",
}


def report(status, phase, progress, message):
    print(f"[{phase}] {progress}% - {message}")
    if not JOB_ID:
        return
    try:
        requests.post(
            SITE_URL + "/api/refresh/agent",
            headers=HEADERS,
            json={
                "id": JOB_ID,
                "status": status,
                "phase": phase,
                "progress": progress,
                "message": message,
            },
            timeout=20,
        )
    except requests.RequestException as exc:
        print(f"  aviso: falha ao reportar progresso pro Atlas: {exc}")


def run_step(args, phase, progress, message):
    report("running", phase, progress, message)
    result = subprocess.run([sys.executable] + args, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"{args[0]} terminou com codigo {result.returncode}")


def main():
    if len(SYNC_KEY) < 32 or not CF_ID or not CF_SECRET:
        sys.exit(
            "Faltam AGM_SITE_SYNC_KEY / CF_ACCESS_CLIENT_ID / "
            "CF_ACCESS_CLIENT_SECRET nos Secrets do repositorio no GitHub."
        )
    try:
        run_step(
            ["sync_bling.py", "--incremental", "--modulos", FAST_MODULES],
            "bling", 10, "Sincronizando dados do Bling",
        )
        run_step(
            ["gerar_sugestao_compras.py"], "calculations", 82,
            "Dados do Bling recebidos; recalculando a inteligencia de compras",
        )
        run_step(
            ["publicar_agm_site.py"], "publishing", 92,
            "Inteligencia recalculada; publicando a nova base no ATLAS",
        )
        report("complete", "complete", 100, "Nova base publicada no ATLAS")
        print("\nATLAS ATUALIZADO COM SUCESSO.")
    except Exception as exc:
        report("failed", "failed", 0, f"Falha na sincronizacao: {exc}")
        print(f"\nERRO: O ATLAS NAO FOI ATUALIZADO. {exc}")
        raise


if __name__ == "__main__":
    main()
