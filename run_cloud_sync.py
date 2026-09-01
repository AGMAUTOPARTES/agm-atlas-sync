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
import datetime as dt
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

# IMPORTANTE (01/09/2026): por que nao existe mais uma unica janela --desde
# global pra todo mundo, como tinha ate hoje de manha.
#
# Cada job do GitHub Actions comeca numa maquina ZERADA -- nao existe CSV de
# uma execucao anterior. O sync_bling.py tem um modo "upsert" (usado no
# --incremental) que MESCLA o que mudou com o CSV que ja existe no disco;
# no agente local isso funciona porque o CSV vai acumulando de sync em sync.
# Na nuvem nao ha o que mesclar -- "CSV existente (vazio) + o que mudou nos
# ultimos N dias" vira, na pratica, SO "o que mudou nos ultimos N dias". E
# como a publicacao no Atlas e um snapshot completo (substitui a base
# inteira), qualquer produto/pedido que nao mudou dentro da janela some do
# Atlas -- nao fica desatualizado, DESAPARECE. Foi o que aconteceu com a
# janela de 7 dias: so vieram os ~1000 produtos alterados na semana, e o
# resto do catalogo foi apagado da base publicada.
#
# A correcao: os modulos abaixo rodam em modo --full (sobrescreve, nunca
# mescla) em vez de --incremental (mescla/upsert), agrupados por
# necessidade real:
#
#   CADASTRO_MODULES -- produtos e contatos. Jonas quer sempre o cadastro
#   INTEIRO do Bling, entao rodam --full SEM --desde (sem limite de data
#   nenhum). estoques_depositos e categorias/situacoes tambem sempre foram
#   completos (nao tem filtro de data).
#
#   METRICAS_MODULES -- pedidos_venda e notas_entrada. Afetam faturamento,
#   dias sem vender, etc no Atlas (que so usa ate 90 dias de janela nos
#   calculos). Rodam --full com --desde de METRICAS_LOOKBACK_DAYS (120 dias,
#   30 de folga sobre o maior calculo do Atlas) em vez de sem limite
#   nenhum, senao seria puxar o historico inteiro de vendas da empresa a
#   cada execucao -- correto, mas provavelmente muito mais lento do que
#   precisa ser pra alimentar metricas de 90 dias.
#
#   pedidos_compra -- Jonas nao usa mais esse modulo no Bling (poucos
#   registros, so teste); vai passar a montar pedido de compra direto no
#   Atlas. Continua rodando --incremental sem --desde (cai no padrao interno
#   de 90 dias do sync_bling.py), exatamente como rodava antes de hoje.
CADASTRO_MODULES = "categorias,situacoes,produtos,estoques_depositos,contatos"
METRICAS_MODULES = "pedidos_venda,notas_entrada"
METRICAS_LOOKBACK_DAYS = 120
PEDIDOS_COMPRA_MODULES = "pedidos_compra"

HEADERS = {
    "x-agm-sync-key": SYNC_KEY,
    "CF-Access-Client-Id": CF_ID,
    "CF-Access-Client-Secret": CF_SECRET,
    "Content-Type": "application/json",
}


def claim():
    """Reivindica o job no Atlas (GET, o mesmo que o agente local faz a
    cada 30s) -- e o unico jeito de mover o status do job de 'requested'
    para 'running' no D1. Sem isso, os POSTs de progresso do report()
    abaixo falham calados (o endpoint so aceita atualizar um job que ja
    esta 'running'), o job fica preso em 'requested', e o watchdog do
    /api/refresh acaba marcando como 'failed' mesmo com a sincronizacao
    rodando normalmente aqui no GitHub Actions."""
    if not JOB_ID:
        return
    try:
        requests.get(SITE_URL + "/api/refresh/agent", headers=HEADERS, timeout=20)
    except requests.RequestException as exc:
        print(f"  aviso: falha ao reivindicar job no Atlas: {exc}")


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
        claim()

        run_step(
            ["sync_bling.py", "--full", "--modulos", CADASTRO_MODULES],
            "bling", 5, "Sincronizando catalogo completo (produtos e contatos)",
        )

        desde_metricas = (
            dt.date.today() - dt.timedelta(days=METRICAS_LOOKBACK_DAYS)
        ).strftime("%Y-%m-%d")
        run_step(
            ["sync_bling.py", "--full", "--modulos", METRICAS_MODULES, "--desde", desde_metricas],
            "bling", 45, "Sincronizando vendas e notas de entrada (120 dias)",
        )

        run_step(
            ["sync_bling.py", "--incremental", "--modulos", PEDIDOS_COMPRA_MODULES],
            "bling", 70, "Sincronizando pedidos de compra",
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
