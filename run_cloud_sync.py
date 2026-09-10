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
AGM_SITE_URL (opcional, ja tem um padrao)

Definidas pelo proprio workflow (nao sao segredo):
TOKEN_STORE=cloud (faz o bling_auth.py usar o D1 em vez de arquivo local)
JOB_ID (id do job criado pelo Atlas; vazio = teste manual,
roda a sincronizacao mas nao reporta heartbeat)

Novo em 10/09/2026 (a pedido do Jonas, depois de descobrir que o cron
automatico de 4 em 4 horas estava perto de estourar a cota gratuita do
GitHub Actions -- ver conversa no Atlas): o cron foi DESLIGADO. A partir
de agora toda sincronizacao e' sob demanda, disparada pelo botao do Atlas,
e o Jonas escolhe (opcionalmente) QUAIS modulos quer atualizar -- em vez
de sempre puxar o catalogo inteiro (que e' o que demora ~50min sozinho).

MODULES (opcional): lista de modulos separados por virgula, usando os
mesmos nomes do sync_bling.py:
  categorias, situacoes, produtos, estoques_depositos, contatos,
  pedidos_venda, notas_entrada, pedidos_compra
Vazio/nao definido = todos os modulos (comportamento antigo, sync completo).

PERIOD_START / PERIOD_END (opcional, "AAAA-MM-DD"): so afetam os modulos
com filtro de data (pedidos_venda, notas_entrada). Sem eles, cai no padrao
de METRICAS_LOOKBACK_DAYS dias pra tras, ate hoje.

IMPORTANTE -- por que cada modulo tem uma "classe de seguranca" fixa (full
sem data, full com janela, ou incremental) que o Jonas NAO escolhe: isso
foi decidido em 01/09/2026 depois de um bug real onde sincronizar
incremental num runner "zerado" (sem o CSV de execucoes anteriores) fazia
produto que nao mudou na janela SUMIR do Atlas em vez de so ficar
desatualizado. Deixar o modulo escolher a classe de novo reabriria esse
risco -- por isso a classe de cada modulo e' fixa aqui embaixo, e so a
LISTA de modulos e' que e' escolhida pelo Jonas.
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

# Classe de seguranca de cada modulo (nao mexer sem entender o comentario
# grande acima): "full_nodate" roda --full sem --desde (cadastro inteiro,
# sempre); "full_windowed" roda --full com --desde (snapshot completo, mas
# so dos ultimos N dias, pra nao puxar o historico inteiro de vendas toda
# vez); "incremental" roda --incremental (mescla com o que ja tem -- unico
# modulo onde isso e' seguro hoje, porque o Jonas praticamente nao usa mais
# esse cadastro no Bling).
MODULE_CLASS = {
    "categorias": "full_nodate",
    "situacoes": "full_nodate",
    "produtos": "full_nodate",
    "estoques_depositos": "full_nodate",
    "contatos": "full_nodate",
    "pedidos_venda": "full_windowed",
    "notas_entrada": "full_windowed",
    "pedidos_compra": "incremental",
}
ALL_MODULES = list(MODULE_CLASS.keys())
METRICAS_LOOKBACK_DAYS = 120

HEADERS = {
    "x-agm-sync-key": SYNC_KEY,
    "CF-Access-Client-Id": CF_ID,
    "CF-Access-Client-Secret": CF_SECRET,
    "Content-Type": "application/json",
}


def parse_modules():
    """Le MODULES do ambiente. Vazio ou ausente = todos os modulos (mesmo
    comportamento de sempre). Nomes desconhecidos sao ignorados (com aviso)
    em vez de quebrar a sincronizacao inteira por um erro de digitacao."""
    raw = os.environ.get("MODULES", "").strip()
    if not raw:
        return list(ALL_MODULES)
    requested, unknown = [], []
    for name in raw.split(","):
        name = name.strip()
        if not name:
            continue
        if name in MODULE_CLASS:
            if name not in requested:
                requested.append(name)
        else:
            unknown.append(name)
    if unknown:
        print(f" aviso: modulo(s) desconhecido(s) ignorado(s): {', '.join(unknown)}")
    return requested or list(ALL_MODULES)


def build_plan(requested_modules):
    """Agrupa os modulos pedidos por classe de seguranca, preservando a
    ordem de MODULE_CLASS, e devolve uma lista de passos
    (classe, [modulos]) so com as classes que tem pelo menos 1 modulo
    pedido."""
    by_class: dict[str, list[str]] = {}
    for module in ALL_MODULES:
        if module in requested_modules:
            by_class.setdefault(MODULE_CLASS[module], []).append(module)
    order = ["full_nodate", "full_windowed", "incremental"]
    return [(cls, by_class[cls]) for cls in order if by_class.get(cls)]


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
        print(f" aviso: falha ao reivindicar job no Atlas: {exc}")


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
        print(f" aviso: falha ao reportar progresso pro Atlas: {exc}")


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

    requested_modules = parse_modules()
    plan = build_plan(requested_modules)
    if not plan:
        sys.exit(f"Nenhum modulo valido em MODULES={os.environ.get('MODULES', '')!r}.")

    period_start = os.environ.get("PERIOD_START", "").strip()
    period_end = os.environ.get("PERIOD_END", "").strip()
    desde_padrao = (dt.date.today() - dt.timedelta(days=METRICAS_LOOKBACK_DAYS)).strftime("%Y-%m-%d")

    labels = {
        "full_nodate": "cadastro completo",
        "full_windowed": "vendas/notas de entrada",
        "incremental": "pedidos de compra",
    }
    print(f"=== AGM - Sync Bling API v3 (sob demanda) - {dt.datetime.now():%d/%m/%Y %H:%M} ===")
    print(f"Modulos pedidos: {', '.join(requested_modules)}")
    plan_descriptions = [f"{labels[cls]} ({','.join(mods)})" for cls, mods in plan]
    print(f"Plano: {' | '.join(plan_descriptions)}")

    try:
        claim()

        n_steps = len(plan)
        # 5% pra abrir, ate 90% distribuido entre os passos do Bling, resto
        # (90-100%) pros calculos + publicacao no final.
        progress_points = [5 + round(85 * (i + 1) / n_steps) for i in range(n_steps)]

        for (cls, modules), progress in zip(plan, progress_points):
            modulos_str = ",".join(modules)
            if cls == "full_nodate":
                run_step(
                    ["sync_bling.py", "--full", "--modulos", modulos_str],
                    "bling", progress, f"Sincronizando {labels[cls]} ({modulos_str})",
                )
            elif cls == "full_windowed":
                args = ["sync_bling.py", "--full", "--modulos", modulos_str,
                        "--desde", period_start or desde_padrao]
                if period_end:
                    args += ["--ate", period_end]
                run_step(
                    args, "bling", progress,
                    f"Sincronizando {labels[cls]} ({modulos_str}, desde {period_start or desde_padrao})",
                )
            else:  # incremental
                run_step(
                    ["sync_bling.py", "--incremental", "--modulos", modulos_str],
                    "bling", progress, f"Sincronizando {labels[cls]} ({modulos_str})",
                )

        run_step(
            ["gerar_sugestao_compras.py"], "calculations", 93,
            "Dados do Bling recebidos; recalculando a inteligencia de compras",
        )
        run_step(
            ["publicar_agm_site.py"], "publishing", 97,
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
