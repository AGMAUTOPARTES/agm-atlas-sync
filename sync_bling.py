"""
AGM Auto Partes - Sync Bling API v3 -> CSV (para a Planilha Mae v4)
======================================================================
Le os dados do Bling via API v3 e grava um CSV por modulo na pasta ./data.
O Power Query da Planilha Mae le esses CSVs e carrega no Data Model.

IMPORTANTE - REGRA DE OURO (pedido do Jonas):
  Todo SKU/produto e relacionado pelo par (id, codigo) do Bling, nunca por
  posicao de linha. O "codigo" e o SKU. O "id" e o identificador interno do
  Bling (inteiro grande, ex: 16671229745) - gravamos como int Python puro,
  sem casting para 32 bits, exatamente o bug que corrompia a v3 anterior
  (toda linha tinha id=2147483647 = estouro de int32 no driver ODBC).

Uso:
    python sync_bling.py --full
        Carga completa, historico inteiro (pode demorar horas se voce tem
        anos de pedidos - veja --desde para carregar por partes).

    python sync_bling.py --full --desde 2025-01-01
        Carga completa, mas so a partir dessa data. Ideal pra deixar a
        planilha utilizavel rapido; depois roda o restante do historico
        separado (ex: de madrugada), sem pressa:
        python sync_bling.py --full --desde 2020-01-01 --ate 2024-12-31 --anexar --modulos pedidos_venda,pedidos_compra,contas_pagar,contas_receber,notas_fiscais,notas_entrada

    python sync_bling.py --incremental
        So o que mudou desde a ultima sync (uso diario).

    python sync_bling.py --full --modulos produtos,contatos
        Roda so os modulos listados.

Saida:
    ./data/<modulo>.csv          (um CSV por tabela da Planilha Mae)
    ./data/LOG_SYNC.csv          (historico de cada execucao, todo modulo)
    ./sync_state.json            (timestamps da ultima sync incremental)
"""

import argparse
import csv
import datetime as dt
import json
import os
import sys
import time
from pathlib import Path

from bling_client import BlingClient, BlingApiError

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATE_FILE = BASE_DIR / "sync_state.json"
LOG_FILE = DATA_DIR / "LOG_SYNC.csv"
STATUS_FILE = DATA_DIR / "SYNC_STATUS.json"

DATA_DIR.mkdir(exist_ok=True)

QUICK_LOOKBACK_DAYS = 7
RECONCILE_LOOKBACK_DAYS = 120


def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    return {}


def save_state(state):
    temp = STATE_FILE.with_suffix(".json.tmp")
    temp.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")
    temp.replace(STATE_FILE)


def write_csv(name, fieldnames, rows, append=False):
    path = DATA_DIR / (name + ".csv")
    final_rows = list(rows)
    if append and path.exists():
        with open(path, encoding="utf-8-sig", newline="") as current:
            final_rows = list(csv.DictReader(current)) + final_rows
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with open(temp_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        for row in final_rows:
            w.writerow(row)
        f.flush()
        os.fsync(f.fileno())
    temp_path.replace(path)
    return path


def write_csv_upsert(name, fieldnames, rows, key_field, allowed_keys=None):
    """Mescla 'rows' no CSV existente por 'key_field' (ex: produto_id):
      - registro cuja chave ja existe -> atualiza os valores (sem duplicar)
      - registro com chave nova -> adiciona
      - registro que nao veio em 'rows' (nao mudou desde a ultima sync) ->
        permanece intacto, como estava
    Usado nos modulos onde o Bling permite filtrar 'o que mudou'
    (dataAlteracaoInicial/Final): produtos, contatos, pedidos_venda. Isso
    evita ter que rebaixar o dataset inteiro a cada ciclo incremental."""
    path = DATA_DIR / (name + ".csv")
    existing = {}
    if path.exists():
        with open(path, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                k = row.get(key_field)
                if k not in (None, ""):
                    existing[str(k)] = row
    for row in rows:
        k = row.get(key_field)
        if k in (None, ""):
            continue
        existing[str(k)] = {fn: row.get(fn) for fn in fieldnames}
    if allowed_keys is not None:
        allowed = {str(key) for key in allowed_keys if key not in (None, "")}
        existing = {key: row for key, row in existing.items() if key in allowed}
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with open(temp_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        for row in existing.values():
            w.writerow(row)
        f.flush()
        os.fsync(f.fileno())
    temp_path.replace(path)
    return path


def write_csv_smart(name, fieldnames, rows, key_field=None, modo_execucao="full", append=False):
    """Full -> sobrescreve tudo (a carga ja trouxe o dataset inteiro, e isso
    tambem remove da planilha o que foi excluido no Bling nesse meio tempo).
    Incremental com key_field -> upsert (so mexe no que mudou, nunca duplica,
    nunca apaga o que nao veio na resposta). Sem key_field -> comportamento
    antigo (overwrite ou append conforme o parametro append)."""
    if modo_execucao != "full" and key_field:
        return write_csv_upsert(name, fieldnames, rows, key_field)
    return write_csv(name, fieldnames, rows, append=append)


def append_log(modulo, modo, registros, status, observacao=""):
    is_new = not LOG_FILE.exists()
    with open(LOG_FILE, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["data_hora", "modulo", "modo", "registros", "status", "observacao"])
        if is_new:
            w.writeheader()
        w.writerow({
            "data_hora": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "modulo": modulo,
            "modo": modo,
            "registros": registros,
            "status": status,
            "observacao": observacao,
        })


RUN_STATUS = {"status": "idle", "mode": None, "startedAt": None, "updatedAt": None, "currentModule": None, "modules": {}}


def save_run_status():
    RUN_STATUS["updatedAt"] = dt.datetime.now().isoformat(timespec="seconds")
    temp = STATUS_FILE.with_suffix(".json.tmp")
    temp.write_text(json.dumps(RUN_STATUS, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(STATUS_FILE)


def now_iso():
    return dt.datetime.now().strftime("%Y-%m-%d")


def run_module(modulo, modo, fn, *args, **kwargs):
    print("")
    print(">> " + modulo + " (" + modo + ") ...")
    t0 = time.time()
    RUN_STATUS["status"] = "running"
    RUN_STATUS["currentModule"] = modulo
    RUN_STATUS["modules"][modulo] = {"status": "running", "startedAt": dt.datetime.now().isoformat(timespec="seconds"), "records": 0, "durationSeconds": 0}
    save_run_status()
    try:
        count = fn(*args, **kwargs)
        dt_s = time.time() - t0
        print("   OK - " + str(count) + " registros em " + format(dt_s, ".1f") + "s")
        append_log(modulo, modo, count, "OK")
        RUN_STATUS["modules"][modulo].update({"status": "complete", "completedAt": dt.datetime.now().isoformat(timespec="seconds"), "records": count, "durationSeconds": round(dt_s, 1)})
        save_run_status()
        return count
    except Exception as exc:
        dt_s = time.time() - t0
        print("   ERRO - " + str(exc))
        append_log(modulo, modo, 0, "ERRO", str(exc)[:500])
        RUN_STATUS["status"] = "failed"
        RUN_STATUS["modules"][modulo].update({"status": "failed", "completedAt": dt.datetime.now().isoformat(timespec="seconds"), "durationSeconds": round(dt_s, 1), "message": str(exc)[:500]})
        save_run_status()
        raise


# ---------------------------------------------------------------------------
# modulos de apoio / cadastros (pequenos, sempre full)
# ---------------------------------------------------------------------------

def sync_categorias(client):
    lista = client.get_all("/categorias/produtos")
    rows = []
    for c in lista:
        rows.append({
            "categoria_id": c.get("id"),
            "descricao": c.get("descricao"),
            "categoriaPai_id": (c.get("categoriaPai") or {}).get("id"),
            "data_sync": now_iso(),
        })
    write_csv("categorias", ["categoria_id", "descricao", "categoriaPai_id", "data_sync"], rows)
    return len(rows)


def sync_depositos(client):
    lista = client.get_all("/depositos")
    rows = []
    for d in lista:
        rows.append({
            "deposito_id": d.get("id"),
            "descricao": d.get("descricao"),
            "situacao": d.get("situacao"),
            "padrao": d.get("padrao"),
            "data_sync": now_iso(),
        })
    write_csv("depositos", ["deposito_id", "descricao", "situacao", "padrao", "data_sync"], rows)
    return len(rows)


def sync_formas_pagamento(client):
    lista = client.get_all("/formas-pagamentos")
    rows = []
    for f in lista:
        rows.append({
            "forma_id": f.get("id"),
            "descricao": f.get("descricao"),
            "situacao": f.get("situacao"),
            "tipoPagamento": f.get("tipoPagamento"),
            "data_sync": now_iso(),
        })
    write_csv("formas_pagamento", ["forma_id", "descricao", "situacao", "tipoPagamento", "data_sync"], rows)
    return len(rows)


def sync_naturezas_operacoes(client):
    """/naturezas-operacoes - resolve o naturezaOperacao_id que aparece cru
    em PEDIDOS_VENDA, NOTAS_SAIDA e NOTAS_ENTRADA (pedido do Jonas: parar
    de mostrar so o ID onde da pra mostrar nome)."""
    lista = client.get_all("/naturezas-operacoes")
    rows = []
    for n in lista:
        rows.append({
            "natureza_id": n.get("id"),
            "descricao": n.get("descricao"),
            "tipo": n.get("tipo"),
            "padrao": n.get("padrao"),
            "data_sync": now_iso(),
        })
    write_csv("naturezas_operacoes", ["natureza_id", "descricao", "tipo", "padrao", "data_sync"], rows)
    return len(rows)


def sync_categorias_financeiras(client):
    """/categorias/receitas-despesas - resolve o categoria_id de
    CONTAS_PAGAR e CONTAS_RECEBER (categoria financeira, diferente da
    categoria de produto que ja sincronizamos em sync_categorias)."""
    lista = client.get_all("/categorias/receitas-despesas")
    rows = []
    for c in lista:
        rows.append({
            "categoria_fin_id": c.get("id"),
            "descricao": c.get("descricao"),
            "categoriaPai_id": (c.get("categoriaPai") or {}).get("id"),
            "tipo": c.get("tipo"),
            "data_sync": now_iso(),
        })
    write_csv("categorias_financeiras", ["categoria_fin_id", "descricao", "categoriaPai_id", "tipo", "data_sync"], rows)
    return len(rows)


def sync_vendedores(client):
    lista = client.get_all("/vendedores")
    rows = []
    for v in lista:
        contato = v.get("contato", {}) or {}
        rows.append({
            "vendedor_id": v.get("id"),
            "contato_id": contato.get("id"),
            "nome": contato.get("nome"),
            "situacao": v.get("situacao"),
            "data_sync": now_iso(),
        })
    write_csv("vendedores", ["vendedor_id", "contato_id", "nome", "situacao", "data_sync"], rows)
    return len(rows)


def sync_situacoes(client):
    data = client.get("/situacoes/modulos")
    modulos = (data or {}).get("data", [])
    rows = []
    for m in modulos:
        id_modulo = m.get("id")
        nome_modulo = m.get("nome")
        sub = client.get("/situacoes/modulos/" + str(id_modulo))
        for s in (sub or {}).get("data", []):
            rows.append({
                "situacao_id": s.get("id"),
                "modulo": nome_modulo,
                "idModuloSistema": id_modulo,
                "descricao": s.get("nome"),
                "idHerdado": s.get("idHerdado"),
                "cor": s.get("cor"),
            })
    write_csv("situacoes", ["situacao_id", "modulo", "idModuloSistema", "descricao", "idHerdado", "cor"], rows)
    return len(rows)


# ---------------------------------------------------------------------------
# produtos + estoque
# ---------------------------------------------------------------------------

PRODUTOS_FIELDS = [
    "produto_id", "codigo", "codigoAntigo", "nome", "tipo", "situacao", "formato",
    "preco", "precoCusto", "unidade", "pesoLiquido", "pesoBruto",
    "dim_largura", "dim_altura", "dim_profundidade", "dim_unidadeMedida",
    "gtin", "gtinEmbalagem", "ncm", "cest", "origem", "spedTipoItem",
    "marca", "descricaoCurta", "linkExterno",
    "estoque_saldoVirtual", "estoque_minimo", "estoque_maximo", "estoque_localizacao", "estoque_crossdocking",
    "categoria_id", "fornecedor_id", "fornecedor_nome", "fornecedor_codigo",
    "fornecedor_precoCusto", "fornecedor_precoCompra",
    "freteGratis", "condicao", "dataValidade",
    "itensPorCaixa", "unidadeCaixa",
    "dataInclusao", "dataAlteracao", "data_sync",
]


def sync_produtos(client, since=None, ate=None, modo_execucao="full"):
    params = {}
    if since:
        params["dataAlteracaoInicial"] = since
        params["dataAlteracaoFinal"] = ate or now_iso()
    lista = client.get_all("/produtos", params=params)
    # A API incremental informa inclusoes/alteracoes, mas nao devolve tombstones
    # de produtos excluidos. Uma listagem leve, sem baixar cada detalhe, vira a
    # fonte de verdade para remover IDs antigos do CSV e impedir acumulacao.
    current_ids = None
    if modo_execucao != "full":
        current_ids = {str(item.get("id")) for item in client.get_all("/produtos") if item.get("id") is not None}
    rows = []
    for p in lista:
        d = client.get_one("/produtos/" + str(p["id"]))
        if not d:
            raise RuntimeError("Produto " + str(p.get("id")) + " retornou detalhe vazio")
        est = d.get("estoque", {}) or {}
        dim = d.get("dimensoes", {}) or {}
        tri = d.get("tributacao", {}) or {}
        forn_raw = d.get("fornecedor") or {}
        if isinstance(forn_raw, list):
            forn = forn_raw[0] if forn_raw else {}
        else:
            forn = forn_raw
        fc = (forn or {}).get("contato", {}) or {}
        volumes = d.get("volumes") or {}
        if isinstance(volumes, list):
            volumes = volumes[0] if volumes else {}
        itens_por_caixa = d.get("itensPorCaixa")
        if itens_por_caixa is None and isinstance(volumes, dict):
            itens_por_caixa = volumes.get("itensPorCaixa") or volumes.get("qtdVolumes")
        unidade_caixa = d.get("unidadeCaixa")
        if unidade_caixa is None and isinstance(volumes, dict):
            unidade_caixa = volumes.get("unidadeCaixa") or volumes.get("tipoEmbalagem")
        rows.append({
            "produto_id": d.get("id"),
            "codigo": d.get("codigo"),
            "codigoAntigo": d.get("codigoAntigo") or d.get("codigoFornecedor"),
            "nome": d.get("nome"),
            "tipo": d.get("tipo"),
            "situacao": d.get("situacao"),
            "formato": d.get("formato"),
            "preco": d.get("preco"),
            "precoCusto": d.get("precoCusto") if d.get("precoCusto") is not None else p.get("precoCusto"),
            "unidade": d.get("unidade"),
            "pesoLiquido": d.get("pesoLiquido"),
            "pesoBruto": d.get("pesoBruto"),
            "dim_largura": dim.get("largura"),
            "dim_altura": dim.get("altura"),
            "dim_profundidade": dim.get("profundidade"),
            "dim_unidadeMedida": dim.get("unidadeMedida"),
            "gtin": d.get("gtin"),
            "gtinEmbalagem": d.get("gtinEmbalagem"),
            "ncm": tri.get("ncm"),
            "cest": tri.get("cest"),
            "origem": tri.get("origem"),
            "spedTipoItem": tri.get("spedTipoItem"),
            "marca": d.get("marca"),
            "descricaoCurta": d.get("descricaoCurta"),
            "linkExterno": d.get("linkExterno"),
            "estoque_saldoVirtual": est.get("saldoVirtualTotal"),
            "estoque_minimo": est.get("minimo"),
            "estoque_maximo": est.get("maximo"),
            "estoque_localizacao": est.get("localizacao"),
            "estoque_crossdocking": est.get("crossdocking"),
            "categoria_id": (d.get("categoria") or {}).get("id"),
            "fornecedor_id": fc.get("id"),
            "fornecedor_nome": fc.get("nome"),
            "fornecedor_codigo": (forn or {}).get("codigo"),
            "fornecedor_precoCusto": (forn or {}).get("precoCusto"),
            "fornecedor_precoCompra": (forn or {}).get("precoCompra"),
            "freteGratis": d.get("freteGratis"),
            "condicao": d.get("condicao"),
            "dataValidade": d.get("dataValidade"),
            "itensPorCaixa": itens_por_caixa,
            "unidadeCaixa": unidade_caixa,
            "dataInclusao": d.get("dataInclusao") or d.get("dataCriacao"),
            "dataAlteracao": d.get("dataAlteracao"),
            "data_sync": now_iso(),
        })
    if modo_execucao == "full":
        write_csv("produtos", PRODUTOS_FIELDS, rows)
    else:
        write_csv_upsert("produtos", PRODUTOS_FIELDS, rows, "produto_id", allowed_keys=current_ids)
    return len(rows)


def _querystring(pairs):
    from urllib.parse import urlencode
    return urlencode(pairs)


def sync_estoques_saldos(client, since=None, ate=None):
    produtos_csv = DATA_DIR / "produtos.csv"
    if not produtos_csv.exists():
        print("   (pulei: rode sync_produtos antes para ter a lista de IDs)")
        return 0
    with open(produtos_csv, encoding="utf-8-sig") as f:
        ids = [row["produto_id"] for row in csv.DictReader(f) if row.get("produto_id")]

    rows = []
    for i in range(0, len(ids), 100):
        bloco = ids[i:i + 100]
        params = [("idsProdutos[]", pid) for pid in bloco]
        data = client.get("/estoques/saldos?" + _querystring(params))
        for item in (data or {}).get("data", []):
            produto = item.get("produto") if isinstance(item.get("produto"), dict) else {}
            produto_id = produto.get("id") or item.get("produtoId") or item.get("id")
            for dep in item.get("depositos", []):
                rows.append({
                    "produto_id": produto_id,
                    "deposito_id": dep.get("id"),
                    "saldoFisico": dep.get("saldoFisicoTotal") if dep.get("saldoFisicoTotal") is not None else dep.get("saldoFisico"),
                    "saldoVirtual": dep.get("saldoVirtualTotal") if dep.get("saldoVirtualTotal") is not None else dep.get("saldoVirtual"),
                    "data_sync": now_iso(),
                })
    write_csv("estoques_depositos", ["produto_id", "deposito_id", "saldoFisico", "saldoVirtual", "data_sync"], rows)
    return len(rows)


# ---------------------------------------------------------------------------
# contatos
# ---------------------------------------------------------------------------

CONTATOS_FIELDS = [
    "contato_id", "nome", "fantasia", "codigo", "situacao", "tipo", "tiposContato",
    "numeroDocumento", "telefone", "celular", "email", "emailNotaFiscal",
    "ie", "rg", "inscricaoMunicipal", "indicadorIe",
    "end_endereco", "end_numero", "end_complemento", "end_bairro", "end_cep", "end_municipio", "end_uf",
    "cobr_endereco", "cobr_cep", "cobr_municipio", "cobr_uf",
    "financeiro_limiteCredito", "financeiro_condicaoPagamento", "financeiro_categoria_id",
    "vendedor_id", "dataNascimento", "dataInclusao", "dataAlteracao", "data_sync",
]


def sync_contatos(client, since=None, ate=None, modo_execucao="full"):
    params = {}
    if since:
        params["dataAlteracaoInicial"] = since
        params["dataAlteracaoFinal"] = ate or now_iso()
    lista = client.get_all("/contatos", params=params)
    rows = []
    for c in lista:
        d = client.get_one("/contatos/" + str(c["id"]))
        if not d:
            raise RuntimeError("Contato " + str(c.get("id")) + " retornou detalhe vazio")
        end = d.get("endereco", {}) or {}
        eg = end.get("geral", {}) or {}
        ec = end.get("cobranca", {}) or {}
        fin = d.get("financeiro", {}) or {}
        tipos = ", ".join(t.get("descricao", "") for t in (d.get("tiposContato") or []))
        rows.append({
            "contato_id": d.get("id"),
            "nome": d.get("nome"),
            "fantasia": d.get("fantasia"),
            "codigo": d.get("codigo"),
            "situacao": d.get("situacao"),
            "tipo": d.get("tipo"),
            "tiposContato": tipos,
            "numeroDocumento": d.get("numeroDocumento"),
            "telefone": d.get("telefone"),
            "celular": d.get("celular"),
            "email": d.get("email"),
            "emailNotaFiscal": d.get("emailNotaFiscal"),
            "ie": d.get("ie"),
            "rg": d.get("rg"),
            "inscricaoMunicipal": d.get("inscricaoMunicipal"),
            "indicadorIe": d.get("indicadorIe"),
            "end_endereco": eg.get("endereco"),
            "end_numero": eg.get("numero"),
            "end_complemento": eg.get("complemento"),
            "end_bairro": eg.get("bairro"),
            "end_cep": eg.get("cep"),
            "end_municipio": eg.get("municipio"),
            "end_uf": eg.get("uf"),
            "cobr_endereco": ec.get("endereco"),
            "cobr_cep": ec.get("cep"),
            "cobr_municipio": ec.get("municipio"),
            "cobr_uf": ec.get("uf"),
            "financeiro_limiteCredito": fin.get("limiteCredito"),
            "financeiro_condicaoPagamento": fin.get("condicaoPagamento"),
            "financeiro_categoria_id": (fin.get("categoria") or {}).get("id"),
            "vendedor_id": (d.get("vendedor") or {}).get("id"),
            "dataNascimento": (d.get("dadosAdicionais") or {}).get("dataNascimento"),
            "dataInclusao": d.get("dataInclusao"),
            "dataAlteracao": d.get("dataAlteracao"),
            "data_sync": now_iso(),
        })
    write_csv_smart("contatos", CONTATOS_FIELDS, rows, key_field="contato_id", modo_execucao=modo_execucao)
    return len(rows)


# ---------------------------------------------------------------------------
# pedidos de venda + itens + parcelas
# ---------------------------------------------------------------------------

PV_FIELDS = [
    "pedido_id", "numero", "numeroLoja", "data", "dataSaida", "dataPrevista",
    "totalProdutos", "total", "contato_id", "contato_nome", "contato_documento",
    "situacao_id", "situacao_valor", "loja_id", "vendedor_id",
    "desconto_valor", "desconto_unidade", "outrasDespesas",
    "frete_valor", "frete_porConta", "transportadora_id", "transportadora_nome",
    "naturezaOperacao_id", "numeroPedidoCompra", "notaFiscal_id",
    "observacoes", "observacoesInternas", "data_sync",
]
IV_FIELDS = [
    "item_id", "pedido_id", "pedido_numero", "pedido_data", "contato_id",
    "produto_id", "codigo", "descricao", "unidade", "quantidade", "valor",
    "desconto", "aliquotaIPI", "comissao_base", "comissao_aliquota", "comissao_valor", "data_sync",
]
PARC_FIELDS = [
    "parcela_id", "origem", "pedido_id", "pedido_numero", "contato_id",
    "dataVencimento", "valor", "formaPagamento_id", "observacoes", "data_sync",
]


def sync_pedidos_venda(client, since=None, ate=None, append=False, modo_execucao="full"):
    params = {}
    if since:
        params["dataAlteracaoInicial"] = since
        params["dataAlteracaoFinal"] = ate or now_iso()
    pedidos = client.get_all("/pedidos/vendas", params=params)
    pv_rows, iv_rows, parc_rows = [], [], []
    total = len(pedidos)
    for idx, p in enumerate(pedidos, 1):
        ct = p.get("contato", {}) or {}
        sit = p.get("situacao", {}) or {}
        d = client.get_one("/pedidos/vendas/" + str(p["id"]))
        if not d:
            raise RuntimeError("Pedido de venda " + str(p.get("id")) + " retornou detalhe vazio")
        desc = d.get("desconto", {}) or {}
        transp = d.get("transporte", {}) or {}
        tc = transp.get("contato", {}) or {}
        pv_rows.append({
            "pedido_id": p["id"], "numero": p.get("numero"), "numeroLoja": p.get("numeroLoja"),
            "data": p.get("data"), "dataSaida": p.get("dataSaida"), "dataPrevista": p.get("dataPrevista"),
            "totalProdutos": p.get("totalProdutos"), "total": p.get("total"),
            "contato_id": ct.get("id"), "contato_nome": ct.get("nome"), "contato_documento": ct.get("numeroDocumento"),
            "situacao_id": sit.get("id"), "situacao_valor": sit.get("valor"),
            "loja_id": p.get("loja", {}).get("id") if p.get("loja") else None,
            "vendedor_id": (d.get("vendedor") or {}).get("id"),
            "desconto_valor": desc.get("valor"), "desconto_unidade": desc.get("unidade"),
            "outrasDespesas": d.get("outrasDespesas"),
            "frete_valor": transp.get("frete"), "frete_porConta": transp.get("fretePorConta"),
            "transportadora_id": tc.get("id"), "transportadora_nome": tc.get("nome"),
            "naturezaOperacao_id": (d.get("naturezaOperacao") or {}).get("id"),
            "numeroPedidoCompra": d.get("numeroPedidoCompra"),
            "notaFiscal_id": (d.get("notaFiscal") or {}).get("id"),
            "observacoes": d.get("observacoes"), "observacoesInternas": d.get("observacoesInternas"),
            "data_sync": now_iso(),
        })
        for it in d.get("itens", []):
            com = it.get("comissao", {}) or {}
            iv_rows.append({
                "item_id": it.get("id"), "pedido_id": p["id"], "pedido_numero": p.get("numero"), "pedido_data": p.get("data"),
                "contato_id": ct.get("id"), "produto_id": (it.get("produto") or {}).get("id"),
                "codigo": it.get("codigo"), "descricao": it.get("descricao"), "unidade": it.get("unidade"),
                "quantidade": it.get("quantidade"), "valor": it.get("valor"), "desconto": it.get("desconto"),
                "aliquotaIPI": it.get("aliquotaIPI"), "comissao_base": com.get("base"),
                "comissao_aliquota": com.get("aliquota"), "comissao_valor": com.get("valor"),
                "data_sync": now_iso(),
            })
        for pa in d.get("parcelas", []):
            parc_rows.append({
                "parcela_id": pa.get("id"), "origem": "venda", "pedido_id": p["id"], "pedido_numero": p.get("numero"),
                "contato_id": ct.get("id"), "dataVencimento": pa.get("dataVencimento"), "valor": pa.get("valor"),
                "formaPagamento_id": (pa.get("formaPagamento") or {}).get("id"),
                "observacoes": pa.get("observacoes"), "data_sync": now_iso(),
            })
        if idx % 100 == 0:
            print("   ... " + str(idx) + "/" + str(total) + " pedidos de venda processados")
    write_csv_smart("pedidos_venda", PV_FIELDS, pv_rows, key_field="pedido_id", modo_execucao=modo_execucao, append=append)
    write_csv_smart("itens_venda", IV_FIELDS, iv_rows, key_field="item_id", modo_execucao=modo_execucao, append=append)
    write_csv_smart("parcelas_venda", PARC_FIELDS, parc_rows, key_field="parcela_id", modo_execucao=modo_execucao, append=append)
    return len(pv_rows)


# ---------------------------------------------------------------------------
# pedidos de compra + itens
# ---------------------------------------------------------------------------

PC_FIELDS = [
    "pedido_id", "numero", "data", "dataPrevista", "totalProdutos", "total",
    "fornecedor_id", "fornecedor_nome", "situacao_id", "situacao_valor",
    "observacoes", "notaEntrada_id", "data_sync",
]
IC_FIELDS = [
    "item_id", "pedido_id", "pedido_numero", "pedido_data", "fornecedor_id",
    "produto_id", "codigo", "descricao", "unidade", "quantidade", "valor",
    "desconto", "aliquotaIPI", "data_sync",
]


def sync_pedidos_compra(client, since=None, ate=None, append=False, modo_execucao="full"):
    params = {}
    if since:
        params["dataInicial"] = since
        params["dataFinal"] = ate or now_iso()
    pedidos = client.get_all("/pedidos/compras", params=params)
    pc_rows, ic_rows = [], []
    for p in pedidos:
        forn = p.get("fornecedor", {}) or {}
        sit = p.get("situacao", {}) or {}
        d = client.get_one("/pedidos/compras/" + str(p["id"]))
        if not d:
            raise RuntimeError("Pedido de compra " + str(p.get("id")) + " retornou detalhe vazio")
        pc_rows.append({
            "pedido_id": p["id"], "numero": p.get("numero"), "data": p.get("data"), "dataPrevista": p.get("dataPrevista"),
            "totalProdutos": p.get("totalProdutos"), "total": p.get("total"),
            "fornecedor_id": forn.get("id"), "fornecedor_nome": forn.get("nome"),
            "situacao_id": sit.get("id"), "situacao_valor": sit.get("valor"),
            "observacoes": d.get("observacoes"), "notaEntrada_id": (d.get("notaFiscal") or {}).get("id"),
            "data_sync": now_iso(),
        })
        for it in d.get("itens", []):
            ic_rows.append({
                "item_id": it.get("id"), "pedido_id": p["id"], "pedido_numero": p.get("numero"), "pedido_data": p.get("data"),
                "fornecedor_id": forn.get("id"), "produto_id": (it.get("produto") or {}).get("id"),
                "codigo": it.get("codigo"), "descricao": it.get("descricao"), "unidade": it.get("unidade"),
                "quantidade": it.get("quantidade"), "valor": it.get("valor"), "desconto": it.get("desconto"),
                "aliquotaIPI": it.get("aliquotaIPI"), "data_sync": now_iso(),
            })
    write_csv_smart("pedidos_compra", PC_FIELDS, pc_rows, key_field="pedido_id", modo_execucao=modo_execucao, append=append)
    write_csv_smart("itens_compra", IC_FIELDS, ic_rows, key_field="item_id", modo_execucao=modo_execucao, append=append)
    return len(pc_rows)


# ---------------------------------------------------------------------------
# financeiro
# ---------------------------------------------------------------------------

CP_FIELDS = [
    "titulo_id", "contato_id", "contato_nome", "historico", "numeroDocumento",
    "dataEmissao", "vencimento", "vencimentoOriginal", "dataPagamento",
    "valor", "valorPago", "saldo", "juros", "desconto", "situacao",
    "categoria_id", "portador_id", "numeroBanco", "ocorrencia_tipo", "data_sync",
]
CR_FIELDS = [
    "titulo_id", "contato_id", "contato_nome", "contato_documento", "historico", "numeroDocumento",
    "dataEmissao", "vencimento", "dataRecebimento", "valor", "valorRecebido", "saldo",
    "juros", "desconto", "acrescimo", "situacao", "categoria_id",
    "contaContabil_id", "contaContabil_descricao", "portador_id",
    "formaPagamento_id", "vendedor_id", "origem_tipo", "origem_id", "origem_numero",
    "linkBoleto", "linkQRCodePix", "data_sync",
]


def sync_contas_pagar(client, since=None, ate=None, append=False, modo_execucao="full"):
    params = {}
    if since:
        params["dataEmissaoInicial"] = since
        params["dataEmissaoFinal"] = ate or now_iso()
    lista = client.get_all("/contas/pagar", params=params)
    rows = []
    for c in lista:
        d = client.get_one("/contas/pagar/" + str(c["id"])) or c
        ct = d.get("contato", {}) or {}
        oc = d.get("ocorrencia", {}) or {}
        rows.append({
            "titulo_id": d.get("id", c.get("id")), "contato_id": ct.get("id"), "contato_nome": ct.get("nome"),
            "historico": d.get("historico"), "numeroDocumento": d.get("numeroDocumento"),
            "dataEmissao": d.get("dataEmissao"), "vencimento": d.get("vencimento"),
            "vencimentoOriginal": d.get("vencimentoOriginal"), "dataPagamento": d.get("dataPagamento"),
            "valor": d.get("valor"), "valorPago": d.get("valorPago"),
            "saldo": d.get("saldo"), "juros": d.get("juros"), "desconto": d.get("desconto"),
            "situacao": d.get("situacao"), "categoria_id": (d.get("categoria") or {}).get("id"),
            "portador_id": (d.get("portador") or {}).get("id"), "numeroBanco": d.get("numeroBanco"),
            "ocorrencia_tipo": oc.get("tipo"), "data_sync": now_iso(),
        })
    write_csv_smart("contas_pagar", CP_FIELDS, rows, key_field="titulo_id", modo_execucao=modo_execucao, append=append)
    return len(rows)


def sync_contas_receber(client, since=None, ate=None, append=False, modo_execucao="full"):
    params = {"tipoFiltroData": "E"}
    if since:
        params["dataInicial"] = since
        params["dataFinal"] = ate or now_iso()
    lista = client.get_all("/contas/receber", params=params)
    rows = []
    for c in lista:
        d = client.get_one("/contas/receber/" + str(c["id"])) or c
        ct = d.get("contato", {}) or {}
        cc = d.get("contaContabil", {}) or {}
        origem = d.get("origem", {}) or {}
        rows.append({
            "titulo_id": d.get("id", c.get("id")), "contato_id": ct.get("id"), "contato_nome": ct.get("nome"),
            "contato_documento": ct.get("numeroDocumento"),
            "historico": d.get("historico"), "numeroDocumento": d.get("numeroDocumento"),
            "dataEmissao": d.get("dataEmissao"), "vencimento": d.get("vencimento"),
            "dataRecebimento": d.get("dataRecebimento"), "valor": d.get("valor"), "valorRecebido": d.get("valorRecebido"),
            "saldo": d.get("saldo"), "juros": d.get("juros"), "desconto": d.get("desconto"), "acrescimo": d.get("acrescimo"),
            "situacao": d.get("situacao"), "categoria_id": (d.get("categoria") or {}).get("id"),
            "contaContabil_id": cc.get("id"), "contaContabil_descricao": cc.get("descricao"),
            "portador_id": (d.get("portador") or {}).get("id"),
            "formaPagamento_id": (d.get("formaPagamento") or {}).get("id"),
            "vendedor_id": (d.get("vendedor") or {}).get("id"),
            "origem_tipo": origem.get("tipoOrigem"), "origem_id": origem.get("id"), "origem_numero": origem.get("numero"),
            "linkBoleto": d.get("linkBoleto"), "linkQRCodePix": d.get("linkQRCodePix"),
            "data_sync": now_iso(),
        })
    write_csv_smart("contas_receber", CR_FIELDS, rows, key_field="titulo_id", modo_execucao=modo_execucao, append=append)
    return len(rows)


# ---------------------------------------------------------------------------
# notas fiscais
# ---------------------------------------------------------------------------

NFE_FIELDS = [
    "nfe_id", "numero", "serie", "tipo", "situacao", "chaveAcesso",
    "dataEmissao", "dataOperacao", "contato_id", "contato_nome", "contato_documento",
    "naturezaOperacao_id", "loja_id", "totalProdutos", "totalNF",
    "totalICMS", "totalIPI", "totalPIS", "totalCOFINS", "frete",
    "pedidoVenda_id", "data_sync",
]
ITENS_NFE_FIELDS = [
    "item_id", "nfe_id", "nfe_numero", "produto_id", "codigo", "descricao",
    "ncm", "cfop", "unidade", "quantidade", "valor", "valorTotal", "data_sync",
]


def sync_nfe(client, since=None, ate=None, append=False, tipo="1", modo_execucao="full"):
    """tipo da NF-e no Bling: 0 = Entrada, 1 = Saida (default da API se omitido).
    Gravamos em arquivos separados pra cada tipo, pra nao misturar compra com venda."""
    params = {"tipo": tipo}
    if since:
        params["dataEmissaoInicial"] = since
        params["dataEmissaoFinal"] = ate or now_iso()
    lista = client.get_all("/nfe", params=params)
    nfe_rows, item_rows = [], []
    for n in lista:
        d = client.get_one("/nfe/" + str(n["id"]))
        if not d:
            raise RuntimeError("NF-e " + str(n.get("id")) + " retornou detalhe vazio")
        ct = d.get("contato", {}) or {}
        tot = d.get("totalNota", {}) if isinstance(d.get("totalNota"), dict) else {}
        nfe_rows.append({
            "nfe_id": d.get("id"), "numero": d.get("numero"), "serie": d.get("serie"),
            "tipo": d.get("tipo"), "situacao": d.get("situacao"), "chaveAcesso": d.get("chaveAcesso"),
            "dataEmissao": d.get("dataEmissao"), "dataOperacao": d.get("dataOperacao"),
            "contato_id": ct.get("id"), "contato_nome": ct.get("nome"), "contato_documento": ct.get("numeroDocumento"),
            "naturezaOperacao_id": (d.get("naturezaOperacao") or {}).get("id"),
            "loja_id": (d.get("loja") or {}).get("id"),
            "totalProdutos": d.get("valorTotalProdutos") or tot.get("valorProdutos"),
            "totalNF": d.get("valorNota") or tot.get("valorNota"),
            "totalICMS": tot.get("icms"), "totalIPI": tot.get("ipi"),
            "totalPIS": tot.get("pis"), "totalCOFINS": tot.get("cofins"),
            "frete": d.get("transporte", {}).get("frete") if d.get("transporte") else None,
            "pedidoVenda_id": (d.get("pedido") or {}).get("id") if d.get("pedido") else None,
            "data_sync": now_iso(),
        })
        for it in d.get("itens", []):
            item_rows.append({
                "item_id": it.get("id"), "nfe_id": d.get("id"), "nfe_numero": d.get("numero"),
                "produto_id": (it.get("produto") or {}).get("id"), "codigo": it.get("codigo"),
                "descricao": it.get("descricao"), "ncm": it.get("ncm"), "cfop": it.get("cfop"),
                "unidade": it.get("unidade"), "quantidade": it.get("quantidade"),
                "valor": it.get("valor"), "valorTotal": it.get("valorTotal"), "data_sync": now_iso(),
            })
    nome_base = "notas_fiscais" if str(tipo) == "1" else "notas_entrada"
    nome_itens = "itens_notas_fiscais" if str(tipo) == "1" else "itens_notas_entrada"
    # A API pode repetir a ultima linha de uma pagina na pagina seguinte.
    # Deduplica a nota por ID e o item por uma chave composta estável quando
    # o Bling não fornece item_id.
    nfe_rows = list({str(r.get("nfe_id")): r for r in nfe_rows if r.get("nfe_id")}.values())
    itens_unicos = {}
    for r in item_rows:
        chave = str(r.get("item_id") or "|").strip()
        if chave in ("", "|"):
            chave = "|".join(str(r.get(k) or "").strip() for k in ("nfe_id", "codigo", "descricao", "quantidade", "valor"))
            r["item_id"] = chave
        itens_unicos[chave] = r
    item_rows = list(itens_unicos.values())
    write_csv_smart(nome_base, NFE_FIELDS, nfe_rows, key_field="nfe_id", modo_execucao=modo_execucao, append=append)
    write_csv_smart(nome_itens, ITENS_NFE_FIELDS, item_rows, key_field="item_id", modo_execucao=modo_execucao, append=append)
    return len(nfe_rows)


def sync_nfe_saida(client, since=None, ate=None, append=False, modo_execucao="full"):
    return sync_nfe(client, since=since, ate=ate, append=append, tipo="1", modo_execucao=modo_execucao)


def sync_nfe_entrada(client, since=None, ate=None, append=False, modo_execucao="full"):
    return sync_nfe(client, since=since, ate=ate, append=append, tipo="0", modo_execucao=modo_execucao)


# ---------------------------------------------------------------------------
# orquestracao
# ---------------------------------------------------------------------------

MODULES = [
    ("categorias", sync_categorias, None),
    ("depositos", sync_depositos, None),
    ("formas_pagamento", sync_formas_pagamento, None),
    ("vendedores", sync_vendedores, None),
    ("situacoes", sync_situacoes, None),
    ("naturezas_operacoes", sync_naturezas_operacoes, None),
    ("categorias_financeiras", sync_categorias_financeiras, None),
    ("produtos", sync_produtos, "alteracao"),
    ("estoques_depositos", sync_estoques_saldos, None),
    ("contatos", sync_contatos, "alteracao"),
    ("pedidos_venda", sync_pedidos_venda, "alteracao"),
    ("pedidos_compra", sync_pedidos_compra, "janela"),
    ("contas_pagar", sync_contas_pagar, "janela"),
    ("contas_receber", sync_contas_receber, "janela"),
    ("notas_fiscais", sync_nfe_saida, "janela"),
    ("notas_entrada", sync_nfe_entrada, "janela"),
]


def compute_since(modo_filtro, modulo, state, modo_execucao, desde_cli):
    if desde_cli:
        return desde_cli
    if modo_execucao == "full" or modo_filtro is None:
        return None
    if modo_execucao == "reconcile":
        return (dt.date.today() - dt.timedelta(days=RECONCILE_LOOKBACK_DAYS)).strftime("%Y-%m-%d")
    if modo_filtro == "alteracao":
        anterior = state.get(modulo)
        if anterior:
            # Sobreposicao de 2 dias protege contra relogio, atraso da API e
            # um estado antigo que tenha avancado antes de uma falha.
            try:
                return (dt.date.fromisoformat(anterior) - dt.timedelta(days=2)).strftime("%Y-%m-%d")
            except ValueError:
                pass
        return anterior
    if modo_filtro == "janela":
        anterior = state.get(modulo)
        if anterior:
            try:
                return (dt.date.fromisoformat(anterior) - dt.timedelta(days=2)).strftime("%Y-%m-%d")
            except ValueError:
                pass
        return (dt.date.today() - dt.timedelta(days=QUICK_LOOKBACK_DAYS)).strftime("%Y-%m-%d")
    return None


def main():
    parser = argparse.ArgumentParser(description="Sincroniza dados do Bling API v3 para CSV (Planilha Mae AGM)")
    parser.add_argument("--full", action="store_true", help="carga completa (ignora estado incremental)")
    parser.add_argument("--incremental", action="store_true", help="so o que mudou desde a ultima sync")
    parser.add_argument("--reconcile", action="store_true", help="reconfere os ultimos 120 dias sem refazer todo o historico")
    parser.add_argument("--modulos", type=str, default=None, help="lista separada por virgula, ex: produtos,contatos")
    parser.add_argument("--desde", type=str, default=None, help="data inicial YYYY-MM-DD para limitar o historico")
    parser.add_argument("--ate", type=str, default=None, help="data final YYYY-MM-DD, usada junto com --desde")
    parser.add_argument("--anexar", action="store_true", help="acrescenta ao CSV existente em vez de sobrescrever")
    args = parser.parse_args()

    if sum(bool(value) for value in (args.full, args.incremental, args.reconcile)) != 1:
        print("Escolha exatamente um modo: --full, --incremental ou --reconcile.")
        sys.exit(1)

    modo_execucao = "full" if args.full else ("reconcile" if args.reconcile else "incremental")
    only = set(m.strip() for m in args.modulos.split(",")) if args.modulos else None

    client = BlingClient()
    state = load_state()
    started_at = dt.datetime.now()
    RUN_STATUS.update({"status": "running", "mode": modo_execucao, "startedAt": started_at.isoformat(timespec="seconds"), "updatedAt": started_at.isoformat(timespec="seconds"), "currentModule": None, "modules": {}})
    save_run_status()

    print("=== AGM - Sync Bling API v3 (" + modo_execucao + ") - " + started_at.strftime("%d/%m/%Y %H:%M") + " ===")
    if args.desde:
        ate_txt = args.ate if args.ate else "hoje"
        print("Filtrando desde " + args.desde + " ate " + ate_txt)

    for modulo, fn, filtro in MODULES:
        if only and modulo not in only:
            continue
        since = compute_since(filtro, modulo, state, modo_execucao, args.desde)
        kwargs = {}
        if filtro:
            kwargs["since"] = since
            kwargs["ate"] = args.ate
        if modulo in ("pedidos_venda", "pedidos_compra", "contas_pagar", "contas_receber", "notas_fiscais", "notas_entrada"):
            kwargs["append"] = args.anexar
        if filtro:
            kwargs["modo_execucao"] = modo_execucao
        run_module(modulo, modo_execucao, fn, client, **kwargs)
        if filtro and not args.desde:
            state[modulo] = dt.date.today().strftime("%Y-%m-%d")
            # Salva o estado logo apos CADA modulo (nao so no final). Assim,
            # se o processo for interrompido/cancelado no meio (janela
            # fechada, PC dormiu, etc.), o progresso ja feito nao se perde -
            # da proxima vez ele continua so do que falta, em vez de
            # reprocessar tudo de novo do zero.
            save_state(state)

    save_state(state)
    elapsed = (dt.datetime.now() - started_at).total_seconds()
    RUN_STATUS.update({"status": "complete", "currentModule": None, "completedAt": dt.datetime.now().isoformat(timespec="seconds"), "durationSeconds": round(elapsed, 1)})
    save_run_status()
    print("")
    print("=== Concluido em " + format(elapsed, ".0f") + "s | " + str(client.total_requests) +
          " requisicoes (" + str(client.total_429) + " bloqueadas por limite) | CSVs em " + str(DATA_DIR) + " ===")


if __name__ == "__main__":
    main()
