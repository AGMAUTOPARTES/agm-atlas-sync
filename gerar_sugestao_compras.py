#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_sugestao_compras.py
==========================
Gera a planilha Sugestao_Compras.xlsx a partir dos CSVs sincronizados do
Bling (pasta data/), com:

  - Um produto por linha, calculado a partir de vendas/estoque/compras reais.
  - Classificacao de saude de estoque (ruptura, risco, saudavel, excesso,
    sem giro, sem historico de venda).
  - Sugestao de quantidade e valor de compra por produto.
  - Custo em cascata: nota de entrada mais recente > pedido de compra mais
    recente > cadastro do produto (com a origem registrada na planilha).
  - Lead time medio por fornecedor (heuristica: pedido_compra -> notaEntrada
    vinculada via pedidos_compra.notaEntrada_id).
  - Selecao de itens (coluna "Comprar?" + "Qtd") com aba Carrinho que reune
    so os itens marcados.
  - Painel de KPIs nas linhas 1-7 da aba Sugestao_Compras (teto de compras,
    valor comprado no mes, valor de custo em estoque, produtos parados,
    produtos a comprar, lead time medio).

Dependencias: apenas biblioteca padrao (csv, datetime, pathlib) + openpyxl.
Sem pandas de proposito -- mesma pilha do sync_bling.py, que ja roda de
forma confiavel via Task Scheduler/.bat neste computador.

Saida: C:\\Users\\Jonas AGM\\Downloads\\JONAS\\Monalisa\\Sugestao_Compras.xlsx
"""

import csv
import datetime as dt
import hashlib
import json
from pathlib import Path
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.worksheet.table import Table, TableStyleInfo

# =====================================================================
# CAMINHOS E PARAMETROS
# =====================================================================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
# No pacote original (Windows) isso apontava pra uma pasta fixa do Jonas.
# Nesta versao portatil (Mac/nuvem), a planilha de cortesia vai pra dentro
# de data/ como qualquer outro artefato do pipeline -- o que realmente
# alimenta o Atlas e' o monalisa_snapshot.json, nao esta xlsx.
OUTPUT_PATH = DATA_DIR / "Sugestao_Compras.xlsx"
SNAPSHOT_PATH = DATA_DIR / "monalisa_snapshot.json"

HOJE = dt.date.today()

TETO_COMPRAS_MES = 90000.0          # teto mensal de compras (R$) - ajustavel na aba Parametros
SEM_MOVIMENTO_LIMIAR_DIAS = 90      # dias sem venda para considerar "sem giro"
COBERTURA_MIN_DIAS = 15             # cobertura abaixo disso = risco de ruptura
COBERTURA_MAX_DIAS = 180            # cobertura acima disso = excesso
JANELA_VELOCIDADE_DIAS = 90         # janela usada para calcular velocidade de venda
META_COBERTURA_REPOSICAO_DIAS = 60  # cobertura-alvo ao sugerir reposicao

HEADER_ROW = 8   # linha do cabecalho da tabela principal; linhas 1-7 = painel de KPIs

NF_ENTRADA_VALIDAS = {"5", "6", "7"}  # Autorizada, Emitida, Registrada (enum oficial Bling)

# =====================================================================
# ESTILOS
# =====================================================================
FONT_NAME = "Calibri"

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
CELL_FONT = Font(name=FONT_NAME, size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

STATUS_FILLS = {
    "RUPTURA COM DEMANDA":       PatternFill("solid", start_color="C00000", end_color="C00000"),
    "RUPTURA SEM VENDA RECENTE": PatternFill("solid", start_color="E6B8B7", end_color="E6B8B7"),
    "RISCO DE RUPTURA":          PatternFill("solid", start_color="FFC000", end_color="FFC000"),
    "SAUDAVEL":                  PatternFill("solid", start_color="C6E0B4", end_color="C6E0B4"),
    "EXCESSO":                   PatternFill("solid", start_color="9DC3E6", end_color="9DC3E6"),
    "SEM GIRO":                  PatternFill("solid", start_color="BFBFBF", end_color="BFBFBF"),
    "SEM HISTORICO DE VENDA":    PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA"),
}
STATUS_FONT_WHITE = {"RUPTURA COM DEMANDA"}

OCIOSIDADE_FILLS = {
    "Ativo":     PatternFill("solid", start_color="C6E0B4", end_color="C6E0B4"),
    "Parando":   PatternFill("solid", start_color="FFE699", end_color="FFE699"),
    "Parado":    PatternFill("solid", start_color="FFC000", end_color="FFC000"),
    "Inativo":   PatternFill("solid", start_color="BFBFBF", end_color="BFBFBF"),
    "Sem Venda": PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA"),
}

KPI_CARD_FILL = PatternFill("solid", start_color="EAF1FB", end_color="EAF1FB")
KPI_LABEL_FONT = Font(name=FONT_NAME, size=8.5, bold=True, color="1F4E78")
KPI_VALUE_FONT = Font(name=FONT_NAME, size=17, bold=True, color="1F4E78")
KPI_SIDE = Side(style="thin", color="9DC3E6")
KPI_BORDER = Border(left=KPI_SIDE, right=KPI_SIDE, top=KPI_SIDE, bottom=KPI_SIDE)
PAINEL_TITLE_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
PAINEL_NOTE_FONT = Font(name=FONT_NAME, size=8, italic=True, color="7F7F7F")

# Layout das colunas da tabela principal - ordem pedida pelo Jonas: da
# esquerda pra direita monta a linha de raciocinio da decisao de compra
# (o que e', situacao de venda/estoque, depois custo/fornecedor, depois
# decisao/valor). condicao_pagamento fica oculta no fim so porque a aba
# Carrinho ainda depende dela via formula - nao faz parte do layout pedido.
COLUNAS = [
    ("codigo", "Codigo", 10),
    ("nome", "Produto", 42),
    ("ociosidade", "Ociosidade", 12),
    ("estoque_atual", "Estoque atual", 11),
    ("cobertura_dias", "Cobertura (d)", 12),
    ("status", "Status estoque", 22),
    ("vendas_30d", "Vendas 30d", 10),
    ("vendas_60d", "Vendas 60d", 10),
    ("vendas_90d", "Vendas 90d", 10),
    ("velocidade_dia", "Vel./dia", 8),
    ("tendencia", "Tendencia", 10),
    ("dias_ate_ruptura", "Dias ate Ruptura", 13),
    ("ultima_venda", "Ultima venda", 12),
    ("dias_sem_vender", "Dias s/ venda", 12),
    ("classe_abc", "Classe ABC", 9),
    ("giro_financeiro", "Giro Financeiro", 13),
    ("rentabilidade_trimestral", "Rentabilidade Trimestral", 15),
    ("lead_time_dias", "Lead time forn. (d)", 14),
    ("risco_ruptura", "Risco de Ruptura", 13),
    ("prioridade", "Prioridade", 10),
    ("qtd_sugerida", "Qtd sugerida", 11),
    ("qtd_comprar", "Qtd a comprar", 11),
    ("ultima_compra", "Ultima compra", 12),
    ("lead_time_ultima_compra_dias", "Lead Time Ult. Compra (d)", 14),
    ("fornecedor", "Fornecedor", 24),
    ("origem_custo", "Origem custo", 12),
    ("custo_unit", "Custo unit.", 11),
    ("melhor_custo", "Melhor Custo", 11),
    ("melhor_fornecedor", "Melhor Fornecedor", 24),
    ("preco_venda", "Preco venda", 11),
    ("markup", "Markup", 8),
    ("margem_pct", "Margem %", 9),
    ("valor_sugerido", "Valor sugerido", 12),
    ("valor_comprar", "Valor a comprar", 12),
    ("score_compra", "Score Compra", 12),
    ("comprar", "Comprar?", 9),
    ("num_carrinho", "Nº no carrinho", 11),
    ("observacoes", "Observacoes", 24),
    ("categoria", "Categoria", 18),
    ("estoque_minimo", "Est. minimo", 10),
    ("estoque_maximo", "Est. maximo", 10),
    ("condicao_pagamento", "Cond. pagamento", 16),   # oculta - so p/ Carrinho
]
COL_INDEX = {chave: i + 1 for i, (chave, _, _) in enumerate(COLUNAS)}
COL_LETTER = {chave: get_column_letter(idx) for chave, idx in COL_INDEX.items()}
N_COLS = len(COLUNAS)


# =====================================================================
# LEITURA DOS CSVS
# =====================================================================
def ler_csv(nome):
    caminho = DATA_DIR / nome
    if not caminho.exists():
        print(f"[AVISO] arquivo nao encontrado: {caminho}")
        return []
    # Alguns exports do Bling trazem bytes NUL embutidos em campos de texto
    # (observado em produtos/notas_fiscais) - isso quebra o csv.reader, entao
    # filtramos antes de parsear em vez de deixar o script cair.
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        linhas_limpas = (linha.replace("\x00", "") for linha in f)
        return list(csv.DictReader(linhas_limpas))


def parse_float(v, default=0.0):
    if v is None or v == "":
        return default
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return default


def parse_date(v):
    if not v:
        return None
    s = str(v).strip()
    if not s:
        return None
    s = s.replace("T", " ")
    date_part = s.split(" ")[0]
    try:
        return dt.date.fromisoformat(date_part)
    except ValueError:
        return None


# =====================================================================
# CARREGAMENTO E PRE-PROCESSAMENTO
# =====================================================================
def carregar_dados():
    dados = {
        "produtos": ler_csv("produtos.csv"),
        "estoques": ler_csv("estoques_depositos.csv"),
        "pedidos_compra": ler_csv("pedidos_compra.csv"),
        "itens_compra": ler_csv("itens_compra.csv"),
        "notas_entrada": ler_csv("notas_entrada.csv"),
        "itens_notas_entrada": ler_csv("itens_notas_entrada.csv"),
        "pedidos_venda": ler_csv("pedidos_venda.csv"),
        "itens_venda": ler_csv("itens_venda.csv"),
        "contatos": ler_csv("contatos.csv"),
        "categorias": ler_csv("categorias.csv"),
        "situacoes": ler_csv("situacoes.csv"),
    }
    for nome, linhas in dados.items():
        print(f"  {nome:20s} {len(linhas):6d} linhas")
    return dados


def montar_indices(dados):
    """Constroi todos os dicionarios auxiliares usados nos calculos por produto."""
    situacoes_map = {r["situacao_id"]: r["descricao"] for r in dados["situacoes"] if r.get("situacao_id")}
    categorias_map = {r["categoria_id"]: r["descricao"] for r in dados["categorias"] if r.get("categoria_id")}
    contatos_map = {r["contato_id"]: r for r in dados["contatos"] if r.get("contato_id")}
    produtos_ids = {r.get("produto_id") for r in dados["produtos"] if r.get("produto_id")}
    produto_id_por_codigo = {(r.get("codigo") or "").strip(): r.get("produto_id") for r in dados["produtos"] if (r.get("codigo") or "").strip()}
    produto_id_por_codigo_antigo = {(r.get("codigoAntigo") or "").strip(): r.get("produto_id") for r in dados["produtos"] if (r.get("codigoAntigo") or "").strip()}
    de_para = {}
    for r in ler_csv("de_para_produtos.csv"):
        antigo = (r.get("codigo_antigo") or "").strip()
        atual = (r.get("codigo_atual") or "").strip()
        if antigo and atual and atual in produto_id_por_codigo:
            de_para[antigo] = produto_id_por_codigo[atual]

    def resolver_produto_id(item):
        pid = item.get("produto_id")
        if pid in produtos_ids:
            return pid
        codigo = (item.get("codigo") or "").strip()
        return produto_id_por_codigo.get(codigo) or produto_id_por_codigo_antigo.get(codigo) or de_para.get(codigo)

    def situacao_venda_valida(situacao_id):
        desc = situacoes_map.get(situacao_id, "")
        return "cancelad" not in desc.lower()

    def situacao_compra_aberta(situacao_id):
        desc = situacoes_map.get(situacao_id, "").lower()
        return ("abert" in desc) or ("andamento" in desc)

    # -------- Estoque atual por produto --------
    # estoques_depositos.csv deveria dar o saldo por deposito, mas nesta base
    # as colunas saldoFisico/saldoVirtual desse arquivo estao 100% vazias
    # (falha de sincronizacao desse endpoint especifico) - usamos o total
    # already agregado em produtos.estoque_saldoVirtual, que esta preenchido.
    estoque_por_produto = defaultdict(float)
    estoques_depositos_vazio = all(
        not (r.get("saldoVirtual") or "").strip() for r in dados["estoques"]
    ) if dados["estoques"] else True
    if not estoques_depositos_vazio:
        for r in dados["estoques"]:
            pid = r.get("produto_id")
            if pid:
                estoque_por_produto[pid] += parse_float(r.get("saldoVirtual"))
    else:
        print("  [AVISO] estoques_depositos.csv sem saldoVirtual preenchido - usando produtos.estoque_saldoVirtual")
        for p in dados["produtos"]:
            pid = p.get("produto_id")
            if pid:
                estoque_por_produto[pid] = parse_float(p.get("estoque_saldoVirtual"))

    # -------- Pedidos de venda validos: mapa pedido_id -> data --------
    pedidos_venda_map = {}
    for r in dados["pedidos_venda"]:
        pid = r.get("pedido_id")
        if not pid:
            continue
        pedidos_venda_map[pid] = {
            "valido": situacao_venda_valida(r.get("situacao_id", "")),
            "data": parse_date(r.get("data")),
        }

    # -------- Vendas por produto: 30/60/90d, ultima venda --------
    vendas_30 = defaultdict(float)
    vendas_60 = defaultdict(float)
    vendas_90 = defaultdict(float)
    faturamento_30 = defaultdict(float)
    faturamento_60 = defaultdict(float)
    faturamento_90 = defaultdict(float)
    ultima_venda = {}
    for r in dados["itens_venda"]:
        pid_pedido = r.get("pedido_id")
        info = pedidos_venda_map.get(pid_pedido)
        if not info or not info["valido"] or not info["data"]:
            continue
        produto_id = resolver_produto_id(r)
        if not produto_id:
            continue
        qtd = parse_float(r.get("quantidade"))
        # Valor real da linha do pedido: quantidade x preco unitario praticado,
        # abatendo o desconto do item. Nao usa o preco atual do cadastro.
        valor_real = max(0.0, qtd * parse_float(r.get("valor")) - parse_float(r.get("desconto")))
        dias_atras = (HOJE - info["data"]).days
        if 0 <= dias_atras <= 30:
            vendas_30[produto_id] += qtd
            faturamento_30[produto_id] += valor_real
        if 0 <= dias_atras <= 60:
            vendas_60[produto_id] += qtd
            faturamento_60[produto_id] += valor_real
        if 0 <= dias_atras <= 90:
            vendas_90[produto_id] += qtd
            faturamento_90[produto_id] += valor_real
        if qtd > 0:
            atual = ultima_venda.get(produto_id)
            if atual is None or info["data"] > atual:
                ultima_venda[produto_id] = info["data"]

    # tendencia: velocidade 0-30d vs 31-60d
    vendas_31_60 = defaultdict(float)
    for r in dados["itens_venda"]:
        info = pedidos_venda_map.get(r.get("pedido_id"))
        if not info or not info["valido"] or not info["data"]:
            continue
        produto_id = resolver_produto_id(r)
        if not produto_id:
            continue
        dias_atras = (HOJE - info["data"]).days
        if 31 <= dias_atras <= 60:
            vendas_31_60[produto_id] += parse_float(r.get("quantidade"))

    # -------- Pedidos de compra: mapa + itens em aberto por produto --------
    pedidos_compra_map = {r["pedido_id"]: r for r in dados["pedidos_compra"] if r.get("pedido_id")}
    qtd_aberta_por_produto = defaultdict(float)
    ultima_compra_pedido = {}
    custo_via_pedido = {}
    for r in dados["itens_compra"]:
        pid_pedido = r.get("pedido_id")
        pedido = pedidos_compra_map.get(pid_pedido)
        produto_id = r.get("produto_id")
        if not pedido or not produto_id:
            continue
        if situacao_compra_aberta(pedido.get("situacao_id", "")):
            qtd_aberta_por_produto[produto_id] += parse_float(r.get("quantidade"))
        data_pedido = parse_date(pedido.get("data"))
        if data_pedido:
            atual = ultima_compra_pedido.get(produto_id)
            if atual is None or data_pedido > atual:
                ultima_compra_pedido[produto_id] = data_pedido
            qtd_item = parse_float(r.get("quantidade"))
            if qtd_item > 0:
                # "valor" no CSV do Bling ja e' o preco unitario da linha, NAO
                # o total (isso e' "valorTotal") - nao dividir pela qtd de novo.
                custo_unit = parse_float(r.get("valor"))
                atual_custo = custo_via_pedido.get(produto_id)
                if custo_unit > 0 and (atual_custo is None or data_pedido >= atual_custo[1]):
                    custo_via_pedido[produto_id] = (custo_unit, data_pedido)

    # -------- Notas de entrada validas: custo mais confiavel + data --------
    notas_entrada_map = {r["nfe_id"]: r for r in dados["notas_entrada"] if r.get("nfe_id")}
    custo_via_nota = {}
    ultima_compra_nota = {}
    lead_time_ultima_compra = {}   # codigo -> dias entre emissao e entrada real, so da NF mais recente
    valor_comprado_mes = 0.0
    for nf in dados["notas_entrada"]:
        if nf.get("situacao") not in NF_ENTRADA_VALIDAS:
            continue
        # dataOperacao = quando a mercadoria realmente deu entrada no estoque
        # (o que aparece na tela de Estoque do Bling); dataEmissao e' so
        # quando o fornecedor emitiu a nota, pode ser dias antes. Usamos
        # dataOperacao com fallback pra dataEmissao quando ela vier vazia.
        data_op = parse_date(nf.get("dataOperacao")) or parse_date(nf.get("dataEmissao"))
        if data_op and data_op.year == HOJE.year and data_op.month == HOJE.month:
            valor_comprado_mes += parse_float(nf.get("totalNF"))

    for r in dados["itens_notas_entrada"]:
        nf = notas_entrada_map.get(r.get("nfe_id"))
        # A API do Bling NAO devolve produto_id nos itens de nota de entrada
        # (fica sempre vazio) - so vem "codigo" (SKU). Por isso o cruzamento
        # com o produto tem que ser feito por codigo, nunca por produto_id.
        codigo_item = (r.get("codigo") or "").strip()
        if not nf or not codigo_item or nf.get("situacao") not in NF_ENTRADA_VALIDAS:
            continue
        data_nf = parse_date(nf.get("dataOperacao")) or parse_date(nf.get("dataEmissao"))
        if not data_nf:
            continue
        atual = ultima_compra_nota.get(codigo_item)
        if atual is None or data_nf > atual:
            ultima_compra_nota[codigo_item] = data_nf
            # Lead time real dessa compra especifica: dias entre o fornecedor
            # emitir a nota (dataEmissao) e a mercadoria de fato entrar no
            # estoque (dataOperacao). So calcula se as duas datas existirem
            # e dataOperacao nao for anterior a dataEmissao (dado inconsistente).
            data_emissao_nf = parse_date(nf.get("dataEmissao"))
            data_operacao_nf = parse_date(nf.get("dataOperacao"))
            if data_emissao_nf and data_operacao_nf and data_operacao_nf >= data_emissao_nf:
                lead_time_ultima_compra[codigo_item] = (data_operacao_nf - data_emissao_nf).days
            else:
                lead_time_ultima_compra[codigo_item] = None
        qtd_item = parse_float(r.get("quantidade"))
        if qtd_item > 0:
            # Mesma correcao: "valor" ja e' o preco unitario ("valorTotal" que
            # e' o total da linha) - dividir por qtd de novo estava gerando
            # custo ~qtd vezes menor que o real (ex: 97,91 virava 8,16 numa
            # linha de 12 unidades).
            custo_unit = parse_float(r.get("valor"))
            atual_custo = custo_via_nota.get(codigo_item)
            if custo_unit > 0 and (atual_custo is None or data_nf >= atual_custo[1]):
                custo_via_nota[codigo_item] = (custo_unit, data_nf)

    # -------- Lead time por fornecedor --------
    # pedidos_compra.notaEntrada_id existe no schema mas esta vazio em 100%
    # dos 56 pedidos desta base (nao e possivel linkar pedido->nota por ID).
    # Heuristica: casar pedido de compra com nota de entrada do MESMO
    # fornecedor que (a) foi emitida depois do pedido, ate 90 dias depois, e
    # (b) tem valor total parecido com o do pedido (ate 20% de diferenca,
    # ou ate R$50 em pedidos pequenos). So exigir proximidade de data (sem
    # checar valor) deu lead times de ~3 dias irreais nesta base, porque ha
    # 225 notas de entrada para so 56 pedidos formais - ou seja, a maior
    # parte das compras nao passa por pedido de compra formal, e "a nota
    # mais proxima em data" quase sempre acerta por coincidencia, nao por
    # ser realmente a entrega daquele pedido. Exigir valor parecido reduz
    # bastante esse falso positivo (ainda e uma heuristica, nao um vinculo
    # direto por ID - ver nota no painel de KPIs).
    notas_por_fornecedor = defaultdict(list)
    for nf in dados["notas_entrada"]:
        forn_id = nf.get("contato_id")
        data_nf = parse_date(nf.get("dataOperacao")) or parse_date(nf.get("dataEmissao"))
        if forn_id and data_nf and nf.get("situacao") in NF_ENTRADA_VALIDAS:
            notas_por_fornecedor[forn_id].append((data_nf, parse_float(nf.get("totalNF"))))
    for forn_id in notas_por_fornecedor:
        notas_por_fornecedor[forn_id].sort(key=lambda t: t[0])

    usadas_por_fornecedor = defaultdict(set)
    lead_times_por_fornecedor = defaultdict(list)
    pedidos_ordenados = sorted(
        (p for p in dados["pedidos_compra"] if p.get("fornecedor_id") and parse_date(p.get("data"))),
        key=lambda p: parse_date(p.get("data")),
    )
    for pedido in pedidos_ordenados:
        forn_id = pedido.get("fornecedor_id")
        data_pedido = parse_date(pedido.get("data"))
        total_pedido = parse_float(pedido.get("total"))
        candidatas = notas_por_fornecedor.get(forn_id, [])
        melhor = None
        for i, (data_nf, total_nf) in enumerate(candidatas):
            if i in usadas_por_fornecedor[forn_id]:
                continue
            delta = (data_nf - data_pedido).days
            if not (0 <= delta <= 90):
                continue
            diff_valor = abs(total_nf - total_pedido)
            valor_compativel = diff_valor <= max(50.0, total_pedido * 0.20)
            if valor_compativel and (melhor is None or delta < melhor[1]):
                melhor = (i, delta)
        if melhor:
            usadas_por_fornecedor[forn_id].add(melhor[0])
            lead_times_por_fornecedor[forn_id].append(melhor[1])

    lead_time_medio_por_fornecedor = {
        forn_id: sum(deltas) / len(deltas)
        for forn_id, deltas in lead_times_por_fornecedor.items()
    }
    todos_deltas = [d for deltas in lead_times_por_fornecedor.values() for d in deltas]

    return {
        "situacoes_map": situacoes_map,
        "categorias_map": categorias_map,
        "contatos_map": contatos_map,
        "estoque_por_produto": estoque_por_produto,
        "vendas_30": vendas_30,
        "vendas_60": vendas_60,
        "vendas_90": vendas_90,
        "faturamento_30": faturamento_30,
        "faturamento_60": faturamento_60,
        "faturamento_90": faturamento_90,
        "vendas_31_60": vendas_31_60,
        "ultima_venda": ultima_venda,
        "qtd_aberta_por_produto": qtd_aberta_por_produto,
        "ultima_compra_pedido": ultima_compra_pedido,
        "custo_via_pedido": custo_via_pedido,
        "ultima_compra_nota": ultima_compra_nota,
        "custo_via_nota": custo_via_nota,
        "lead_time_ultima_compra": lead_time_ultima_compra,
        "lead_time_medio_por_fornecedor": lead_time_medio_por_fornecedor,
        "lead_time_geral_medio": (sum(todos_deltas) / len(todos_deltas)) if todos_deltas else None,
        "valor_comprado_mes": valor_comprado_mes,
    }


# =====================================================================
# CALCULO POR PRODUTO
# =====================================================================
def classificar_status(estoque, vendas_90d, cobertura_dias, dias_sem_vender, nunca_vendeu):
    if nunca_vendeu:
        return "SEM HISTORICO DE VENDA"
    if estoque <= 0:
        return "RUPTURA COM DEMANDA" if vendas_90d > 0 else "RUPTURA SEM VENDA RECENTE"
    if dias_sem_vender is not None and dias_sem_vender >= SEM_MOVIMENTO_LIMIAR_DIAS:
        return "SEM GIRO"
    if cobertura_dias is not None:
        if cobertura_dias < COBERTURA_MIN_DIAS:
            return "RISCO DE RUPTURA"
        if cobertura_dias > COBERTURA_MAX_DIAS:
            return "EXCESSO"
        return "SAUDAVEL"
    return "BAIXO GIRO"


# Limiares de ociosidade (dias sem vender). Independente do estoque - aqui o
# que importa e' ha quanto tempo o produto nao gira, nao se falta ou sobra
# peca fisica. "Parando" = ainda vende, mas a velocidade recente (30d) caiu
# forte frente aos 31-60d anteriores (mesma logica da coluna Tendencia).
OCIOSIDADE_PARADO_DIAS = SEM_MOVIMENTO_LIMIAR_DIAS   # 90 dias
OCIOSIDADE_INATIVO_DIAS = 180


def classificar_ociosidade(dias_sem_vender, tendencia, nunca_vendeu):
    if nunca_vendeu or dias_sem_vender is None:
        return "Sem Venda"
    if dias_sem_vender > OCIOSIDADE_INATIVO_DIAS:
        return "Inativo"
    if dias_sem_vender > OCIOSIDADE_PARADO_DIAS:
        return "Parado"
    if tendencia == "CAINDO":
        return "Parando"
    return "Ativo"


def montar_linhas(dados, idx):
    produtos_ativos = [p for p in dados["produtos"] if p.get("situacao") == "A"]

    # Classe ABC por faturamento nos ultimos 90 dias
    faturamento_90d = {}
    for p in produtos_ativos:
        pid = p.get("produto_id")
        faturamento_90d[pid] = idx["faturamento_90"].get(pid, 0.0)
    ranking = sorted(faturamento_90d.items(), key=lambda kv: kv[1], reverse=True)
    total_fat = sum(v for _, v in ranking) or 1.0
    classe_por_produto = {}
    acumulado = 0.0
    for pid, v in ranking:
        acumulado += v
        pct = acumulado / total_fat
        classe_por_produto[pid] = "A" if pct <= 0.80 else ("B" if pct <= 0.95 else "C")

    linhas = []
    for p in produtos_ativos:
        pid = p.get("produto_id")
        estoque = idx["estoque_por_produto"].get(pid)
        if estoque is None:
            estoque = parse_float(p.get("estoque_saldoVirtual"))

        v30 = idx["vendas_30"].get(pid, 0.0)
        v60 = idx["vendas_60"].get(pid, 0.0)
        v90 = idx["vendas_90"].get(pid, 0.0)
        fat30 = idx["faturamento_30"].get(pid, 0.0)
        fat60 = idx["faturamento_60"].get(pid, 0.0)
        fat90 = idx["faturamento_90"].get(pid, 0.0)
        v31_60 = idx["vendas_31_60"].get(pid, 0.0)
        ultima_venda = idx["ultima_venda"].get(pid)
        nunca_vendeu = ultima_venda is None
        dias_sem_vender = (HOJE - ultima_venda).days if ultima_venda else None

        velocidade_dia = v90 / JANELA_VELOCIDADE_DIAS
        cobertura_dias = (estoque / velocidade_dia) if velocidade_dia > 0 else None

        status = classificar_status(estoque, v90, cobertura_dias, dias_sem_vender, nunca_vendeu)

        if v30 > v31_60 * 1.15:
            tendencia = "SUBINDO"
        elif v30 < v31_60 * 0.85:
            tendencia = "CAINDO"
        else:
            tendencia = "ESTAVEL"

        ociosidade = classificar_ociosidade(dias_sem_vender, tendencia, nunca_vendeu)

        # ---- custo em cascata: nota de entrada > cadastro ----
        # Pedido de compra e' so uma intencao de compra (o que foi pedido ao
        # fornecedor), nao o custo que efetivamente foi pago - por isso NAO
        # entra mais nessa cascata; a nota de entrada e' a fonte real.
        codigo_produto = (p.get("codigo") or "").strip()
        custo_nota = idx["custo_via_nota"].get(codigo_produto)
        custo_cadastro = parse_float(p.get("precoCusto"))
        if custo_nota:
            custo_unit, origem_custo = custo_nota[0], "Nota de entrada"
        elif custo_cadastro > 0:
            custo_unit, origem_custo = custo_cadastro, "Cadastro"
        else:
            custo_unit, origem_custo = 0.0, "Sem custo confiavel"

        preco_venda = parse_float(p.get("preco"))
        margem_pct = ((preco_venda - custo_unit) / preco_venda) if preco_venda > 0 else None
        markup = (preco_venda / custo_unit) if custo_unit > 0 else None

        ultima_compra = idx["ultima_compra_nota"].get(codigo_produto) or idx["ultima_compra_pedido"].get(pid)
        dias_sem_comprar = (HOJE - ultima_compra).days if ultima_compra else None
        # Lead time REAL da ultima compra desse produto especifico (emissao ->
        # entrada no estoque), vindo da mesma nota de entrada que definiu
        # "Ultima compra" acima. So existe quando a origem foi nota de entrada
        # (pedido de compra nao tem essa granularidade de datas por item).
        lead_time_ultima_compra_dias = idx["lead_time_ultima_compra"].get(codigo_produto)

        qtd_aberta = idx["qtd_aberta_por_produto"].get(pid, 0.0)

        # ---- sugestao de compra ----
        # So sugerimos compra quando ha evidencia de demanda real (venda nos
        # ultimos 90 dias) ou o produto esta "saudavel"/"baixo giro" (ainda
        # vendendo, mesmo que pouco). Para SEM HISTORICO DE VENDA, SEM GIRO,
        # EXCESSO e RUPTURA SEM VENDA RECENTE nao sugerimos compra so por o
        # estoque estar abaixo do minimo cadastrado - o cadastro pode estar
        # desatualizado e comprar aqui seria repetir o problema de estoque
        # parado identificado no diagnostico empresarial.
        qtd_sugerida = 0.0
        if status in ("RUPTURA COM DEMANDA", "RISCO DE RUPTURA"):
            meta = velocidade_dia * META_COBERTURA_REPOSICAO_DIAS
            qtd_sugerida = max(0.0, round(meta - estoque - qtd_aberta))
        elif status in ("SAUDAVEL", "BAIXO GIRO"):
            est_min = parse_float(p.get("estoque_minimo"))
            est_max = parse_float(p.get("estoque_maximo"))
            if est_min > 0 and estoque < est_min and est_max > est_min:
                qtd_sugerida = max(0.0, round(est_max - estoque - qtd_aberta))
        valor_sugerido = qtd_sugerida * custo_unit

        fornecedor_id = p.get("fornecedor_id") or ""
        fornecedor_nome = p.get("fornecedor_nome") or ""
        contato_forn = idx["contatos_map"].get(fornecedor_id, {})
        condicao_pagamento = contato_forn.get("financeiro_condicaoPagamento", "")
        limite_credito = parse_float(contato_forn.get("financeiro_limiteCredito"), default=None)
        lead_time_forn = idx["lead_time_medio_por_fornecedor"].get(fornecedor_id)

        peso_bruto_unit = parse_float(p.get("pesoBruto"))
        peso_bruto_total = round(qtd_sugerida * peso_bruto_unit, 2) if peso_bruto_unit else None

        classe_abc_val = classe_por_produto.get(pid, "C")

        # ---- metricas novas de apoio a decisao ----
        # Dias ate ruptura: quanto tempo o estoque atual ainda aguenta no
        # ritmo de venda atual. 0 quando ja esta zerado; branco quando nao
        # ha velocidade de venda pra estimar (produto parado/sem historico).
        if estoque <= 0:
            dias_ate_ruptura = 0.0
        elif velocidade_dia > 0:
            dias_ate_ruptura = round(estoque / velocidade_dia, 1)
        else:
            dias_ate_ruptura = None

        # Giro financeiro: faturamento dos ultimos 30 dias a preco atual de
        # venda (proxy rapido de quanto dinheiro esse SKU gira por mes).
        giro_financeiro = round(fat30, 2)

        # Rentabilidade trimestral: lucro bruto estimado nos ultimos 90 dias
        # (preco de venda - custo) x volume vendido no trimestre.
        rentabilidade_trimestral = (
            round((preco_venda - custo_unit) * v90, 2) if (preco_venda and custo_unit) else None
        )

        lead_time_forn_val = round(lead_time_forn, 1) if lead_time_forn is not None else None

        # Risco de ruptura: so faz sentido enquanto ainda ha estoque - sinaliza
        # "SIM" quando a cobertura atual e' menor ou igual ao lead time medio
        # do fornecedor, ou seja, o estoque some antes do proximo pedido chegar.
        risco_ruptura = ""
        if estoque > 0 and cobertura_dias is not None and lead_time_forn_val is not None:
            if cobertura_dias <= lead_time_forn_val:
                risco_ruptura = "SIM"

        # Prioridade / Score Compra: cruza peso financeiro do SKU (Classe ABC)
        # com a urgencia do status de estoque, so quando ha sugestao de compra.
        # Score Compra e' a mesma logica em numero, pra poder ordenar a coluna.
        prioridade = ""
        score_compra = None
        if qtd_sugerida > 0:
            pontos = (2 if classe_abc_val == "A" else 1 if classe_abc_val == "B" else 0)
            if status in ("RUPTURA COM DEMANDA", "RUPTURA SEM VENDA RECENTE"):
                pontos += 2
            elif status == "RISCO DE RUPTURA":
                pontos += 1
            prioridade = "ALTA" if pontos >= 3 else ("MEDIA" if pontos >= 1 else "BAIXA")
            score_compra = pontos * 100000 + (giro_financeiro or 0)

        # Melhor Custo / Melhor Fornecedor: hoje o script so acompanha um
        # fornecedor por produto (o do cadastro), entao por enquanto isso e'
        # igual a Custo unit./Fornecedor. Fica pronto pra quando houver
        # comparacao entre varios fornecedores do mesmo produto.
        melhor_custo = custo_unit if custo_unit else None
        melhor_fornecedor = fornecedor_nome

        linha = {
            "codigo": p.get("codigo", ""),
            "nome": p.get("nome", ""),
            "categoria": idx["categorias_map"].get(p.get("categoria_id"), ""),
            "fornecedor": fornecedor_nome,
            "melhor_custo": round(melhor_custo, 2) if melhor_custo else None,
            "melhor_fornecedor": melhor_fornecedor,
            "estoque_atual": estoque,
            "estoque_minimo": parse_float(p.get("estoque_minimo"), default=None),
            "estoque_maximo": parse_float(p.get("estoque_maximo"), default=None),
            "vendas_30d": v30,
            "vendas_60d": v60,
            "vendas_90d": v90,
            "faturamento_30d": round(fat30, 2),
            "faturamento_60d": round(fat60, 2),
            "faturamento_90d": round(fat90, 2),
            "velocidade_dia": round(velocidade_dia, 3),
            "cobertura_dias": round(cobertura_dias, 1) if cobertura_dias is not None else None,
            "dias_ate_ruptura": dias_ate_ruptura,
            "dias_sem_vender": dias_sem_vender,
            "ultima_venda": ultima_venda,
            "ultima_compra": ultima_compra,
            "dias_sem_comprar": dias_sem_comprar,
            "lead_time_ultima_compra_dias": lead_time_ultima_compra_dias,
            "preco_venda": preco_venda if preco_venda else None,
            "custo_unit": custo_unit if custo_unit else None,
            "origem_custo": origem_custo,
            "margem_pct": round(margem_pct * 100, 1) if margem_pct is not None else None,
            "markup": round(markup, 2) if markup is not None else None,
            "status": status,
            "tendencia": tendencia,
            "ociosidade": ociosidade,
            "classe_abc": classe_abc_val,
            "giro_financeiro": giro_financeiro,
            "rentabilidade_trimestral": rentabilidade_trimestral,
            "qtd_pedidos_abertos": qtd_aberta if qtd_aberta else None,
            "lead_time_dias": lead_time_forn_val,
            "risco_ruptura": risco_ruptura,
            "prioridade": prioridade,
            "score_compra": score_compra,
            "qtd_sugerida": qtd_sugerida if qtd_sugerida else None,
            "valor_sugerido": round(valor_sugerido, 2) if valor_sugerido else None,
            "peso_bruto_total": peso_bruto_total,
            "condicao_pagamento": condicao_pagamento,
            "limite_credito": limite_credito,
            "comprar": False,
            "qtd_comprar": qtd_sugerida if qtd_sugerida else None,
            "valor_comprar": None,   # calculado via formula na planilha
            "observacoes": "",
            "num_carrinho": None,     # calculado via formula na planilha
            "_produto_id": pid,
            "_custo_unit_raw": custo_unit,
        }
        linhas.append(linha)

    ordem_status = {
        "RUPTURA COM DEMANDA": 0, "RISCO DE RUPTURA": 1, "RUPTURA SEM VENDA RECENTE": 2,
        "BAIXO GIRO": 3, "SAUDAVEL": 4, "EXCESSO": 5, "SEM GIRO": 6, "SEM HISTORICO DE VENDA": 7,
    }
    linhas.sort(key=lambda r: (ordem_status.get(r["status"], 9), r["nome"] or ""))
    return linhas


# =====================================================================
# PAINEL DE KPIS (linhas 1-7)
# =====================================================================
def calcular_valor_comprado_mes(idx):
    return idx["valor_comprado_mes"]


def calcular_prazo_medio_recebimento(idx):
    return idx["lead_time_geral_medio"]


def _kpi_card(ws, row_top, col_left, col_right, label, valor_texto, value_font=None):
    ws.merge_cells(start_row=row_top, start_column=col_left, end_row=row_top, end_column=col_right)
    ws.merge_cells(start_row=row_top + 1, start_column=col_left, end_row=row_top + 1, end_column=col_right)
    for r in (row_top, row_top + 1):
        for c in range(col_left, col_right + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = KPI_CARD_FILL
            cell.border = KPI_BORDER
    lbl = ws.cell(row=row_top, column=col_left, value=label)
    lbl.font = KPI_LABEL_FONT
    lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    val = ws.cell(row=row_top + 1, column=col_left, value=valor_texto)
    val.font = value_font or KPI_VALUE_FONT
    val.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def desenhar_painel_kpis(ws, painel):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N_COLS)
    titulo = ws.cell(row=1, column=1, value=f"PAINEL DE COMPRAS — AGM AUTO PARTES   |   gerado em {HOJE.strftime('%d/%m/%Y')}")
    titulo.fill = PAINEL_TITLE_FILL
    titulo.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    titulo.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 20

    # linha 2-3: 4 cards
    largura = N_COLS // 4
    limites = [1, 1 + largura, 1 + 2 * largura, 1 + 3 * largura, N_COLS + 1]
    cards_topo = [
        ("Teto de compras no mes", f"R$ {painel['teto']:,.0f}".replace(",", ".")),
        ("Comprado no mes atual", f"R$ {painel['comprado_mes']:,.0f}".replace(",", ".")),
        ("Valor de custo em estoque", f"R$ {painel['valor_estoque']:,.0f}".replace(",", ".")),
        ("Produtos parados (sem giro/excesso)", f"{painel['n_parados']}"),
    ]
    for i, (label, valor) in enumerate(cards_topo):
        _kpi_card(ws, 2, limites[i], limites[i + 1] - 1, label, valor)

    # linha 4-5: 3 cards
    largura2 = N_COLS // 3
    limites2 = [1, 1 + largura2, 1 + 2 * largura2, N_COLS + 1]
    cards_baixo = [
        ("Produtos que precisam de compra", f"{painel['n_a_comprar']}"),
        ("Lead time medio dos fornecedores", f"{painel['lead_time']:.0f} dias" if painel['lead_time'] is not None else "sem dados"),
        ("% do teto ja utilizado no mes", f"{painel['pct_teto']:.0f}%"),
    ]
    for i, (label, valor) in enumerate(cards_baixo):
        _kpi_card(ws, 4, limites2[i], limites2[i + 1] - 1, label, valor)

    # linha 6: barra de progresso (uso do teto de compras)
    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=N_COLS)
    barra = ws.cell(row=6, column=1, value=min(painel["pct_teto"], 999) / 100.0)
    barra.number_format = "0%"
    barra.alignment = Alignment(horizontal="center", vertical="center")
    ws.conditional_formatting.add(
        f"A6:{get_column_letter(N_COLS)}6",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                    color="1F4E78", showValue=True, minLength=None, maxLength=None),
    )

    # linha 7: nota explicativa
    ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=N_COLS)
    texto_nota = (
        "Nota: nao ha KPI de \"prazo medio de pagamento real\" (boleto/pix) porque a API do Bling nao "
        "retorna a data efetiva de pagamento de contas a pagar (confirmado na documentacao oficial). "
        "O lead time acima e estimado casando pedido de compra com a nota de entrada mais proxima do "
        "mesmo fornecedor (por data e valor) - so 56 pedidos formais existem contra 225 notas de entrada, "
        "entao e uma aproximacao, nao um vinculo direto."
    )
    if painel.get("lead_time_suspeito"):
        texto_nota += (" ATENCAO: varios fornecedores deram 0 dias, o que sugere que o pedido de compra "
                        "e lancado no Bling junto com a nota de entrada (nao antes) - trate este numero com cautela.")
    nota = ws.cell(row=7, column=1, value=texto_nota)
    nota.font = PAINEL_NOTE_FONT
    nota.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[7].height = 26


# =====================================================================
# MONTAGEM DO WORKBOOK
# =====================================================================
def montar_workbook(linhas, idx, painel):
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Sugestao_Compras")
    desenhar_painel_kpis(ws, painel)

    # cabecalho da tabela
    for col_idx, (chave, titulo, largura) in enumerate(COLUNAS, start=1):
        cell = ws.cell(row=HEADER_ROW, column=col_idx, value=titulo)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = largura
    ws.column_dimensions[COL_LETTER["condicao_pagamento"]].hidden = True
    ws.row_dimensions[HEADER_ROW].height = 30
    ws.freeze_panes = f"E{HEADER_ROW + 1}"

    campos = [c[0] for c in COLUNAS]
    for i, linha in enumerate(linhas, start=HEADER_ROW + 1):
        for col_idx, campo in enumerate(campos, start=1):
            valor = linha.get(campo)
            if isinstance(valor, dt.date):
                cell = ws.cell(row=i, column=col_idx, value=valor)
                cell.number_format = "dd/mm/yyyy"
            else:
                cell = ws.cell(row=i, column=col_idx, value=valor)
            cell.font = CELL_FONT
            cell.border = BORDER

        r = i
        status_cell = ws.cell(row=r, column=COL_INDEX["status"])
        fill = STATUS_FILLS.get(linha["status"])
        if fill:
            status_cell.fill = fill
            if linha["status"] in STATUS_FONT_WHITE:
                status_cell.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")

        ociosidade_cell = ws.cell(row=r, column=COL_INDEX["ociosidade"])
        fill_oc = OCIOSIDADE_FILLS.get(linha["ociosidade"])
        if fill_oc:
            ociosidade_cell.fill = fill_oc

        ws.cell(row=r, column=COL_INDEX["margem_pct"]).number_format = '0.0"%"'
        ws.cell(row=r, column=COL_INDEX["preco_venda"]).number_format = "#,##0.00"
        ws.cell(row=r, column=COL_INDEX["custo_unit"]).number_format = "#,##0.00"
        ws.cell(row=r, column=COL_INDEX["melhor_custo"]).number_format = "#,##0.00"
        ws.cell(row=r, column=COL_INDEX["giro_financeiro"]).number_format = "#,##0.00"
        ws.cell(row=r, column=COL_INDEX["rentabilidade_trimestral"]).number_format = "#,##0.00"
        ws.cell(row=r, column=COL_INDEX["valor_sugerido"]).number_format = "#,##0.00"

        # formulas: valor a comprar e numero no carrinho
        col_comprar = COL_LETTER["comprar"]
        col_qtd_comprar = COL_LETTER["qtd_comprar"]
        col_custo = COL_LETTER["custo_unit"]
        col_valor_comprar = COL_LETTER["valor_comprar"]
        col_num_carrinho = COL_LETTER["num_carrinho"]

        ws.cell(row=r, column=COL_INDEX["valor_comprar"],
                value=f'=IF({col_comprar}{r}=TRUE,{col_qtd_comprar}{r}*{col_custo}{r},"")')
        ws.cell(row=r, column=COL_INDEX["valor_comprar"]).number_format = "#,##0.00"
        ws.cell(row=r, column=COL_INDEX["num_carrinho"],
                value=f'=IF({col_comprar}{r}=TRUE,COUNTIF(${col_comprar}${HEADER_ROW + 1}:{col_comprar}{r},TRUE),"")')

    last_row = HEADER_ROW + len(linhas)

    # checkbox "Comprar?" com validacao de lista TRUE/FALSE
    dv = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws.add_data_validation(dv)
    col_comprar_letra = COL_LETTER["comprar"]
    dv.add(f"{col_comprar_letra}{HEADER_ROW + 1}:{col_comprar_letra}{last_row}")
    for r in range(HEADER_ROW + 1, last_row + 1):
        ws.cell(row=r, column=COL_INDEX["comprar"], value=False)

    # destaque visual: linha marcada como "Comprar?" = TRUE fica verde
    ws.conditional_formatting.add(
        f"{col_comprar_letra}{HEADER_ROW + 1}:{col_comprar_letra}{last_row}",
        CellIsRule(operator="equal", formula=["TRUE"],
                   fill=PatternFill("solid", start_color="C6E0B4", end_color="C6E0B4")),
    )

    # tabela nomeada
    ultima_col = get_column_letter(N_COLS)
    tabela = Table(displayName="Sugestao_Compras_Tbl", ref=f"A{HEADER_ROW}:{ultima_col}{last_row}")
    tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tabela)

    # destaque condicional: cobertura de estoque baixa
    col_cob = COL_LETTER["cobertura_dias"]
    ws.conditional_formatting.add(
        f"{col_cob}{HEADER_ROW + 1}:{col_cob}{last_row}",
        CellIsRule(operator="lessThan", formula=[str(COBERTURA_MIN_DIAS)],
                   fill=PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")),
    )
    # destaque: valor a comprar alto (acima de 10% do teto num item so)
    col_valor_comprar_letra = COL_LETTER["valor_comprar"]
    ws.conditional_formatting.add(
        f"{col_valor_comprar_letra}{HEADER_ROW + 1}:{col_valor_comprar_letra}{last_row}",
        CellIsRule(operator="greaterThan", formula=[str(TETO_COMPRAS_MES * 0.10)],
                   fill=PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")),
    )

    _montar_aba_carrinho(wb, last_row, painel)
    _montar_aba_parametros(wb)

    return wb


def _montar_aba_carrinho(wb, last_row_sugestao, painel):
    ws = wb.create_sheet("Carrinho")
    titulo = ws.cell(row=1, column=1, value="Carrinho de compras — itens marcados em Sugestao_Compras")
    titulo.font = Font(name=FONT_NAME, size=12, bold=True, color="1F4E78")
    ws.merge_cells("A1:H1")

    headers = ["Nº", "Codigo", "Produto", "Fornecedor", "Qtd a comprar", "Custo unit.", "Valor total", "Cond. pagamento"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER
    ws.freeze_panes = "A4"

    max_itens = last_row_sugestao - HEADER_ROW
    src = "Sugestao_Compras"
    col = {
        "codigo": COL_LETTER["codigo"], "nome": COL_LETTER["nome"], "fornecedor": COL_LETTER["fornecedor"],
        "qtd_comprar": COL_LETTER["qtd_comprar"], "custo": COL_LETTER["custo_unit"],
        "valor": COL_LETTER["valor_comprar"], "num_carrinho": COL_LETTER["num_carrinho"],
        "condicao": COL_LETTER["condicao_pagamento"],
    }
    faixa = f"{HEADER_ROW + 1}:{last_row_sugestao}"

    for i in range(1, max_itens + 1):
        r = i + 3
        pos_formula = f"MATCH({i},{src}!${col['num_carrinho']}${faixa.split(':')[0]}:${col['num_carrinho']}${faixa.split(':')[1]},0)"
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=f'=IFERROR(INDEX({src}!${col["codigo"]}:${col["codigo"]},{pos_formula}+{HEADER_ROW}),"")')
        ws.cell(row=r, column=3, value=f'=IFERROR(INDEX({src}!${col["nome"]}:${col["nome"]},{pos_formula}+{HEADER_ROW}),"")')
        ws.cell(row=r, column=4, value=f'=IFERROR(INDEX({src}!${col["fornecedor"]}:${col["fornecedor"]},{pos_formula}+{HEADER_ROW}),"")')
        ws.cell(row=r, column=5, value=f'=IFERROR(INDEX({src}!${col["qtd_comprar"]}:${col["qtd_comprar"]},{pos_formula}+{HEADER_ROW}),"")')
        ws.cell(row=r, column=6, value=f'=IFERROR(INDEX({src}!${col["custo"]}:${col["custo"]},{pos_formula}+{HEADER_ROW}),"")')
        ws.cell(row=r, column=7, value=f'=IFERROR(INDEX({src}!${col["valor"]}:${col["valor"]},{pos_formula}+{HEADER_ROW}),"")')
        ws.cell(row=r, column=8, value=f'=IFERROR(INDEX({src}!${col["condicao"]}:${col["condicao"]},{pos_formula}+{HEADER_ROW}),"")')
        for c in range(1, 9):
            ws.cell(row=r, column=c).font = CELL_FONT
            ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=6).number_format = "#,##0.00"
        ws.cell(row=r, column=7).number_format = "#,##0.00"

    total_row = max_itens + 5
    ws.cell(row=total_row, column=6, value="TOTAL SELECIONADO:").font = Font(name=FONT_NAME, bold=True)
    ws.cell(row=total_row, column=7, value=f"=SUM(G4:G{max_itens + 3})").font = Font(name=FONT_NAME, bold=True)
    ws.cell(row=total_row, column=7).number_format = "#,##0.00"
    ws.cell(row=total_row + 1, column=6, value="Teto do mes:").font = Font(name=FONT_NAME, italic=True, size=9)
    ws.cell(row=total_row + 1, column=7, value=f"R$ {painel['teto']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")).font = Font(name=FONT_NAME, italic=True, size=9)

    larguras = [5, 10, 40, 24, 12, 11, 12, 16]
    for c, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w


def _montar_aba_parametros(wb):
    ws = wb.create_sheet("Parametros")
    ws.cell(row=1, column=1, value="Parametros usados nesta geracao").font = Font(name=FONT_NAME, size=12, bold=True)
    params = [
        ("Data de geracao", HOJE.strftime("%d/%m/%Y")),
        ("Teto de compras no mes (R$)", TETO_COMPRAS_MES),
        ("Limiar 'sem giro' (dias sem venda)", SEM_MOVIMENTO_LIMIAR_DIAS),
        ("Cobertura minima antes de 'risco de ruptura' (dias)", COBERTURA_MIN_DIAS),
        ("Cobertura acima da qual e 'excesso' (dias)", COBERTURA_MAX_DIAS),
        ("Janela usada p/ velocidade de venda (dias)", JANELA_VELOCIDADE_DIAS),
        ("Cobertura-alvo ao sugerir reposicao (dias)", META_COBERTURA_REPOSICAO_DIAS),
    ]
    for i, (label, valor) in enumerate(params, start=3):
        c1 = ws.cell(row=i, column=1, value=label)
        c2 = ws.cell(row=i, column=2, value=valor)
        c1.font = Font(name=FONT_NAME, size=10)
        c2.font = Font(name=FONT_NAME, size=10, bold=True, color="1F4E78")
    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 16

    nota_row = len(params) + 5
    ws.cell(row=nota_row, column=1,
            value=("Como usar: marque \"TRUE\" na coluna Comprar? da aba Sugestao_Compras (ou selecione da "
                   "lista) para os itens que quer comprar, ajuste a Qtd a comprar se precisar, e confira a "
                   "aba Carrinho para o resumo e o valor total selecionado."))
    ws.cell(row=nota_row, column=1).font = Font(name=FONT_NAME, size=9, italic=True, color="7F7F7F")
    ws.merge_cells(start_row=nota_row, start_column=1, end_row=nota_row + 2, end_column=6)
    ws.cell(row=nota_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")


# =====================================================================
# MAIN
# =====================================================================
def salvar_snapshot_site(linhas, painel, dados):
    produtos = []
    for l in linhas:
        produtos.append({
            "code": l["codigo"], "name": l["nome"], "idle": l["ociosidade"],
            "stock": l["estoque_atual"], "coverage": l["cobertura_dias"], "status": l["status"],
            "sales30": l["vendas_30d"], "sales60": l["vendas_60d"], "sales90": l["vendas_90d"],
            "revenue30": l["faturamento_30d"], "revenue60": l["faturamento_60d"], "revenue90": l["faturamento_90d"],
            "velocity": l["velocidade_dia"], "trend": l["tendencia"], "daysToBreak": l["dias_ate_ruptura"],
            "abc": l["classe_abc"], "lead": l["lead_time_dias"], "priority": l["prioridade"],
            "suggestedQty": l["qtd_sugerida"], "buyQty": l["qtd_comprar"], "supplier": l["fornecedor"],
            "cost": l["custo_unit"], "price": l["preco_venda"], "margin": l["margem_pct"],
            "suggestedValue": l["valor_sugerido"], "score": l["score_compra"], "category": l["categoria"],
            "daysSinceSale": l["dias_sem_vender"], "daysSincePurchase": l["dias_sem_comprar"],
            "costOrigin": l["origem_custo"],
            "lastSaleDate": l["ultima_venda"].isoformat() if l["ultima_venda"] else None,
            "lastPurchaseDate": l["ultima_compra"].isoformat() if l["ultima_compra"] else None,
            "lastPurchaseLeadTime": l["lead_time_ultima_compra_dias"],
            "openPurchaseQty": l["qtd_pedidos_abertos"] or 0,
        })
    codigos = [p["code"] for p in produtos]
    if any(not c for c in codigos):
        raise RuntimeError("Snapshot recusado: existe produto sem codigo")
    if len(codigos) != len(set(codigos)):
        raise RuntimeError("Snapshot recusado: existem codigos duplicados")
    codigo_checksum = "\n".join(sorted(codigos))
    resumo = {
        "teto": painel["teto"], "compradoMes": painel["comprado_mes"], "valorEstoque": painel["valor_estoque"],
        "produtosParados": painel["n_parados"], "produtosComprar": painel["n_a_comprar"], "leadTimeMedio": painel["lead_time"],
    }
    itens_por_nota = defaultdict(list)
    for item in dados["itens_notas_entrada"]:
        codigo = (item.get("codigo") or "").strip()
        quantidade = parse_float(item.get("quantidade"))
        if item.get("nfe_id") and codigo and quantidade > 0:
            custo_unitario = parse_float(item.get("valor"))
            valor_total = parse_float(item.get("valorTotal"))
            if valor_total <= 0 and custo_unitario > 0:
                valor_total = quantidade * custo_unitario
            itens_por_nota[item["nfe_id"]].append({
                "code": codigo,
                "name": (item.get("descricao") or "").strip(),
                "quantity": quantidade,
                "unitCost": custo_unitario if custo_unitario > 0 else None,
                "totalValue": valor_total if valor_total > 0 else None,
            })
    recebimentos = []
    for nota in dados["notas_entrada"]:
        entrada = parse_date(nota.get("dataOperacao"))
        itens = itens_por_nota.get(nota.get("nfe_id"), [])
        if nota.get("situacao") not in NF_ENTRADA_VALIDAS or not entrada or not itens:
            continue
        emissao = parse_date(nota.get("dataEmissao"))
        recebimentos.append({
            "noteNumber": nota.get("numero", ""), "supplier": nota.get("contato_nome", ""),
            "issueDate": emissao.isoformat() if emissao else None,
            "entryDate": entrada.isoformat(),
            "totalValue": parse_float(nota.get("totalNF")) or sum((item.get("totalValue") or 0) for item in itens),
            "items": itens,
        })
    snapshot = {
        "schemaVersion": 2, "generatedAt": dt.datetime.now().isoformat(timespec="seconds"),
        "expectedCount": len(produtos), "checksum": hashlib.sha256(codigo_checksum.encode("utf-8")).hexdigest(),
        "summary": resumo, "products": produtos, "receipts": recebimentos,
    }
    SNAPSHOT_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp = SNAPSHOT_PATH.with_suffix(".json.tmp")
    temp.write_text(json.dumps(snapshot, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    temp.replace(SNAPSHOT_PATH)
    print(f"  Snapshot do site validado: {len(produtos)} produtos")


def main():
    print("Carregando dados sincronizados do Bling...")
    dados = carregar_dados()

    print("\nCalculando indices (vendas, estoque, compras, lead time)...")
    idx = montar_indices(dados)

    print("Montando linhas por produto...")
    linhas = montar_linhas(dados, idx)

    n_a_comprar = sum(1 for l in linhas if l.get("qtd_sugerida"))
    n_parados = sum(1 for l in linhas if l["status"] in ("SEM GIRO", "EXCESSO", "SEM HISTORICO DE VENDA"))
    valor_estoque_total = sum(
        (l["estoque_atual"] or 0) * (l["_custo_unit_raw"] or 0)
        for l in linhas if (l["estoque_atual"] or 0) > 0
    )
    comprado_mes = calcular_valor_comprado_mes(idx)
    lead_time_medio = calcular_prazo_medio_recebimento(idx)
    leads_por_forn = list(idx["lead_time_medio_por_fornecedor"].values())
    n_zero = sum(1 for v in leads_por_forn if v <= 0.5)
    lead_time_suspeito = bool(leads_por_forn) and (n_zero / len(leads_por_forn)) >= 0.4

    painel = {
        "teto": TETO_COMPRAS_MES,
        "comprado_mes": comprado_mes,
        "valor_estoque": valor_estoque_total,
        "n_parados": n_parados,
        "n_a_comprar": n_a_comprar,
        "lead_time": lead_time_medio,
        "lead_time_suspeito": lead_time_suspeito,
        "pct_teto": (comprado_mes / TETO_COMPRAS_MES * 100) if TETO_COMPRAS_MES else 0,
    }

    print(f"\nProdutos ativos processados: {len(linhas)}")
    print(f"  Produtos a comprar (sugestao > 0): {n_a_comprar}")
    print(f"  Produtos parados (sem giro/excesso/sem historico): {n_parados}")
    print(f"  Valor de custo em estoque: R$ {valor_estoque_total:,.2f}")
    print(f"  Comprado este mes: R$ {comprado_mes:,.2f}  ({painel['pct_teto']:.1f}% do teto de R$ {TETO_COMPRAS_MES:,.2f})")
    if lead_time_medio is not None:
        print(f"  Lead time medio dos fornecedores: {lead_time_medio:.1f} dias")
    else:
        print("  Lead time medio dos fornecedores: sem dados suficientes")

    print("\nMontando planilha...")
    wb = montar_workbook(linhas, idx, painel)

    print("Gerando snapshot seguro para o AGM 2.0...")
    salvar_snapshot_site(linhas, painel, dados)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"\nOK - planilha salva em: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
