"""Publica no AGM 2.0 somente um snapshot previamente validado.

Credenciais obrigatoriamente em variaveis de ambiente:
  AGM_SITE_SYNC_KEY       chave gerada na tela Configurar integracao
  CF_ACCESS_CLIENT_ID    ID do token de servico Cloudflare Access
  CF_ACCESS_CLIENT_SECRET segredo do token de servico Cloudflare Access
"""
import datetime as dt
import json
import os
import ssl
import subprocess
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# No Windows original, essas credenciais vinham de variaveis de ambiente
# do sistema (setx), definidas uma vez fora do .env. No Mac (e em qualquer
# outra maquina nova) isso nao existe ainda, entao carregamos do mesmo
# .env que os outros scripts (bling_auth.py) ja usam -- evita ter que
# configurar a credencial em dois lugares diferentes.
try:
    from dotenv import load_dotenv
    load_dotenv(BASE_DIR / ".env")
except ImportError:
    pass
SNAPSHOT_PATH = BASE_DIR / "data" / "monalisa_snapshot.json"
SITE_URL = os.environ.get("AGM_SITE_URL", "https://atlas-agm.comprasagmautopartes.workers.dev").rstrip("/")
SYNC_KEY = os.environ.get("AGM_SITE_SYNC_KEY", "").strip()
CF_ACCESS_CLIENT_ID = os.environ.get("CF_ACCESS_CLIENT_ID", "").strip()
CF_ACCESS_CLIENT_SECRET = os.environ.get("CF_ACCESS_CLIENT_SECRET", "").strip()
CHUNK_SIZE = 200

def post_via_windows(payload):
    """Usa o HTTPS nativo do Windows quando o OpenSSL do Python estiver
    com uma cadeia de certificados incompatível. A validação TLS continua
    ativa; nenhum verify=False é utilizado."""
    command = r'''
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$body = Get-Content -Raw -Encoding UTF8 -LiteralPath $env:AGM_PAYLOAD_PATH
$headers = @{ "x-agm-sync-key" = $env:AGM_SITE_SYNC_KEY }
if ($env:CF_ACCESS_CLIENT_ID -and $env:CF_ACCESS_CLIENT_SECRET) {
    $headers["CF-Access-Client-Id"] = $env:CF_ACCESS_CLIENT_ID
    $headers["CF-Access-Client-Secret"] = $env:CF_ACCESS_CLIENT_SECRET
}
$result = Invoke-RestMethod `
    -Uri ($env:AGM_SITE_URL.TrimEnd('/') + "/api/sync") `
    -Method Post `
    -Headers $headers `
    -ContentType "application/json; charset=utf-8" `
    -Body ([Text.Encoding]::UTF8.GetBytes($body))
$result | ConvertTo-Json -Compress -Depth 10
'''
    child_env = os.environ.copy()
    child_env["AGM_SITE_URL"] = SITE_URL
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", suffix=".json", delete=False
        ) as temp:
            json.dump(payload, temp, ensure_ascii=False, separators=(",", ":"))
            temp_path = temp.name
        child_env["AGM_PAYLOAD_PATH"] = temp_path
        completed = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
            text=True, encoding="utf-8", errors="replace", capture_output=True,
            env=child_env, timeout=90,
        )
    finally:
        if temp_path:
            try:
                os.unlink(temp_path)
            except FileNotFoundError:
                pass
    if completed.returncode != 0:
        message = (completed.stderr or completed.stdout).strip()
        raise RuntimeError("Falha HTTPS do Windows: " + message)
    result = json.loads(completed.stdout)
    if not result.get("ok"):
        raise RuntimeError(str(result))
    return result

def is_certificate_error(exc):
    reason = getattr(exc, "reason", None)
    return isinstance(reason, ssl.SSLCertVerificationError) or "CERTIFICATE_VERIFY_FAILED" in str(exc)

def post(payload):
    if os.name == "nt":
        return post_via_windows(payload)
    headers={
        "Content-Type":"application/json",
        "x-agm-sync-key":SYNC_KEY,
        # Sem um User-Agent "de navegador", o Cloudflare trata o
        # Python-urllib padrao como bot e bloqueia com "error code: 1010"
        # antes mesmo de checar as credenciais. No Windows original isso
        # nunca aparecia porque o script sempre usava post_via_windows()
        # (Invoke-RestMethod do PowerShell, que tem seu proprio User-Agent).
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    }
    if CF_ACCESS_CLIENT_ID and CF_ACCESS_CLIENT_SECRET:
        headers["CF-Access-Client-Id"]=CF_ACCESS_CLIENT_ID
        headers["CF-Access-Client-Secret"]=CF_ACCESS_CLIENT_SECRET
    body=json.dumps(payload,ensure_ascii=False,separators=(",",":")).encode("utf-8")
    last=None
    for attempt in range(1,6):
        try:
            req=urllib.request.Request(SITE_URL+"/api/sync",data=body,headers=headers,method="POST")
            with urllib.request.urlopen(req,timeout=60) as response:
                result=json.loads(response.read().decode("utf-8"))
                if not result.get("ok"): raise RuntimeError(str(result))
                return result
        except urllib.error.URLError as exc:
            if is_certificate_error(exc) and os.name == "nt":
                print("  HTTPS do Python incompativel; usando validacao nativa do Windows...")
                return post_via_windows(payload)
            last=exc
            if isinstance(exc, urllib.error.HTTPError):
                try:
                    err_body = exc.read().decode("utf-8", errors="replace")[:500]
                    print("  [tentativa " + str(attempt) + "] HTTP " + str(exc.code) + ": " + err_body)
                except Exception:
                    pass
            if attempt<5: time.sleep(min(2**attempt,20))
        except (TimeoutError,RuntimeError) as exc:
            last=exc
            if attempt<5: time.sleep(min(2**attempt,20))
    raise RuntimeError("Falha ao publicar no AGM 2.0 apos 5 tentativas: "+str(last))

def main():
    if len(SYNC_KEY)<32: raise RuntimeError("AGM_SITE_SYNC_KEY ausente. Gere a chave no site e salve como variavel de ambiente.")
    if not CF_ACCESS_CLIENT_ID or not CF_ACCESS_CLIENT_SECRET: raise RuntimeError("Credenciais do Cloudflare Access ausentes.")
    if not SNAPSHOT_PATH.exists(): raise RuntimeError("Snapshot nao encontrado: "+str(SNAPSHOT_PATH))
    snapshot=json.loads(SNAPSHOT_PATH.read_text(encoding="utf-8"))
    products=snapshot.get("products") or []
    expected=int(snapshot.get("expectedCount") or 0)
    if expected!=len(products) or expected<1: raise RuntimeError("Snapshot incompleto ou inconsistente")
    if int(snapshot.get("schemaVersion") or 0) < 2: raise RuntimeError("Snapshot antigo: gere novamente antes de publicar")
    required=("costOrigin","daysSinceSale","daysSincePurchase","lastSaleDate","lastPurchaseDate","openPurchaseQty","revenue30","revenue60","revenue90")
    missing={key:sum(1 for p in products if key not in p) for key in required}
    if any(missing.values()): raise RuntimeError("Snapshot sem campos obrigatorios: "+str(missing))
    print("  Cobertura dos campos novos: OK ("+str(expected)+" produtos)")
    codes=[str(p.get("code") or "") for p in products]
    if any(not c for c in codes) or len(codes)!=len(set(codes)): raise RuntimeError("Codigos vazios ou duplicados no snapshot")
    run_id="agm_"+dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    post({"action":"start","schemaVersion":snapshot["schemaVersion"],"runId":run_id,"expectedCount":expected,"checksum":snapshot["checksum"],"summary":snapshot.get("summary",{})})
    for start in range(0,expected,CHUNK_SIZE):
        result=post({"action":"chunk","runId":run_id,"products":products[start:start+CHUNK_SIZE]})
        print("  Enviados: "+str(result.get("receivedCount"))+"/"+str(expected))
    result=post({"action":"finalize","runId":run_id})
    print("  Base principal publicada: "+str(result.get("publishedCount"))+" produtos")
    receipts=snapshot.get("receipts") or []
    # O processamento de recebimentos cruza cada item com pedidos em aberto.
    # Lotes de 100 notas ultrapassavam o limite de tempo do Worker, mesmo com
    # a base principal ja publicada. Cinco notas por chamada mantem cada
    # operacao curta e permite acompanhar o progresso real.
    receipt_batch_size=5
    for start in range(0,len(receipts),receipt_batch_size):
        batch=receipts[start:start+receipt_batch_size]
        post({"action":"receipts","receipts":batch})
        print("  Entradas processadas: "+str(min(start+len(batch),len(receipts)))+"/"+str(len(receipts)))
    if receipts:
        print("  Notas de entrada interpretadas: "+str(len(receipts)))
    daily_sales=snapshot.get("dailySales") or []
    # Mesma logica de lotes das entradas: manda em pedacos pra cada chamada
    # ao Worker ficar curta. Vendas diarias sao so leitura+upsert (sem cruzar
    # pedidos em aberto como as entradas fazem), entao um lote maior cabe sem
    # estourar o tempo do Worker.
    daily_sales_batch_size=500
    for start in range(0,len(daily_sales),daily_sales_batch_size):
        batch=daily_sales[start:start+daily_sales_batch_size]
        post({"action":"dailySales","entries":batch})
        print("  Vendas diarias enviadas: "+str(min(start+len(batch),len(daily_sales)))+"/"+str(len(daily_sales)))
    if daily_sales:
        print("  Dias com venda registrados: "+str(len(daily_sales)))
    daily_sales_by_client=snapshot.get("dailySalesByClient") or []
    # Mesmo padrao de lotes de dailySales acima (22/09/2026, drill-down por
    # cliente no grafico financeiro) -- so leitura+upsert, sem cruzar pedidos
    # em aberto, entao cabe um lote grande sem estourar o tempo do Worker.
    daily_sales_by_client_batch_size=500
    for start in range(0,len(daily_sales_by_client),daily_sales_by_client_batch_size):
        batch=daily_sales_by_client[start:start+daily_sales_by_client_batch_size]
        post({"action":"dailySalesByClient","entries":batch})
        print("  Vendas diarias por cliente enviadas: "+str(min(start+len(batch),len(daily_sales_by_client)))+"/"+str(len(daily_sales_by_client)))
    if daily_sales_by_client:
        print("  Combinacoes dia+cliente registradas: "+str(len(daily_sales_by_client)))
    print("OK - ATLAS atualizado com "+str(result.get("publishedCount"))+" produtos em "+str(result.get("updatedAt")))

if __name__=="__main__": main()
