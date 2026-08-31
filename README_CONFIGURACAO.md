# Atlas rodando na nuvem — passo a passo (Jonas)

Isso aqui faz o botão **"Atualizar agora"** do Atlas funcionar de
**qualquer lugar** (celular incluído), sem precisar de nenhum computador
ligado. Quem faz o trabalho pesado passa a ser o **GitHub** (de graça, na
sua cota de uso normal) em vez do seu Mac.

Nenhum dos passos abaixo exige que você me passe uma senha ou chave —
você sempre cola os valores direto no GitHub ou no Cloudflare, nunca aqui
no chat.

## O que já está pronto nesta pasta

- Os mesmos scripts Python que já rodam no seu Mac (`sync_bling.py`,
  `gerar_sugestao_compras.py`, `publicar_agm_site.py`), só que ajustados
  pra rodar sozinhos, sem Mac.
- `run_cloud_sync.py` — o "regente" que roda os três na ordem certa e
  avisa o Atlas do progresso (é o que faz a barra de progresso continuar
  funcionando igual).
- `.github/workflows/atlas-sync.yml` — a receita que diz ao GitHub *como*
  e *quando* rodar isso.
- `worker-changes/` — duas mudanças pequenas no código do site do Atlas
  (não nos scripts) pra ele saber chamar o GitHub.

## Passo 1 — Criar o repositório no GitHub

1. Se você não tem conta no [github.com](https://github.com), crie uma
   (gratuita).
2. Clique em **New repository**.
3. Nome sugerido: `agm-atlas-sync`. Marque como **Private** (privado).
4. Não marque nenhuma opção de "adicionar README" — deixe vazio.
5. Depois de criado, clique em **uploading an existing file** (ou
   "Add file" → "Upload files") e arraste **todos os arquivos desta
   pasta** (menos a pasta `worker-changes/`, essa não vai pro GitHub —
   ela é pro código do Atlas, ver Passo 4). Inclua a pasta `.github/`
   também — pode arrastar tudo junto, o GitHub organiza sozinho.
6. Clique em **Commit changes**.

## Passo 2 — Configurar os "Secrets" do repositório

No repositório que você acabou de criar: **Settings** → **Secrets and
variables** → **Actions** → **New repository secret**.

Crie um secret pra cada linha abaixo (o nome à esquerda, exatamente
assim; o valor é o que você já tem hoje no arquivo `.env` do Mac —
pode abrir esse arquivo lá e copiar de lá pra cá):

| Nome do Secret | De onde vem o valor |
|---|---|
| `BLING_CLIENT_ID` | mesmo do `.env` do Mac |
| `BLING_CLIENT_SECRET` | mesmo do `.env` do Mac |
| `BLING_BASE_URL` | mesmo do `.env` do Mac (ex: `https://api.bling.com.br/Api/v3`) |
| `AGM_SITE_SYNC_KEY` | mesmo do `.env` do Mac |
| `CF_ACCESS_CLIENT_ID` | mesmo do `.env` do Mac |
| `CF_ACCESS_CLIENT_SECRET` | mesmo do `.env` do Mac |
| `AGM_SITE_URL` | `https://atlas-agm.comprasagmautopartes.workers.dev` |

## Passo 3 — Enviar o token do Bling pra nuvem (só uma vez)

Isso deixa o GitHub apto a renovar sozinho o acesso ao Bling, sem você
precisar logar de novo. Roda **no Mac**, dentro da pasta `Bling_Sync`:

```
source .venv/bin/activate
python seed_bling_token_to_cloud.py
```

(Esse comando só funciona depois que o Passo 4 abaixo — a mudança no
Worker — já tiver sido publicada, porque ele envia o token pro endpoint
novo que ainda não existe no site.)

## Passo 4 — Aplicar as duas mudanças no código do Atlas

Isso é código do próprio site do Atlas (não desses scripts). Se você
tiver o código do Atlas aberto num editor, ou se preferir, me chama numa
próxima sessão que eu aplico essas duas mudanças diretamente e te aviso
exatamente o comando de publicar — só listando aqui pra você saber o que
são:

1. **Arquivo novo**: `app/api/bling-token/route.ts` — conteúdo em
   `worker-changes/app-api-bling-token-route.ts` (só copiar).
2. **Arquivo substituído**: `app/api/refresh/route.ts` — conteúdo em
   `worker-changes/app-api-refresh-route.ts` (substitui o que já existe;
   a única mudança real é uma função nova que avisa o GitHub quando
   alguém clica em "Atualizar agora").

Depois, no site do Atlas (`wrangler`), mais 4 secrets (esses ficam no
Cloudflare, não no GitHub):

```
wrangler secret put GITHUB_TOKEN
wrangler secret put GITHUB_REPO
wrangler secret put GITHUB_WORKFLOW_FILE
wrangler secret put GITHUB_REF
```

Valores:
- `GITHUB_TOKEN` — um **Personal Access Token** do GitHub, criado em
  *Settings → Developer settings → Personal access tokens → Fine-grained
  tokens*, com acesso **só ao repositório `agm-atlas-sync`** e permissão
  **Actions: Read and write** (nada além disso — quanto mais restrito,
  melhor).
- `GITHUB_REPO` — `seu-usuario/agm-atlas-sync`
- `GITHUB_WORKFLOW_FILE` — `atlas-sync.yml`
- `GITHUB_REF` — `main`

Depois é só rodar o deploy do Atlas do jeito que você já faz hoje
(`ATUALIZAR_ATLAS_CLOUDFLARE.ps1`, ou os comandos `wrangler` de sempre).

## Passo 5 — Testar

Clica em "Atualizar agora" no Atlas — pode ser do celular, de qualquer
navegador, sem precisar do Mac ligado. Se tudo estiver certo, o GitHub
assume a sincronização sozinho e a barra de progresso do Atlas se move
normalmente.

## Importante

- **O Mac continua funcionando em paralelo, se você quiser manter os
  dois** — nada nisso desliga o agente local; é só mais um jeito de
  disparar a sincronização, e o GitHub e o Mac nunca vão rodar a mesma
  sincronização ao mesmo tempo (o Atlas só deixa um job ativo por vez).
- **Custo**: 100% dentro da cota gratuita do GitHub pro seu volume de
  uso (a sincronização "rápida" leva minutos, não horas, depois que o
  primeiro dia de carga completa já passou — que já passou, no seu
  caso).
- Se o `wrangler deploy` reclamar de `GITHUB_TOKEN` "não existir no tipo
  Env", roda `wrangler types` uma vez pra ele gerar os tipos de novo — é
  normal depois de adicionar secrets novos, não é erro de verdade.
- Guarda o Personal Access Token do GitHub com o mesmo cuidado que uma
  senha — quem tiver ele consegue disparar sincronizações nesse
  repositório (só isso, nada além, por causa do escopo restrito acima).
