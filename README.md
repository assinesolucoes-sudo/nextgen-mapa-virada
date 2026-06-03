# Mapa da Virada — Mentoria NextGen

Painel da mentorada (Paula) + **modo administrativo** pra Elaine acompanhar, com **persistência real no Google Sheets**. Construído pra Frente 1 (a virada de Pleno → Sênior na trilha de Atração e Seleção) e o horizonte de liderança.

Mesmo padrão de hospedagem do formulário diagnóstico — Streamlit Community Cloud, grátis.

## O que tem aqui

- `app.py` — o app: painel da mentorada + modo `?admin=`
- `requirements.txt` — dependências (Streamlit, gspread, google-auth, pandas)
- `.streamlit/config.toml` — paleta visual NextGen (roxos)
- `.streamlit/secrets.toml.example` — modelo dos segredos (credencial do Google + admin_key)

## Como funciona a persistência (o que faz "ser guardado")

O `session_state` do Streamlit é efêmero — quem guarda é o **Google Sheets**. A planilha ganha duas abas, criadas sozinhas na primeira execução:

- **estado** — uma linha por mentorada com o estado completo (JSON). Cada *Salvar progresso* atualiza essa linha.
- **marcos** — uma linha por marco salvo (data + índice de senioridade + índice de liderança). É a trajetória.

O modo admin lê **da mesma planilha**. Por isso, à medida que a Paula salva, você vê — encontro após encontro. A gravação acontece no botão *Salvar progresso* e automaticamente ao guardar uma evidência ou um marco (não a cada tecla, pra não estourar o limite da API).

## Como rodar localmente (pra testar antes de publicar)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre em `http://localhost:8501`. Sem os segredos configurados, o app funciona pra navegar, mas avisa que **não** está guardando — útil pra ver a interface antes de ligar o Sheets.

## Como publicar no Streamlit Community Cloud

**Passo 1 — Criar repositório no GitHub.** Crie um repo novo (sugestão: `nextgen-mapa-virada`) na conta `assinesolucoes-sudo` e suba estes arquivos, mantendo a pasta `.streamlit`.

**Passo 2 — Ligar a persistência no Google (uma vez).**
1. No Google Cloud, crie/selecione um projeto e ative as APIs **Google Sheets** e **Google Drive**.
2. Crie uma **conta de serviço**, gere uma **chave JSON** e baixe.
3. Crie uma Google Sheet vazia e **compartilhe** com o e-mail da conta de serviço (`...@...iam.gserviceaccount.com`) como **Editor**.

**Passo 3 — Publicar.** No Streamlit Community Cloud, faça *New app* apontando pra este repo e pro `app.py`. Escolha o **subdomínio** que quiser (ex.: `mapa-virada` → `mapa-virada.streamlit.app`).

**Passo 4 — Colar os segredos.** Em *Settings → Secrets* do app, cole o conteúdo de `.streamlit/secrets.toml.example` preenchido com os campos do JSON, a URL da planilha e o `admin_key`. Mantenha as quebras de linha do `private_key` como `\n`. **Nunca** suba o segredo real pro GitHub.

> Importante: no Community Cloud os **secrets são por app**, não por repositório. A credencial do Sheets fica colada só neste app.

## Os dois links

- **Mentorada (Paula):** `https://mapa-virada.streamlit.app/?m=Paula`
- **Você (admin):** `https://mapa-virada.streamlit.app/?admin=elaine`

(Troque `mapa-virada` pelo subdomínio que você escolher. O `?admin=` usa o valor que você definir em `admin_key` nos secrets.)

## Observações

- A proteção do `?admin=` é leve (igual à do formulário atual). Dá pra somar uma senha depois, se quiser.
- A seção de liderança usa hoje os eixos do Guia NextGen + o DISC da Paula. Quando vier a régua formal de Coordenador da TBC, troca-se pelos critérios reais.
