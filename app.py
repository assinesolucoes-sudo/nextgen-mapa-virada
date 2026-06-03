# -*- coding: utf-8 -*-
"""
Mapa da Virada — Mentoria NextGen (Streamlit + Google Sheets)
Painel da mentorada + modo admin (?admin=...). Persistência real no Google Sheets.
Identidade visual alinhada ao formulário diagnóstico NextGen.
"""
import json
import datetime as dt
import streamlit as st
import pandas as pd

# ============================================================ CONSTANTES (Book + DISC)
LEVELS = ["Noções", "Conhecimento", "Domínio", "Expert"]
STEPS = ["Preciso de validação", "Faço, revisam", "Faço sozinha", "Faço e oriento outros"]
LSCALE = ["Em desenvolvimento", "Emergente", "Consolidada", "Referência"]
NA = "—"

COMPS = {
    "analitica": {
        "name": "Capacidade Analítica",
        "desc": "Produzir informação a partir de dados de várias fontes pra apoiar a decisão.",
        "behav": [
            ["Compreender os indicadores da área e de onde vêm os dados.",
             "Organizar dados de várias fontes de forma racional.",
             "Retroalimentar sistemas e relatórios."],
            ["Analisar dados identificando padrões e sinalizando desvios.",
             "Buscar proativamente informação em outras fontes.",
             "Conduzir análises de baixa complexidade."],
            ["Análises complexas (múltiplas fontes/áreas) apoiando a decisão.",
             "Identificar inconsistências e propor correção de rota.",
             "Definir formatos/padrões e sintetizar riscos e oportunidades."],
            ["Criar novos indicadores a partir de histórico, tendências e objetivos do negócio.",
             "Decidir a partir de dados complexos, com clareza das implicações.",
             "Estabelecer os parâmetros de decisão da área."],
        ],
    },
    "influencia": {
        "name": "Influência e Persuasão",
        "desc": "Compreender o ambiente do cliente/área e influenciar decisões com ética.",
        "behav": [
            ["Mapear influenciadores e o ambiente do cliente.",
             "Compreender diferenciais pra argumentar a favor.",
             "Manter-se atualizada sobre o segmento."],
            ["Reconhecer forças, fraquezas e o mapa de poder.",
             "Aplicar técnicas de persuasão pra fechar acordos.",
             "Estabelecer confiança que influencia a decisão."],
            ["Criar e fortalecer vínculos de confiança com stakeholders.",
             "Antecipar necessidades de toda a cadeia.",
             "Influenciar OUTRAS áreas da TOTVS com melhorias e produtos."],
            ["Criar estratégias pra ampliar território, influenciando outros TOTVERS.",
             "Ser reconhecida como referência no tema.",
             "Atuar como voz ativa no futuro do segmento."],
        ],
        "note": ("Do seu DISC: o I (influência) é forte e estável. O ponto a observar é a abertura "
                 "a opiniões divergentes — chegar a Domínio aqui passa por influenciar outras áreas, "
                 "e flexibilidade ajuda."),
    },
}

RESP_BASE = [
    "Alinhar a demanda da vaga com o gestor (perfil, requisitos)",
    "Triar e abordar candidatos (ATS, LinkedIn)",
    "Entrevistar candidatos (cultura e competências)",
    "Garantir o SLA do processo seletivo",
    "Conduzir projetos de Employer Branding",
    "Gerar e analisar indicadores estratégicos da área",
    "Construir testes/avaliações junto às áreas",
    "Dar devolutivas e apresentar candidatos aos gestores",
    "Mapear mercado e construir pipeline de talentos",
]

REQ = [
    ("tempo", "3+ anos em Recrutamento e Seleção", "Tempo mínimo do nível III", False),
    ("gerencia", "Atuação em vagas de nível Gerência", "II vai até Coordenação; III inclui Gerência", False),
    ("pos", "Pós-graduação completa", "Gestão de Pessoas ou áreas correlatas", False),
    ("agil", "Metodologias ágeis", "Listado como desejável no III", True),
]
REQ_HARD = ["tempo", "gerencia", "pos"]

LEAD = [
    ("infl", "Influência", "Seu I é estável e o D subiu — base forte aqui."),
    ("fb", "Dar feedback que desenvolve", "Eixo a observar; menos evidente no DISC."),
    ("com", "Comunicação e narrativa", "Alta influência favorece; estruture pra liderança."),
    ("prot", "Postura protagonista", "D em alta empurra pra assumir a frente."),
    ("dec", "Tomada de decisão", "DISC: dificuldade em 2024/25, mais firme em 2026."),
    ("rel", "Gestão de pessoas e relacionamentos", "Ponto de atenção: abertura a divergência / flexibilidade."),
]
LEAD_KEYS = [k for k, _, _ in LEAD]

ARC = [("Analista II", "você hoje"), ("Analista III", "a virada (Frente 1)"),
       ("Especialista / Coordenador", "fork: técnica ou liderança"), ("Diretor", "horizonte")]
DISC = [("2024", "IS"), ("2025", "IS"), ("2026", "ID")]

# ============================================================ ARMAZENAMENTO
HEAD_ESTADO = ["mentee", "updated_at", "json"]
HEAD_MARCOS = ["mentee", "data", "senioridade", "lideranca", "ts"]


@st.cache_resource(show_spinner=False)
def _open_sheet():
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds).open_by_url(st.secrets["app"]["sheet_url"])


def storage_ready():
    return ("gcp_service_account" in st.secrets) and ("app" in st.secrets) and \
           ("sheet_url" in st.secrets.get("app", {}))


def _ws(sh, title, header):
    try:
        return sh.worksheet(title)
    except Exception:
        ws = sh.add_worksheet(title=title, rows=200, cols=max(6, len(header)))
        ws.append_row(header)
        return ws


def save_state(mentee, state):
    ws = _ws(_open_sheet(), "estado", HEAD_ESTADO)
    payload = json.dumps(state, ensure_ascii=False)
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    try:
        cell = ws.find(mentee, in_column=1)
    except Exception:
        cell = None
    if cell:
        ws.update(f"B{cell.row}:C{cell.row}", [[now, payload]])
    else:
        ws.append_row([mentee, now, payload])


def load_state(mentee):
    ws = _ws(_open_sheet(), "estado", HEAD_ESTADO)
    for rec in ws.get_all_records():
        if str(rec.get("mentee", "")).strip() == mentee:
            try:
                return json.loads(rec.get("json") or "{}")
            except Exception:
                return {}
    return {}


def append_marco(mentee, sen, lid):
    ws = _ws(_open_sheet(), "marcos", HEAD_MARCOS)
    ws.append_row([mentee, dt.date.today().strftime("%d/%m"), sen, lid,
                   dt.datetime.now().isoformat(timespec="seconds")])


def load_marcos(mentee):
    ws = _ws(_open_sheet(), "marcos", HEAD_MARCOS)
    return [{"d": r.get("data"), "sen": int(r.get("senioridade") or 0), "lid": int(r.get("lideranca") or 0)}
            for r in ws.get_all_records() if str(r.get("mentee", "")).strip() == mentee]


def list_mentees():
    ws = _ws(_open_sheet(), "estado", HEAD_ESTADO)
    return [str(r.get("mentee")) for r in ws.get_all_records() if r.get("mentee")]


# ============================================================ CÁLCULOS
def _lvl(label): return ([NA] + LEVELS).index(label) if label in ([NA] + LEVELS) else 0
def _stp(label): return ([NA] + STEPS).index(label) if label in ([NA] + STEPS) else 0
def _lsc(label): return ([NA] + LSCALE).index(label) if label in ([NA] + LSCALE) else 0
def resps(): return RESP_BASE + st.session_state.get("custom", [])


def build_state():
    ss = st.session_state
    rs = resps()
    return {
        "comp": {"analitica": _lvl(ss.get("comp_analitica", NA)),
                 "influencia": _lvl(ss.get("comp_influencia", NA))},
        "aut": [{"n": _stp(ss.get(f"aut_{i}", NA)), "ev": ss.get(f"autev_{i}", "")} for i in range(len(rs))],
        "req": {k: bool(ss.get(f"req_{k}", False)) for k, _, _, _ in REQ},
        "lead": {k: _lsc(ss.get(f"lead_{k}", NA)) for k in LEAD_KEYS},
        "ev": ss.get("ev", []),
        "custom": ss.get("custom", []),
    }


def senioridade_idx(s):
    c = (s["comp"]["analitica"] + s["comp"]["influencia"]) / 8.0
    rs = max(1, len(s["aut"]))
    aut_score = sum(a["n"] for a in s["aut"]) / (rs * 4.0)
    aut_sr = sum(1 for a in s["aut"] if a["n"] >= 3)
    req_hard = sum(1 for k in REQ_HARD if s["req"].get(k))
    req_score = (req_hard + (0.5 if s["req"].get("agil") else 0)) / 3.5
    total = round((c * 0.35 + aut_score * 0.40 + req_score * 0.25) * 100)
    return {"total": total, "aut_sr": aut_sr, "req_hard": req_hard, "rs": len(s["aut"])}


def lideranca_idx(s):
    return round((sum(s["lead"].get(k, 0) for k in LEAD_KEYS) / (len(LEAD_KEYS) * 4.0)) * 100)


# ============================================================ ESTILO (identidade NextGen)
def inject_css():
    st.markdown("""
    <style>
    .stApp { background: linear-gradient(180deg, #faf9fc 0%, #f3eefa 100%); }
    .block-container { padding-top: 2rem !important; padding-bottom: 4rem !important; max-width: 840px !important; }
    #MainMenu, footer { visibility: hidden; }
    .nextgen-header { background: linear-gradient(135deg,#4A2D85 0%,#5D3A9B 40%,#8B5FBF 100%);
      padding: 44px 40px; border-radius: 24px; margin-bottom: 28px;
      box-shadow: 0 12px 40px rgba(93,58,155,.20); position: relative; overflow: hidden; }
    .nextgen-header::before { content:''; position:absolute; top:-50%; right:-20%; width:60%; height:200%;
      background: radial-gradient(circle, rgba(255,255,255,.18) 0%, transparent 70%); }
    .nextgen-brand { color: rgba(255,255,255,.85); font-size: 12px; font-weight: 700; letter-spacing: 3px;
      text-transform: uppercase; margin-bottom: 12px; position: relative; }
    .nextgen-title { color:#fff; font-size: 34px; font-weight: 700; margin: 0 0 12px 0; line-height: 1.12;
      position: relative; letter-spacing: -.5px; }
    .nextgen-subtitle { color: rgba(255,255,255,.95); font-size: 15px; line-height: 1.6; margin: 0;
      position: relative; max-width: 640px; }
    .block-strip { background: linear-gradient(135deg,#5D3A9B 0%,#8B5FBF 100%); color:#fff;
      padding: 22px 26px; border-radius: 16px; margin: 36px 0 8px 0; box-shadow: 0 6px 24px rgba(93,58,155,.15);
      position: relative; overflow: hidden; }
    .block-strip::before { content:''; position:absolute; top:0; right:0; width:200px; height:100%;
      background: radial-gradient(circle at right, rgba(255,255,255,.15) 0%, transparent 70%); }
    .block-strip-label { color: rgba(255,255,255,.85); font-size: 11px; font-weight: 700; letter-spacing: 2.5px;
      text-transform: uppercase; margin-bottom: 8px; position: relative; }
    .block-strip-title { color:#fff; font-size: 22px; font-weight: 700; margin: 0 0 6px 0; position: relative;
      letter-spacing: -.3px; }
    .block-strip-desc { color: rgba(255,255,255,.92); font-size: 13.5px; line-height: 1.6; margin: 0;
      position: relative; max-width: 95%; }
    div[data-testid="stVerticalBlockBorderWrapper"] { background:#fff !important; border-radius: 16px !important;
      padding: 16px 22px !important; box-shadow: 0 4px 18px rgba(93,58,155,.07) !important;
      border: 1px solid #f0e8f9 !important; }
    div[data-testid="stRadio"] label { padding: 9px 14px !important; background:#faf9fc !important;
      border-radius: 12px !important; border: 1.5px solid transparent !important; transition: all .2s ease !important; }
    div[data-testid="stRadio"] label:hover { background:#f0e8f9 !important; border-color:#d4c3e8 !important; }
    div[data-testid="stRadio"] label p { font-size: 14px !important; color:#2c2c2a !important; }
    div[data-testid="stCheckbox"] { background:#faf9fc !important; padding: 11px 15px !important;
      border-radius: 12px !important; margin: 6px 0 !important; border: 1.5px solid transparent !important; }
    div[data-testid="stCheckbox"]:hover { background:#f0e8f9 !important; border-color:#d4c3e8 !important; }
    .stTextInput input, .stTextArea textarea { background:#faf9fc !important; border:1.5px solid #E8DDF5 !important;
      border-radius: 12px !important; color:#2c2c2a !important; }
    .stTextInput input:focus, .stTextArea textarea:focus { border-color:#8B5FBF !important;
      box-shadow: 0 0 0 4px rgba(139,95,191,.12) !important; }
    .stButton button { background: linear-gradient(135deg,#4A2D85 0%,#5D3A9B 50%,#8B5FBF 100%) !important;
      color:#fff !important; border:none !important; border-radius: 14px !important; padding: 12px 24px !important;
      font-weight: 700 !important; transition: all .3s ease !important; box-shadow: 0 6px 20px rgba(93,58,155,.25) !important; }
    .stButton button:hover { transform: translateY(-2px) !important; box-shadow: 0 10px 28px rgba(93,58,155,.35) !important; }
    .metric-cap { color:#6c6580; font-size: 12.5px; margin-top: 2px; }
    .pill { display:inline-block; background:#efe9fb; color:#6d28d9; border-radius:99px; padding:2px 10px;
      font-size:12px; font-weight:600; margin-right:6px; }
    .arc { display:flex; gap:6px; margin: 6px 0 14px; flex-wrap:wrap; }
    .arc div { flex:1; min-width:120px; border-radius:10px; padding:11px; text-align:center; background:#f3f1f8;
      border:1px solid #e6e2ee; font-size:11px; color:#6c6580; }
    .arc b { display:block; font-size:13px; color:#1b1726; margin-bottom:2px; }
    .arc .now { background:#2b2540; color:#fff; } .arc .now b { color:#fff; }
    .arc .next { background:#f6edd4; border-color:#e6d49a; }
    .arc .fork { background:#e0f1f2; border-color:#bfe2e4; }
    .leitura { background: linear-gradient(135deg,#2b2540 0%,#3a2d63 100%); color:#efeaf9; border-radius:16px;
      padding: 22px 24px; }
    .leitura h3 { color:#fff; font-size:18px; margin:0 0 10px; }
    </style>
    """, unsafe_allow_html=True)


def block_strip(label, title, desc):
    st.markdown(f"<div class='block-strip'><div class='block-strip-label'>{label}</div>"
                f"<div class='block-strip-title'>{title}</div>"
                f"<div class='block-strip-desc'>{desc}</div></div>", unsafe_allow_html=True)


def _evolucao_chart(marcos):
    if not marcos:
        st.caption("Nenhum marco salvo ainda. Salve o primeiro abaixo pra fixar seu ponto de partida.")
        return
    labels = [f"{i+1} · {m['d']}" for i, m in enumerate(marcos)]
    rows = []
    for i, m in enumerate(marcos):
        rows.append({"Marco": labels[i], "Pontos": m["sen"], "Indicador": "Senioridade"})
        rows.append({"Marco": labels[i], "Pontos": m.get("lid", 0), "Indicador": "Liderança"})
    df = pd.DataFrame(rows)
    try:
        import altair as alt
        ch = (alt.Chart(df).mark_line(point=True, strokeWidth=3)
              .encode(x=alt.X("Marco:N", sort=labels, title=None),
                      y=alt.Y("Pontos:Q", title="Pontos (0–100)", scale=alt.Scale(domain=[0, 100])),
                      color=alt.Color("Indicador:N", title=None,
                                      scale=alt.Scale(domain=["Senioridade", "Liderança"],
                                                      range=["#5D3A9B", "#B47FE0"])))
              .properties(height=240))
        st.altair_chart(ch, use_container_width=True)
    except Exception:
        st.line_chart(df, x="Marco", y="Pontos", color="Indicador")
    if len(marcos) > 1:
        d = marcos[-1]["sen"] - marcos[0]["sen"]
        st.caption(f"{len(marcos)} marcos. Senioridade: {d:+d} pts desde o primeiro.")


# ============================================================ HIDRATAÇÃO
def hydrate(mentee):
    if st.session_state.get("_loaded") == mentee:
        return
    data = load_state(mentee) if storage_ready() else {}
    comp = data.get("comp", {})
    st.session_state["comp_analitica"] = LEVELS[comp.get("analitica", 0) - 1] if comp.get("analitica") else NA
    st.session_state["comp_influencia"] = LEVELS[comp.get("influencia", 0) - 1] if comp.get("influencia") else NA
    st.session_state["custom"] = data.get("custom", [])
    aut = data.get("aut", [])
    for i in range(len(RESP_BASE) + len(st.session_state["custom"])):
        a = aut[i] if i < len(aut) else {"n": 0, "ev": ""}
        st.session_state[f"aut_{i}"] = STEPS[a["n"] - 1] if a.get("n") else NA
        st.session_state[f"autev_{i}"] = a.get("ev", "")
    req = data.get("req", {})
    for k, _, _, _ in REQ:
        st.session_state[f"req_{k}"] = bool(req.get(k, False))
    lead = data.get("lead", {})
    for k in LEAD_KEYS:
        st.session_state[f"lead_{k}"] = LSCALE[lead.get(k, 0) - 1] if lead.get(k) else NA
    st.session_state["ev"] = data.get("ev", [])
    st.session_state["marcos"] = load_marcos(mentee) if storage_ready() else []
    st.session_state["_loaded"] = mentee


# ============================================================ PÁGINA MENTORADA
def page_mentee(mentee):
    hydrate(mentee)
    inject_css()
    st.markdown(f"<div class='nextgen-header'><div class='nextgen-brand'>Mentoria NextGen · Frente 1 + Horizonte</div>"
                f"<div class='nextgen-title'>Mapa da Virada</div>"
                f"<div class='nextgen-subtitle'>{mentee} · Pleno (Analista II) &rarr; Sênior (Analista III), "
                f"trilha de Atração e Seleção — e o passo seguinte rumo à liderança.</div></div>",
                unsafe_allow_html=True)

    if not storage_ready():
        st.warning("Persistência não configurada — o que você preencher não será guardado.", icon="⚠️")

    s = build_state()
    k = senioridade_idx(s)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Prontidão p/ nível III", f"{k['total']}%")
    c2.metric("Autonomia sênior", f"{k['aut_sr']}/{k['rs']}")
    c3.metric("Requisitos formais", f"{k['req_hard']}/3")
    c4.metric("Prontidão p/ liderança", f"{lideranca_idx(s)}%")
    st.markdown("<div class='metric-cap'>Retrato de hoje, não nota. Atualizam conforme você responde.</div>",
                unsafe_allow_html=True)
    with st.expander("Como esses números são calculados?"):
        st.markdown("A **prontidão pro nível III** soma três partes da régua: a autonomia nas responsabilidades "
                    "(peso maior), o nível nas duas competências do Book e os requisitos formais. Quanto mais "
                    "itens em nível sênior, mais alto. A **prontidão pra liderança** é a média do que você marcar "
                    "na seção Horizonte. Nada vira nota — é um espelho do momento, que muda a cada resposta sua.")

    # ---- EVOLUÇÃO (perto do topo) ----
    block_strip("SUA EVOLUÇÃO", "Linha do tempo",
                "Salve um marco a cada encontro (ou quando algo mudar). Os pontos vão se somando — nada se apaga.")
    _evolucao_chart(st.session_state.get("marcos", []))
    cma, cmb = st.columns([1, 2])
    if cma.button("Salvar marco de hoje"):
        s2 = build_state()
        sen, lid = senioridade_idx(s2)["total"], lideranca_idx(s2)
        st.session_state["marcos"].append({"d": dt.date.today().strftime("%d/%m"), "sen": sen, "lid": lid})
        if storage_ready():
            save_state(mentee, s2)
            append_marco(mentee, sen, lid)
        st.rerun()
    with st.expander("Como funciona a linha do tempo?"):
        st.markdown("Cada vez que você clica em **Salvar marco de hoje**, eu tiro uma foto dos seus dois "
                    "indicadores naquele dia (senioridade e liderança) e guardo com a data. Conforme você "
                    "revisita e ajusta ao longo dos meses, salvando novos marcos, eles vão **se somando** no "
                    "gráfico — você enxerga os pontos subindo. **Você nunca apaga nada**: a jornada inteira fica registrada.")

    # ---- COMPETÊNCIAS ----
    block_strip("RÉGUA · PARTE 1 DE 3", "As duas competências do Book",
                "São as duas competências específicas de Relações Humanas no Book de Carreira. A régua completa "
                "tem três partes: estas competências, a autonomia (parte 2) e os requisitos formais (parte 3).")
    with st.expander("De onde vêm essas duas competências?"):
        st.markdown("Direto do **Book de Carreira de Relações Humanas** da TOTVS. Lá, pra essa carreira, estão "
                    "definidas exatamente duas competências específicas: **Capacidade Analítica** e **Influência "
                    "e Persuasão**. Os comportamentos de cada nível abaixo também são os do próprio Book.")
    for key in ("analitica", "influencia"):
        c = COMPS[key]
        st.markdown(f"**{c['name']}** · {c['desc']}")
        st.radio(c["name"], [NA] + LEVELS, key=f"comp_{key}", horizontal=True, label_visibility="collapsed")
        sel = _lvl(st.session_state[f"comp_{key}"])
        if sel:
            with st.container(border=True):
                st.markdown("\n".join(["• " + b for b in c["behav"][sel - 1]]))
            if sel >= 3:
                st.success("Território sênior — conta a favor da virada.", icon="✅")
        if key == "influencia" and c.get("note"):
            st.info(c["note"], icon="🧭")
        st.divider()

    # ---- AUTONOMIA ----
    block_strip("RÉGUA · PARTE 2 DE 3", "Autonomia — onde mora a virada",
                "As responsabilidades são as MESMAS no II e no III; muda o nível de supervisão (moderado → baixo). "
                "Marque como você opera hoje. De 'faço sozinha' pra cima = nível III.")
    with st.expander("Por que a autonomia mede a virada?"):
        st.markdown("No Book, a lista de responsabilidades de um Analista II e de um III é igual. O que separa os "
                    "dois é **o quanto você faz sem supervisão**. Por isso, mais do que 'o que você faz', o que "
                    "evidencia a senioridade é **com que autonomia** você faz cada coisa.")
    rs = resps()
    for i, label in enumerate(rs):
        if f"aut_{i}" not in st.session_state:
            st.session_state[f"aut_{i}"] = NA
            st.session_state[f"autev_{i}"] = ""
        cols = st.columns([3, 1])
        cols[0].markdown(f"**{label}**")
        n = _stp(st.session_state[f"aut_{i}"])
        if n >= 3:
            cols[1].markdown("<span class='pill'>nível III</span>", unsafe_allow_html=True)
        elif n:
            cols[1].caption("nível II")
        st.radio(label, [NA] + STEPS, key=f"aut_{i}", horizontal=True, label_visibility="collapsed")
        st.text_input("Evidência rápida", key=f"autev_{i}", placeholder="um exemplo recente em que operou assim",
                      label_visibility="collapsed")
        if i >= len(RESP_BASE):
            if st.button("remover esta responsabilidade", key=f"delresp_{i}"):
                st.session_state["custom"].pop(i - len(RESP_BASE))
                st.session_state["_loaded"] = None
                st.rerun()
        st.divider()
    with st.form("add_resp", clear_on_submit=True):
        novo = st.text_input("Adicionar uma responsabilidade sua que não está na lista")
        if st.form_submit_button("Adicionar responsabilidade") and novo.strip():
            st.session_state["custom"].append(novo.strip())
            st.rerun()

    # ---- REQUISITOS ----
    block_strip("RÉGUA · PARTE 3 DE 3", "Requisitos formais",
                "Os critérios 'duros' do III — sem subjetividade. Às vezes a promoção trava aqui, independente da entrega.")
    for kk, b, sub, opt in REQ:
        st.checkbox(b + (" · desejável" if opt else ""), key=f"req_{kk}", help=sub)

    # ---- COFRE ----
    block_strip("EVIDÊNCIAS", "Cofre de evidências",
                "As entregas que provam sua senioridade. Vá guardando ao longo do tempo — comece pelas 3 que "
                "você considera nível sênior. Ele acumula e fica como seu acervo.")
    with st.form("add_ev", clear_on_submit=True):
        t = st.text_input("Título da entrega")
        tag = st.selectbox("Marcar como", ["Capacidade Analítica", "Influência e Persuasão",
                                           "Autonomia / processo", "Liderança"])
        d = st.text_input("2 linhas: o que era + o impacto")
        prova = st.text_input("Onde está a prova (link/local) — opcional")
        if st.form_submit_button("Guardar evidência") and t.strip():
            st.session_state["ev"].append({"t": t.strip(), "tag": tag, "d": d.strip(),
                                           "p": prova.strip(), "date": dt.date.today().strftime("%d/%m/%Y")})
            if storage_ready():
                save_state(mentee, build_state())
            st.rerun()
    if st.session_state["ev"]:
        for j, e in enumerate(st.session_state["ev"]):
            with st.container(border=True):
                cc = st.columns([6, 1])
                meta = f"guardado em {e.get('date','')}" + (f" · prova: {e['p']}" if e.get("p") else "")
                cc[0].markdown(f"<span class='pill'>{e['tag']}</span> **{e['t']}**  \n{e.get('d','')}  \n"
                               f"<small style='color:#a59ebb'>{meta}</small>", unsafe_allow_html=True)
                if cc[1].button("excluir", key=f"delev_{j}"):
                    st.session_state["ev"].pop(j)
                    if storage_ready():
                        save_state(mentee, build_state())
                    st.rerun()
    else:
        st.caption("_Nenhuma evidência ainda._")

    # ---- HORIZONTE ----
    block_strip("HORIZONTE", "Da senioridade à liderança",
                "A virada pra sênior é o passo 1. Olhando além: no Book, depois do Analista III a carreira bifurca "
                "em Especialista (referência) ou Coordenador (liderança).")
    arc_html = "<div class='arc'>"
    for (b, sub), cls in zip(ARC, ["now", "next", "fork", ""]):
        arc_html += f"<div class='{cls}'><b>{b}</b>{sub}</div>"
    arc_html += "</div>"
    st.markdown(arc_html, unsafe_allow_html=True)
    st.info("Arraste cada barra pra onde você se vê hoje. **Nada se move sozinho** — o texto do DISC ao lado é só "
            "contexto pra ajudar a refletir. A evolução acontece quando você re-marca ao longo do tempo.", icon="🎚️")
    with st.expander("De onde vem esta leitura de liderança?"):
        st.markdown("De três fontes: o **fork de carreira** do Book (III → Especialista ou Coordenador), os **eixos "
                    "de liderança do Guia do Mentor NextGen** (as barras abaixo) e o seu **DISC** "
                    "(2024 IS → 2026 ID, com a Dominância subindo). Quando você trouxer a régua formal do cargo de "
                    "Coordenador da TBC, a gente troca esses eixos pelos critérios reais.")
    for kk, b, note in LEAD:
        st.select_slider(f"**{b}** — {note}", options=[NA] + LSCALE, key=f"lead_{kk}")

    # ---- LEITURA ----
    s = build_state()
    k = senioridade_idx(s)
    an, inf = s["comp"]["analitica"], s["comp"]["influencia"]
    lname = lambda v: LEVELS[v - 1] if v else "a definir"
    faltam = [t for kk, t in [("tempo", "tempo (3 anos)"), ("gerencia", "vagas de gerência"),
                              ("pos", "pós-graduação")] if not s["req"].get(kk)]
    block_strip("SÍNTESE", "Leitura de hoje", "Um resumo do retrato atual, pra você levar pra conversa.")
    st.markdown(
        f"<div class='leitura'>"
        f"<p>Você opera com <b>autonomia de sênior</b> em <b>{k['aut_sr']} de {k['rs']}</b> responsabilidades.</p>"
        f"<p>Competências do Book: Capacidade Analítica em <b>{lname(an)}</b>, Influência e Persuasão em <b>{lname(inf)}</b>.</p>"
        f"<p>Requisitos formais: " + ("faltam <b>" + ", ".join(faltam) + "</b>." if faltam else "todos atendidos.") + "</p>"
        f"<p>Horizonte de liderança: <b>{lideranca_idx(s)}%</b> — o DISC (D no topo em 2026) sustenta o caminho de "
        f"Coordenação/HRBP; o eixo a cuidar é a flexibilidade na gestão de pessoas.</p>"
        f"<p>Cofre: <b>{len(s['ev'])}</b> evidência(s).</p></div>", unsafe_allow_html=True)

    st.write("")
    if st.button("💾 Salvar progresso", type="primary", use_container_width=True):
        if storage_ready():
            save_state(mentee, build_state())
            st.success("Progresso salvo. Pode fechar e voltar quando quiser — fica guardado.")
        else:
            st.error("Persistência não configurada — não foi possível guardar.")


# ============================================================ PÁGINA ADMIN
def page_admin():
    inject_css()
    st.markdown("<div class='nextgen-header'><div class='nextgen-brand'>Modo administrativo · Elaine</div>"
                "<div class='nextgen-title'>Acompanhamento</div>"
                "<div class='nextgen-subtitle'>Leitura do que cada mentorada preencheu. Atualiza a cada salvamento delas.</div></div>",
                unsafe_allow_html=True)
    if not storage_ready():
        st.error("Persistência não configurada. Configure os secrets do Google Sheets.")
        return
    nomes = list_mentees()
    if not nomes:
        st.info("Ainda não há respostas guardadas.")
        return
    mentee = st.selectbox("Mentorada", nomes)
    s = load_state(mentee)
    if not s:
        st.warning("Sem estado salvo pra esta mentorada.")
        return
    k = senioridade_idx(s)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Prontidão p/ III", f"{k['total']}%")
    c2.metric("Autonomia sênior", f"{k['aut_sr']}/{k['rs']}")
    c3.metric("Requisitos", f"{k['req_hard']}/3")
    c4.metric("Liderança", f"{lideranca_idx(s)}%")

    block_strip("EVOLUÇÃO", "Linha do tempo", "Marcos salvos pela mentorada ao longo da jornada.")
    _evolucao_chart(load_marcos(mentee))

    block_strip("DETALHE", "Autonomia por responsabilidade", "")
    rs = RESP_BASE + s.get("custom", [])
    for i, label in enumerate(rs):
        a = s["aut"][i] if i < len(s["aut"]) else {"n": 0, "ev": ""}
        nivel = STEPS[a["n"] - 1] if a.get("n") else "—"
        st.markdown(f"- **{label}** — {nivel}" + (f"  \n  _{a['ev']}_" if a.get("ev") else ""))

    block_strip("DETALHE", "Cofre de evidências", "")
    if s.get("ev"):
        for e in s["ev"]:
            st.markdown(f"- <span class='pill'>{e.get('tag')}</span> **{e.get('t')}** — {e.get('d','')} "
                        f"<small>({e.get('date','')})</small>", unsafe_allow_html=True)
    else:
        st.caption("_vazio_")

    with st.expander("Estado bruto (JSON)"):
        st.json(s)


# ============================================================ ROTEAMENTO
def main():
    st.set_page_config(page_title="Mapa da Virada — NextGen", page_icon="✦",
                       layout="centered", initial_sidebar_state="collapsed")
    qp = st.query_params
    admin_val = qp.get("admin")
    admin_key = st.secrets.get("app", {}).get("admin_key", "elaine") if "app" in st.secrets else "elaine"
    if admin_val and admin_val == admin_key:
        page_admin()
    else:
        mentee = (qp.get("m") or "Paula").strip() or "Paula"
        page_mentee(mentee)


if __name__ == "__main__":
    main()
