# -*- coding: utf-8 -*-
"""
Mapa da Virada — Mentoria NextGen (Streamlit + Google Sheets)
Painel da mentorada + modo administrativo (?admin=...).

Persistência REAL no Google Sheets:
- aba "estado": uma linha por mentorada com o estado completo (JSON). Upsert.
- aba "marcos": append-only, um marco por linha (pra trajetória).

O session_state do Streamlit é efêmero; quem guarda é o Sheets.
Cada "Salvar progresso" grava o estado; cada "Salvar marco" anexa uma linha.
O modo admin lê do MESMO Sheets — por isso você acompanha entre encontros.
"""
import json
import datetime as dt

import streamlit as st

# ----------------------------------------------------------------------------
# CONSTANTES (espelham o painel HTML, ancoradas no Book de Carreira RH)
# ----------------------------------------------------------------------------
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
                 "a opiniões divergentes — chegar a Domínio aqui passa por influenciar OUTRAS áreas, "
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

ARC = [
    ("Analista II", "você hoje"),
    ("Analista III", "a virada (Frente 1)"),
    ("Especialista / Coordenador", "fork: técnica ou liderança"),
    ("Diretor", "horizonte"),
]
DISC = [("2024", "IS"), ("2025", "IS"), ("2026", "ID")]


# ----------------------------------------------------------------------------
# ARMAZENAMENTO (Google Sheets)
# ----------------------------------------------------------------------------
HEAD_ESTADO = ["mentee", "updated_at", "json"]
HEAD_MARCOS = ["mentee", "data", "senioridade", "lideranca", "ts"]


@st.cache_resource(show_spinner=False)
def _open_sheet():
    """Abre a planilha usando a conta de serviço dos secrets. Retorna o objeto Spreadsheet."""
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_url(st.secrets["app"]["sheet_url"])


def storage_ready():
    return ("gcp_service_account" in st.secrets) and ("app" in st.secrets) and \
           ("sheet_url" in st.secrets.get("app", {}))


def _ws(sh, title, header):
    try:
        ws = sh.worksheet(title)
    except Exception:
        ws = sh.add_worksheet(title=title, rows=200, cols=max(6, len(header)))
        ws.append_row(header)
    return ws


def save_state(mentee, state):
    sh = _open_sheet()
    ws = _ws(sh, "estado", HEAD_ESTADO)
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
    sh = _open_sheet()
    ws = _ws(sh, "estado", HEAD_ESTADO)
    for rec in ws.get_all_records():
        if str(rec.get("mentee", "")).strip() == mentee:
            try:
                return json.loads(rec.get("json") or "{}")
            except Exception:
                return {}
    return {}


def append_marco(mentee, sen, lid):
    sh = _open_sheet()
    ws = _ws(sh, "marcos", HEAD_MARCOS)
    today = dt.date.today().strftime("%d/%m")
    ws.append_row([mentee, today, sen, lid, dt.datetime.now().isoformat(timespec="seconds")])


def load_marcos(mentee):
    sh = _open_sheet()
    ws = _ws(sh, "marcos", HEAD_MARCOS)
    out = []
    for rec in ws.get_all_records():
        if str(rec.get("mentee", "")).strip() == mentee:
            out.append({"d": rec.get("data"),
                        "sen": int(rec.get("senioridade") or 0),
                        "lid": int(rec.get("lideranca") or 0)})
    return out


def list_mentees():
    sh = _open_sheet()
    ws = _ws(sh, "estado", HEAD_ESTADO)
    return [str(r.get("mentee")) for r in ws.get_all_records() if r.get("mentee")]


# ----------------------------------------------------------------------------
# CÁLCULOS
# ----------------------------------------------------------------------------
def _lvl(label):  # competências
    return ([NA] + LEVELS).index(label) if label in ([NA] + LEVELS) else 0


def _stp(label):  # autonomia
    return ([NA] + STEPS).index(label) if label in ([NA] + STEPS) else 0


def _lsc(label):  # liderança
    return ([NA] + LSCALE).index(label) if label in ([NA] + LSCALE) else 0


def resps():
    return RESP_BASE + st.session_state.get("custom", [])


def build_state():
    ss = st.session_state
    rs = resps()
    return {
        "comp": {"analitica": _lvl(ss.get("comp_analitica", NA)),
                 "influencia": _lvl(ss.get("comp_influencia", NA))},
        "aut": [{"n": _stp(ss.get(f"aut_{i}", NA)), "ev": ss.get(f"autev_{i}", "")}
                for i in range(len(rs))],
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
    v = sum(s["lead"].get(k, 0) for k in LEAD_KEYS)
    return round((v / (len(LEAD_KEYS) * 4.0)) * 100)


# ----------------------------------------------------------------------------
# ESTILO
# ----------------------------------------------------------------------------
def inject_css():
    st.markdown("""
    <style>
      .vhead{background:radial-gradient(120% 140% at 85% -10%,#3a2d63 0%,#1b1726 55%);
        color:#fff;padding:26px 28px;border-radius:16px;margin-bottom:8px}
      .vhead .k{font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:#c9bef0;font-weight:700}
      .vhead h1{font-size:30px;margin:6px 0 4px;font-weight:700}
      .vhead p{color:#d8d1f0;margin:0;font-size:14px}
      .arc{display:flex;gap:6px;margin:6px 0 14px;flex-wrap:wrap}
      .arc div{flex:1;min-width:120px;border-radius:10px;padding:10px;text-align:center;
        background:#f3f1f8;border:1px solid #e6e2ee;font-size:11px;color:#6c6580}
      .arc b{display:block;font-size:13px;color:#1b1726}
      .arc .now{background:#1b1726;color:#fff}.arc .now b{color:#fff}
      .arc .next{background:#f6edd4;border-color:#e6d49a}
      .arc .fork{background:#e0f1f2;border-color:#bfe2e4}
      .pill{display:inline-block;background:#efe9fb;color:#6d28d9;border-radius:99px;
        padding:2px 10px;font-size:12px;font-weight:600;margin-right:6px}
    </style>
    """, unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# INICIALIZAÇÃO DO ESTADO NA SESSÃO
# ----------------------------------------------------------------------------
def hydrate(mentee):
    """Carrega do Sheets pra session_state (uma vez por mentorada/sessão)."""
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


# ----------------------------------------------------------------------------
# PÁGINA DA MENTORADA
# ----------------------------------------------------------------------------
def page_mentee(mentee):
    hydrate(mentee)
    inject_css()
    st.markdown(f"""<div class="vhead"><div class="k">Mentoria NextGen · Frente 1 + Horizonte</div>
      <h1>Mapa da Virada</h1>
      <p>{mentee} · Pleno (Analista II) &rarr; Sênior (Analista III), trilha de Atração e Seleção —
      e o passo seguinte rumo à liderança.</p></div>""", unsafe_allow_html=True)

    if not storage_ready():
        st.warning("Persistência não configurada (sem secrets do Google Sheets). "
                   "Dá pra navegar, mas o que você preencher NÃO será guardado.", icon="⚠️")

    s = build_state()
    k = senioridade_idx(s)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Prontidão p/ nível III", f"{k['total']}%")
    c2.metric("Autonomia sênior", f"{k['aut_sr']}/{k['rs']}")
    c3.metric("Requisitos formais", f"{k['req_hard']}/3")
    c4.metric("Prontidão p/ liderança", f"{lideranca_idx(s)}%")
    st.caption("Os números são retrato de hoje, não nota. Atualizam conforme você responde.")

    # --- COMPETÊNCIAS ---
    st.subheader("Os dois pontos da régua")
    st.caption("As duas competências de Relações Humanas no Book. Escolha o nível que mais descreve você — "
               "os comportamentos são os do próprio Book.")
    for key in ("analitica", "influencia"):
        c = COMPS[key]
        st.markdown(f"**{c['name']}** · {c['desc']}")
        st.radio(c["name"], [NA] + LEVELS, key=f"comp_{key}", horizontal=True,
                 label_visibility="collapsed")
        sel = _lvl(st.session_state[f"comp_{key}"])
        if sel:
            with st.container(border=True):
                st.markdown("  \n".join(["• " + b for b in c["behav"][sel - 1]]))
            if sel >= 3:
                st.success("Território sênior — conta a favor da virada.", icon="✅")
        if key == "influencia" and c.get("note"):
            st.info(c["note"], icon="🧭")
        st.divider()

    # --- AUTONOMIA ---
    st.subheader("Autonomia — onde mora a virada")
    st.caption("As responsabilidades são as MESMAS no II e no III; muda o nível de supervisão "
               "(moderado → baixo). De 'faço sozinha' pra cima = nível III.")
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
        st.text_input("Evidência rápida (exemplo recente)", key=f"autev_{i}",
                      placeholder="um exemplo concreto em que operou assim")
        if i >= len(RESP_BASE):
            if st.button("remover esta responsabilidade", key=f"delresp_{i}"):
                st.session_state["custom"].pop(i - len(RESP_BASE))
                st.session_state["_loaded"] = None  # força re-hidratar índices
                st.rerun()
        st.divider()
    with st.form("add_resp", clear_on_submit=True):
        novo = st.text_input("Adicionar uma responsabilidade sua que não está na lista")
        if st.form_submit_button("Adicionar responsabilidade") and novo.strip():
            st.session_state["custom"].append(novo.strip())
            st.rerun()

    # --- REQUISITOS ---
    st.subheader("Requisitos formais")
    st.caption("Critérios 'duros' do III. Às vezes a promoção trava aqui, independente da entrega.")
    for kk, b, sub, opt in REQ:
        st.checkbox(f"{b}" + (" · desejável" if opt else ""), key=f"req_{kk}", help=sub)
    st.divider()

    # --- COFRE ---
    st.subheader("Cofre de evidências")
    st.caption("As entregas que provam sua senioridade. Vá guardando — começa pelas 3 que considera nível sênior.")
    with st.form("add_ev", clear_on_submit=True):
        t = st.text_input("Título da entrega")
        tag = st.selectbox("Marcar como", ["Capacidade Analítica", "Influência e Persuasão",
                                           "Autonomia / processo", "Liderança"])
        d = st.text_input("2 linhas: o que era + o impacto")
        prova = st.text_input("Onde está a prova (link/local) — opcional")
        if st.form_submit_button("Guardar evidência") and t.strip():
            st.session_state["ev"].append({
                "t": t.strip(), "tag": tag, "d": d.strip(), "p": prova.strip(),
                "date": dt.date.today().strftime("%d/%m/%Y")})
            if storage_ready():
                save_state(mentee, build_state())
            st.rerun()
    if st.session_state["ev"]:
        for j, e in enumerate(st.session_state["ev"]):
            with st.container(border=True):
                cc = st.columns([6, 1])
                meta = f"guardado em {e.get('date','')}"
                if e.get("p"):
                    meta += f" · prova: {e['p']}"
                cc[0].markdown(f"<span class='pill'>{e['tag']}</span> **{e['t']}**  \n{e.get('d','')}  \n"
                               f"<small style='color:#a59ebb'>{meta}</small>", unsafe_allow_html=True)
                if cc[1].button("excluir", key=f"delev_{j}"):
                    st.session_state["ev"].pop(j)
                    if storage_ready():
                        save_state(mentee, build_state())
                    st.rerun()
    else:
        st.caption("_Nenhuma evidência ainda._")
    st.divider()

    # --- HORIZONTE / LIDERANÇA ---
    st.subheader("Horizonte — da senioridade à liderança")
    arc_html = "<div class='arc'>"
    cls = ["now", "next", "fork", ""]
    for (b, sub), c in zip(ARC, cls):
        arc_html += f"<div class='{c}'><b>{b}</b>{sub}</div>"
    arc_html += "</div>"
    st.markdown(arc_html, unsafe_allow_html=True)
    st.caption("DISC: " + " → ".join([f"{y} {t}" for y, t in DISC]) +
               " — Dominância sobe ao topo em 2026 (ID), leitura de liderança mais orientada a resultados.")
    for kk, b, note in LEAD:
        st.select_slider(f"**{b}** — {note}", options=[NA] + LSCALE, key=f"lead_{kk}")
    st.divider()

    # --- TRAJETÓRIA ---
    st.subheader("Trajetória")
    marcos = st.session_state.get("marcos", [])
    cur = senioridade_idx(build_state())["total"]
    sug = ""
    if not marcos and cur > 0:
        sug = "Você já preencheu bastante. Que tal salvar seu primeiro marco?"
    elif marcos and abs(cur - marcos[-1]["sen"]) >= 8:
        sug = f"Mudou bastante desde o último marco ({cur - marcos[-1]['sen']:+d} pts). Salvar um novo?"
    if sug:
        st.info(sug, icon="📍")
    if marcos:
        try:
            import pandas as pd
            df = pd.DataFrame({
                "senioridade": [m["sen"] for m in marcos],
                "liderança": [m.get("lid", 0) for m in marcos],
            }, index=[m["d"] for m in marcos])
            st.line_chart(df)
        except Exception:
            st.write(marcos)
        delta = marcos[-1]["sen"] - marcos[0]["sen"]
        st.caption(f"{len(marcos)} marco(s). Movimento: {delta:+d} pts desde o primeiro.")
    else:
        st.caption("_Nenhum marco salvo ainda._")
    if st.button("Salvar marco de hoje"):
        s2 = build_state()
        sen, lid = senioridade_idx(s2)["total"], lideranca_idx(s2)
        st.session_state["marcos"].append({"d": dt.date.today().strftime("%d/%m"), "sen": sen, "lid": lid})
        if storage_ready():
            save_state(mentee, s2)
            append_marco(mentee, sen, lid)
        st.rerun()
    st.divider()

    # --- LEITURA + SALVAR ---
    s = build_state()
    k = senioridade_idx(s)
    an, inf = s["comp"]["analitica"], s["comp"]["influencia"]
    ln = lambda v: LEVELS[v - 1] if v else "a definir"
    faltam = [t for kk, t in [("tempo", "tempo (3 anos)"), ("gerencia", "vagas de gerência"),
                              ("pos", "pós-graduação")] if not s["req"].get(kk)]
    st.subheader("Leitura de hoje")
    st.markdown(
        f"- Você opera com **autonomia de sênior** em **{k['aut_sr']} de {k['rs']}** responsabilidades.\n"
        f"- Competências: Capacidade Analítica em **{ln(an)}**, Influência e Persuasão em **{ln(inf)}**.\n"
        f"- Requisitos formais: " + ("faltam: **" + ", ".join(faltam) + "**" if faltam else "todos atendidos") + ".\n"
        f"- Horizonte de liderança: **{lideranca_idx(s)}%** — o DISC (D no topo em 2026) sustenta o caminho de "
        f"Coordenação/HRBP; cuidar da flexibilidade na gestão de pessoas.\n"
        f"- Cofre: **{len(s['ev'])}** evidência(s). Meta do Encontro 03: 3 entregas nível sênior bem descritas."
    )

    st.write("")
    if st.button("💾 Salvar progresso", type="primary", use_container_width=True):
        if storage_ready():
            save_state(mentee, build_state())
            st.success("Progresso salvo. Você pode fechar e voltar depois — fica guardado.")
        else:
            st.error("Persistência não configurada — não foi possível guardar.")


# ----------------------------------------------------------------------------
# PÁGINA ADMIN
# ----------------------------------------------------------------------------
def page_admin():
    inject_css()
    st.markdown("""<div class="vhead"><div class="k">Modo administrativo · Elaine</div>
      <h1>Acompanhamento — Mapa da Virada</h1>
      <p>Leitura do que cada mentorada preencheu. Atualiza a cada salvamento delas.</p></div>""",
                unsafe_allow_html=True)
    if not storage_ready():
        st.error("Persistência não configurada. Configure os secrets do Google Sheets pra ver os dados.")
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

    st.subheader("Autonomia por responsabilidade")
    rs = RESP_BASE + s.get("custom", [])
    for i, label in enumerate(rs):
        a = s["aut"][i] if i < len(s["aut"]) else {"n": 0, "ev": ""}
        nivel = STEPS[a["n"] - 1] if a.get("n") else "—"
        st.markdown(f"- **{label}** — {nivel}" + (f"  \n  _{a['ev']}_" if a.get("ev") else ""))

    st.subheader("Cofre de evidências")
    if s.get("ev"):
        for e in s["ev"]:
            st.markdown(f"- <span class='pill'>{e.get('tag')}</span> **{e.get('t')}** — {e.get('d','')} "
                        f"<small>({e.get('date','')})</small>", unsafe_allow_html=True)
    else:
        st.caption("_vazio_")

    st.subheader("Trajetória (marcos)")
    marcos = load_marcos(mentee)
    if marcos:
        try:
            import pandas as pd
            df = pd.DataFrame({"senioridade": [m["sen"] for m in marcos],
                               "liderança": [m["lid"] for m in marcos]},
                              index=[m["d"] for m in marcos])
            st.line_chart(df)
            st.download_button("Baixar marcos (CSV)", df.to_csv().encode("utf-8"),
                               file_name=f"marcos_{mentee}.csv")
        except Exception:
            st.write(marcos)
    else:
        st.caption("_sem marcos ainda_")

    with st.expander("Estado bruto (JSON)"):
        st.json(s)


# ----------------------------------------------------------------------------
# ROTEAMENTO
# ----------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Mapa da Virada — NextGen", page_icon="🧭", layout="centered")
    qp = st.query_params
    admin_val = qp.get("admin")
    admin_key = st.secrets.get("app", {}).get("admin_key", "elaine") if "app" in st.secrets else "elaine"
    if admin_val and admin_val == admin_key:
        page_admin()
    else:
        mentee = (qp.get("m") or "").strip()
        if not mentee:
            mentee = st.text_input("Seu nome (pra eu guardar seu progresso)", value="Paula").strip() or "Paula"
        page_mentee(mentee)


if __name__ == "__main__":
    main()
