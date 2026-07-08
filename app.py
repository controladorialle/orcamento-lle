"""
LLE Orçamento — Acompanhamento orçado x realizado (Despesas)
Streamlit + Supabase. Segredos: SUPABASE_URL e SUPABASE_ANON_KEY.
Repaginação: menu horizontal (tabs), filtros no corpo, resumo em 2 colunas
(Mês x Acumulado YTD) e drill-down por CR -> conta para os desvios.
"""
import re
import unicodedata
import pandas as pd
import streamlit as st
from supabase import create_client

# ---------------------------------------------------------------- identidade LLE
AZUL_PROFUNDO = "#071639"; AZUL_CORP = "#183F78"; AMARELO = "#F8B11E"
VERDE = "#0F8C3B"; VERMELHO = "#C0392B"; CINZA_TXT = "#6B7583"
CINZA_BG = "#F5F7FA"; LINHA = "#E4E8F0"

st.set_page_config(page_title="LLE Orçamento", page_icon="📊", layout="wide",
                   initial_sidebar_state="collapsed")
URL = st.secrets.get("SUPABASE_URL", ""); ANON = st.secrets.get("SUPABASE_ANON_KEY", "")
MESES = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
MABREV = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
STATUS_LABEL = {"PENDENTE": "Pendente", "JUSTIFICADO": "Justificado", "EM_REVISAO": "Em revisão", "DEVOLVIDO": "Devolvido", "APROVADO": "Aprovado"}

LOGO = f"""<svg width="42" height="54" viewBox="0 0 32 44" xmlns="http://www.w3.org/2000/svg">
<polygon points="16,2 23,9 16,16 9,9" fill="{AMARELO}"/><polygon points="16,10 23,17 16,24 9,17" fill="{AMARELO}"/>
<polygon points="16,18 23,25 16,32 9,25" fill="{VERDE}"/><polygon points="16,26 23,33 16,40 9,33" fill="{AZUL_CORP}"/></svg>"""

def inject_css():
    st.markdown(f"""<style>
      @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&display=swap');
      html, body, [class*="css"], .stApp {{ font-family:'Montserrat',sans-serif; }}
      .stApp {{ background:{CINZA_BG}; }}
      #MainMenu, footer, header[data-testid="stHeader"] {{ visibility:hidden; }}
      .block-container {{ padding-top:1.1rem; max-width:1360px; }}
      div.stButton>button[kind="primary"] {{ background:{AZUL_CORP}; border-color:{AZUL_CORP}; }}
      div.stButton>button[kind="primary"]:hover {{ background:{AZUL_PROFUNDO}; border-color:{AZUL_PROFUNDO}; }}

      /* header */
      .lle-header {{ background:linear-gradient(135deg,{AZUL_PROFUNDO} 0%,{AZUL_CORP} 100%);
        border-bottom:3px solid {AMARELO}; border-radius:12px; padding:15px 24px; display:flex;
        align-items:center; justify-content:space-between; margin-bottom:10px;
        box-shadow:0 4px 16px rgba(7,22,57,.15); }}
      .lle-hl {{ display:flex; align-items:center; gap:15px; }}
      .lle-header h1 {{ color:#fff; font-size:20px; font-weight:700; margin:0; letter-spacing:-.3px; }}
      .lle-header p {{ color:rgba(255,255,255,.72); font-size:12px; margin:3px 0 0; font-weight:500; }}
      .lle-badge {{ color:#fff; font-size:12px; font-weight:600; background:rgba(255,255,255,.12);
        border:1px solid rgba(255,255,255,.22); border-radius:20px; padding:6px 14px; white-space:nowrap; }}
      .lle-badge span {{ color:{AMARELO}; }}
      .modtag {{ font-size:15px; font-weight:700; color:{AZUL_PROFUNDO}; margin:2px 0 2px; }}
      .modsub {{ font-size:12px; color:{CINZA_TXT}; margin:0 0 8px; }}

      /* tabs horizontais (barra navy + sublinhado dourado) */
      .stTabs [data-baseweb="tab-list"] {{ gap:2px; background:{AZUL_PROFUNDO}; padding:0 10px;
        border-radius:10px; margin-bottom:18px; box-shadow:0 3px 10px rgba(7,22,57,.14); }}
      .stTabs [data-baseweb="tab"] {{ background:transparent; color:rgba(255,255,255,.72);
        padding:13px 22px; font-weight:600; font-size:14px; border-bottom:3px solid transparent; }}
      .stTabs [data-baseweb="tab"]:hover {{ color:#fff; background:rgba(255,255,255,.05); }}
      .stTabs [aria-selected="true"] {{ color:{AMARELO} !important; background:transparent !important;
        border-bottom:3px solid {AMARELO} !important; }}
      .stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none; }}

      /* resumo em 2 colunas */
      .resumo {{ display:grid; grid-template-columns:1fr 1fr; gap:14px; margin:4px 0 6px; }}
      @media (max-width:820px) {{ .resumo {{ grid-template-columns:1fr; }} }}
      .rpanel {{ background:#fff; border:1px solid {LINHA}; border-radius:12px; overflow:hidden;
        box-shadow:0 2px 8px rgba(7,22,57,.05); }}
      .rphead {{ background:{AZUL_PROFUNDO}; color:#fff; font-weight:700; font-size:13px;
        letter-spacing:.04em; padding:10px 16px; display:flex; justify-content:space-between; align-items:center; }}
      .rphead .tag {{ font-size:10px; font-weight:600; color:{AMARELO}; }}
      .rprow {{ display:flex; justify-content:space-between; align-items:center; padding:9px 16px;
        border-bottom:1px solid #F0F2F6; font-size:13px; }}
      .rprow:last-child {{ border-bottom:none; }}
      .rprow span {{ color:{CINZA_TXT}; }}
      .rprow b {{ color:{AZUL_PROFUNDO}; font-variant-numeric:tabular-nums; font-size:14px; }}
      .rprow.big b {{ font-size:16px; }}

      /* chips de status */
      .chip {{ display:inline-block; font-size:11px; font-weight:700; padding:3px 11px; border-radius:20px; }}

      /* cards contadores */
      .cards {{ display:flex; gap:12px; flex-wrap:wrap; margin:6px 0 4px; }}
      .card {{ flex:1; min-width:150px; background:#fff; border:1px solid {LINHA}; border-radius:12px;
        padding:13px 16px; box-shadow:0 2px 8px rgba(7,22,57,.05); }}
      .card .lab {{ font-size:11px; color:{CINZA_TXT}; text-transform:uppercase; letter-spacing:.04em; }}
      .card .val {{ font-size:21px; color:{AZUL_PROFUNDO}; font-weight:700; margin-top:3px; }}

      /* tabelas */
      table.lle {{ border-collapse:collapse; width:100%; font-size:13px; background:#fff;
        border-radius:10px; overflow:hidden; }}
      table.lle th {{ background:{AZUL_CORP}; color:#fff; padding:9px 12px; text-align:right; font-weight:600; }}
      table.lle th:first-child {{ text-align:left; }}
      table.lle td {{ padding:7px 12px; text-align:right; border-bottom:1px solid {LINHA}; font-variant-numeric:tabular-nums; }}
      table.lle td:first-child {{ text-align:left; }}
      table.lle tr.total td {{ font-weight:700; border-top:2px solid {AZUL_CORP}; background:#EEF2F8; }}
      table.lle tr.mark td {{ background:#FFF7E6; }}
      .scroll {{ overflow-x:auto; }}

      /* linha-cabeça do drill (CR) */
      .drow {{ display:grid; grid-template-columns:3.1fr 1.5fr 1.5fr 1.5fr 0.9fr 1.5fr; gap:0;
        align-items:center; background:#fff; border:1px solid {LINHA}; border-radius:9px;
        padding:9px 14px; font-size:13px; margin-bottom:2px; }}
      .drow.head {{ background:{AZUL_PROFUNDO}; color:#fff; font-weight:600; border:none; }}
      .drow .r {{ text-align:right; font-variant-numeric:tabular-nums; }}
      .drow .nm {{ font-weight:600; color:{AZUL_PROFUNDO}; }}
      div[data-testid="column"] div.stButton>button {{ padding:2px 0; border:1px solid {LINHA};
        background:#fff; color:{AZUL_CORP}; font-weight:700; border-radius:7px; }}

      /* rodapé */
      .lle-foot {{ margin-top:34px; padding:16px 22px; background:linear-gradient(135deg,{AZUL_PROFUNDO} 0%,{AZUL_CORP} 100%);
        border-radius:12px; display:flex; align-items:center; justify-content:space-between; }}
      .lle-foot .t {{ color:#fff; font-weight:600; font-size:13px; }}
      .lle-foot .s {{ color:#B8C5D9; font-size:11px; margin-top:3px; }}
      .lle-foot .v {{ color:#B8C5D9; font-size:11px; text-align:right; }}
    </style>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- helpers
def norm(s): return unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode().lower().strip()
def brl(n): return "R$ " + f"{(n or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
def pct_txt(n): return f"{n:+.1f}".replace(".", ",") + "%"
def chunks(seq, n=500):
    for i in range(0, len(seq), n): yield seq[i:i + n]

def chip(label, cor):
    return f"<span class='chip' style='color:{cor}; background:{cor}1A; border:1px solid {cor}55;'>{label}</span>"

def classifica(raw, pct, conta_cod, banda):
    """Retorna (label, cor) respeitando a convenção de sinal e a faixa neutra."""
    invert = str(conta_cod)[:1] in ("3", "6")     # receita/dedução: sinal invertido
    impacto = -raw if invert else raw               # >0 = pior resultado
    if abs(pct) <= banda:
        return "Neutro", CINZA_TXT
    return ("Desfavorável", VERMELHO) if impacto > 0 else ("Favorável", VERDE)

def var_de(vp, vr):
    raw = (vr or 0) - (vp or 0)
    pct = raw / vp * 100 if vp else (100 if vr else 0)
    return raw, pct

# ---------------------------------------------------------------- supabase
def client():
    c = create_client(URL, ANON)
    tok, rtok = st.session_state.get("access_token"), st.session_state.get("refresh_token")
    if tok and rtok:
        try: c.auth.set_session(tok, rtok)
        except Exception: pass
    return c

def get_faixa(c):
    try:
        r = c.table("config").select("valor").eq("chave", "faixa_neutra_pct").execute()
        return float(r.data[0]["valor"]) if r.data else 2.0
    except Exception:
        return 2.0

def set_faixa(c, valor):
    c.table("config").upsert({"chave": "faixa_neutra_pct", "valor": str(valor)}, on_conflict="chave").execute()

def get_cobranca(c):
    try:
        r = c.table("config").select("valor").eq("chave", "mes_inicio_cobranca").execute()
        return int(float(r.data[0]["valor"])) if r.data else 1
    except Exception:
        return 1

def set_cobranca(c, mes):
    c.table("config").upsert({"chave": "mes_inicio_cobranca", "valor": str(int(mes))}, on_conflict="chave").execute()

def ler_tudo(c, tabela, ano):
    """Le todas as linhas de uma tabela para o ano, paginando de 1000 em 1000."""
    linhas, passo, ini = [], 1000, 0
    while True:
        lote = c.table(tabela).select("*").eq("ano", ano).range(ini, ini + passo - 1).execute().data or []
        linhas.extend(lote)
        if len(lote) < passo:
            break
        ini += passo
    return linhas

def perfil(c, email):
    r = c.table("gestor_usuario").select("gestor_codigo, senha_provisoria, gestor(nome, papel)").eq("email", email).execute()
    if not r.data: return None
    g = r.data[0].get("gestor") or {}
    return {"gestor_codigo": r.data[0]["gestor_codigo"], "nome": g.get("nome", email),
            "papel": g.get("papel", "gestor"), "senha_provisoria": bool(r.data[0].get("senha_provisoria", False))}

# ---------------------------------------------------------------- leituras cacheadas
# Cache por TOKEN do usuário: respeita o RLS (cada perfil só cacheia o que pode ver)
# e é limpo em TODA escrita via limpar_cache(); TTL curto só como respaldo.
def _cli_tok(tok, rtok):
    cc = create_client(URL, ANON)
    if tok and rtok:
        try: cc.auth.set_session(tok, rtok)
        except Exception: pass
    return cc

@st.cache_data(ttl=300, show_spinner=False)
def _q_orc(tok, rtok, ano):
    cc = _cli_tok(tok, rtok)
    cols = "ano,mes,uni_cod,unidade,cr_cod,cr_nome,cr_grupo,conta_cod,conta_desc,valor_planejado,valor_realizado,tipo_conta,classificacao"
    linhas, passo, ini = [], 1000, 0
    while True:
        lote = cc.table("orc_realizado").select(cols).eq("ano", ano).range(ini, ini + passo - 1).execute().data or []
        linhas.extend(lote)
        if len(lote) < passo: break
        ini += passo
    return linhas

@st.cache_data(ttl=300, show_spinner=False)
def _q_cr_gestor(tok, rtok):
    cc = _cli_tok(tok, rtok)
    return cc.table("cr_gestor").select("uni_cod, cr_cod, cr_nome, gestor(nome)").execute().data or []

@st.cache_data(ttl=120, show_spinner=False)
def _q_justif(tok, rtok, ano, mes):
    cc = _cli_tok(tok, rtok)
    return cc.table("justificativa").select("*").eq("ano", ano).eq("mes", mes).execute().data or []

@st.cache_data(ttl=300, show_spinner=False)
def _q_oper_mes(tok, rtok, ano, mes):
    cc = _cli_tok(tok, rtok)
    linhas, passo, ini = [], 1000, 0
    while True:
        lote = (cc.table("operacional_detalhe").select("uni_cod,cr_cod,conta_cod,num_doc,valor,historico")
                .eq("ano", ano).eq("mes", mes).range(ini, ini + passo - 1).execute().data or [])
        linhas.extend(lote)
        if len(lote) < passo: break
        ini += passo
    return linhas

def _tok(): return st.session_state.get("access_token"), st.session_state.get("refresh_token")
def carregar_orc(ano): return _q_orc(*_tok(), ano)
def carregar_cr_gestor(): return _q_cr_gestor(*_tok())
def carregar_justificativas(ano, mes): return _q_justif(*_tok(), ano, mes)
def carregar_operacional(ano, mes): return _q_oper_mes(*_tok(), ano, mes)

def limpar_cache():
    """Chamar após QUALQUER escrita para forçar dados frescos na próxima leitura."""
    try: st.cache_data.clear()
    except Exception: pass

# ---------------------------------------------------------------- login / senha
def tela_trocar_senha(c, email):
    inject_css()
    _, c2, _ = st.columns([1, 1.3, 1])
    with c2:
        st.markdown(f"<div style='text-align:center; padding:20px 0 6px;'>{LOGO}</div>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown("##### Defina sua senha")
            st.caption("Este é seu primeiro acesso. Crie uma senha pessoal para continuar.")
            with st.form("trocar"):
                s1 = st.text_input("Nova senha (mínimo 6 caracteres)", type="password")
                s2 = st.text_input("Confirme a nova senha", type="password")
                ok = st.form_submit_button("Salvar e entrar", type="primary", use_container_width=True)
            if ok:
                if len(s1) < 6:
                    st.error("A senha precisa ter ao menos 6 caracteres.")
                elif s1 != s2:
                    st.error("As duas senhas não conferem.")
                else:
                    try:
                        c.auth.update_user({"password": s1})
                        c.table("gestor_usuario").update({"senha_provisoria": False}).eq("email", email).execute()
                        st.success("Senha definida! Recarregando...")
                        st.rerun()
                    except Exception:
                        st.error("Não consegui salvar a senha. Saia e entre novamente para tentar.")


def tela_login():
    inject_css()
    _, c2, _ = st.columns([1, 1.3, 1])
    with c2:
        st.markdown(f"""<div style="text-align:center; padding:26px 0 6px;"><div>{LOGO}</div>
            <h1 style="color:{AZUL_PROFUNDO}; margin:12px 0 2px;">Sistema de Acompanhamento de Orçamento</h1>
            <p style="color:{AZUL_CORP}; font-size:13px; margin:0;">LLE Ferragens · Controladoria</p></div>""", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown("##### Acesso ao sistema")
            with st.form("login"):
                email = st.text_input("E-mail corporativo", placeholder="seu.email@grupolle.com.br")
                senha = st.text_input("Senha", type="password")
                ok = st.form_submit_button("Entrar", type="primary", use_container_width=True)
            if ok:
                try:
                    cli = create_client(URL, ANON)
                    res = cli.auth.sign_in_with_password({"email": email.strip(), "password": senha})
                    st.session_state.access_token = res.session.access_token
                    st.session_state.refresh_token = res.session.refresh_token
                    st.session_state.email = email.strip(); st.rerun()
                except Exception:
                    st.error("E-mail ou senha inválidos.")
            st.caption("Use o e-mail cadastrado pela controladoria.")

def header(prof):
    papel = "Controladoria" if prof["papel"] == "controladoria" else "Gestor"
    st.markdown(f"""<div class="lle-header">
        <div class="lle-hl">{LOGO}<div>
          <h1>Sistema de Acompanhamento de Orçamento — LLE Ferragens</h1>
          <p>{prof['nome']}</p></div></div>
        <div class="lle-badge">GRUPO LLE <span>—</span> {papel}</div></div>""", unsafe_allow_html=True)

def rodape():
    st.markdown(f"""<div class="lle-foot">
        <div style="display:flex; align-items:center; gap:12px;">{LOGO}
          <div><div class="t">Sistema de Acompanhamento de Orçamento</div>
          <div class="s">Controladoria · Grupo LLE Ferragens · 2026</div></div></div>
        <div class="v"><div>Módulo Despesas · Orçado x Realizado</div>
          <div style="opacity:.7; margin-top:3px;">Todos os direitos reservados</div></div></div>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- blocos de insight
def resumo_colunas(d_mes, d_ytd, banda, mes):
    vp, vr = d_mes["valor_planejado"].sum(), d_mes["valor_realizado"].sum()
    raw, pct = var_de(vp, vr); lab, cor = classifica(raw, pct, "5", banda)
    yp, yr = d_ytd["valor_planejado"].sum(), d_ytd["valor_realizado"].sum()
    yraw, ypct = var_de(yp, yr); ylab, ycor = classifica(yraw, ypct, "5", banda)

    def panel(titulo, tag, o, r, va, p, lab, cor):
        return f"""<div class='rpanel'>
          <div class='rphead'><span>{titulo}</span><span class='tag'>{tag}</span></div>
          <div class='rprow'><span>Orçado</span><b>{brl(o)}</b></div>
          <div class='rprow'><span>Realizado</span><b>{brl(r)}</b></div>
          <div class='rprow'><span>Variação (R$)</span><b style='color:{cor}'>{brl(va)}</b></div>
          <div class='rprow'><span>Variação (%)</span><b style='color:{cor}'>{pct_txt(p)}</b></div>
          <div class='rprow big'><span>Status</span>{chip(lab, cor)}</div>
        </div>"""

    st.markdown("<div class='resumo'>"
        + panel("MÊS", MESES[mes] + "/2026", vp, vr, raw, pct, lab, cor)
        + panel("ACUMULADO YTD", f"Jan–{MABREV[mes]}/2026", yp, yr, yraw, ypct, ylab, ycor)
        + "</div>", unsafe_allow_html=True)

def contadores(df_mes, banda):
    fav = desf = neu = 0
    for _, r in df_mes.iterrows():
        raw, pct = var_de(r["valor_planejado"], r["valor_realizado"])
        lab, _ = classifica(raw, pct, r["conta_cod"], banda)
        fav += lab == "Favorável"; desf += lab == "Desfavorável"; neu += lab == "Neutro"
    html = "<div class='cards'>"
    for lab, val, cor in [("Lançamentos", len(df_mes), AZUL_PROFUNDO), ("Favoráveis", fav, VERDE),
                          ("Desfavoráveis", desf, VERMELHO), ("Neutros", neu, CINZA_TXT)]:
        html += f"<div class='card'><div class='lab'>{lab}</div><div class='val' style='color:{cor}'>{val}</div></div>"
    st.markdown(html + "</div>", unsafe_allow_html=True)

def tabela_evolucao(df, banda, mes_sel):
    g = df.groupby("mes")[["valor_planejado", "valor_realizado"]].sum().reindex(range(1, 13), fill_value=0)
    cum = g.cumsum()
    linhas = ""
    for m in range(1, 13):
        vp, vr = g.loc[m, "valor_planejado"], g.loc[m, "valor_realizado"]
        raw, pct = var_de(vp, vr)
        if vr == 0:
            st_html = f"<td colspan='3' style='color:{CINZA_TXT}'>Sem realizado</td>"
        else:
            lab, cor = classifica(raw, pct, "5", banda)
            st_html = f"<td style='color:{cor}'>{brl(raw)}</td><td style='color:{cor}'>{pct_txt(pct)}</td><td>{chip(lab, cor)}</td>"
        yvp, yvr = cum.loc[m, "valor_planejado"], cum.loc[m, "valor_realizado"]
        yraw, ypct = var_de(yvp, yvr); ylab, ycor = classifica(yraw, ypct, "5", banda)
        yhtml = (f"<td>{brl(yvr)}</td><td style='color:{ycor}'>{brl(yraw)}</td>"
                 f"<td style='color:{ycor}'>{pct_txt(ypct)}</td>") if yvr else "<td>—</td><td>—</td><td>—</td>"
        cls = " class='mark'" if m == mes_sel else ""
        linhas += f"<tr{cls}><td>{MESES[m]}</td><td>{brl(vp)}</td><td>{brl(vr) if vr else '—'}</td>{st_html}{yhtml}</tr>"
    st.markdown(f"""<div class='scroll'><table class="lle"><tr>
        <th>Mês</th><th>Orçado</th><th>Realizado</th><th>Var. (R$)</th><th>Var. (%)</th><th>Status</th>
        <th>Realizado YTD</th><th>Var. YTD (R$)</th><th>Var. YTD (%)</th></tr>{linhas}</table></div>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- DRILL-DOWN por CR -> conta
def drill_desvios(d_mes, banda, mes):
    if d_mes.empty:
        st.info("Sem lançamentos para os filtros atuais."); return
    only_desf = st.checkbox("Mostrar apenas centros com desvio desfavorável", value=False, key=f"drill_only_{mes}")

    g = d_mes.groupby(["cr_cod", "cr_nome"], as_index=False)[["valor_planejado", "valor_realizado"]].sum()
    rows = []
    for _, r in g.iterrows():
        raw, pct = var_de(r["valor_planejado"], r["valor_realizado"])
        lab, cor = classifica(raw, pct, "5", banda)
        rows.append((int(r["cr_cod"]), r["cr_nome"], r["valor_planejado"], r["valor_realizado"], raw, pct, lab, cor))
    if only_desf:
        rows = [x for x in rows if x[6] == "Desfavorável"]
    rows.sort(key=lambda x: x[4], reverse=True)
    if not rows:
        st.success("Nenhum centro de resultado com desvio desfavorável neste recorte."); return

    st.caption("Clique em ▶ para abrir as contas do centro de resultado. Ordenado do maior gasto acima do previsto para o menor.")
    # cabeçalho alinhado (mesma grade do corpo, com espaço para o botão)
    ch1, ch2 = st.columns([0.05, 0.95])
    ch2.markdown("""<div class="drow head"><div class="nm">Centro de resultado</div>
        <div class="r">Orçado</div><div class="r">Realizado</div><div class="r">Var. (R$)</div>
        <div class="r">Var. (%)</div><div class="r">Status</div></div>""", unsafe_allow_html=True)

    for cr_cod, cr_nome, vp, vr, raw, pct, lab, cor in rows:
        key = f"drill_{mes}_{cr_cod}"
        if key not in st.session_state:
            st.session_state[key] = False
        exp = st.session_state[key]
        cbtn, cbody = st.columns([0.05, 0.95])
        with cbtn:
            if st.button("▼" if exp else "▶", key=f"btn_{key}"):
                st.session_state[key] = not exp; st.rerun()
        with cbody:
            st.markdown(f"""<div class="drow"><div class="nm">{cr_nome}</div>
                <div class="r">{brl(vp)}</div><div class="r">{brl(vr)}</div>
                <div class="r" style="color:{cor};font-weight:600">{brl(raw)}</div>
                <div class="r" style="color:{cor}">{pct_txt(pct)}</div>
                <div class="r">{chip(lab, cor)}</div></div>""", unsafe_allow_html=True)
            if exp:
                sub = d_mes[d_mes["cr_cod"] == cr_cod]
                det = []
                for _, v in sub.iterrows():
                    rw, pc = var_de(v["valor_planejado"], v["valor_realizado"])
                    lb, co = classifica(rw, pc, v["conta_cod"], banda)
                    det.append((f"{v['conta_cod']} · {v.get('conta_desc','')}", v["valor_planejado"], v["valor_realizado"], rw, pc, lb, co))
                det.sort(key=lambda x: x[3], reverse=True)
                linhas = ""
                for nome, o, r, rw, pc, lb, co in det:
                    linhas += (f"<tr><td>{nome}</td><td>{brl(o)}</td><td>{brl(r)}</td>"
                               f"<td style='color:{co}'>{brl(rw)}</td><td style='color:{co}'>{pct_txt(pc)}</td>"
                               f"<td>{chip(lb, co)}</td></tr>")
                st.markdown(f"""<div class='scroll' style='margin:2px 0 10px 8px;'><table class="lle"><tr>
                    <th>Conta</th><th>Orçado</th><th>Realizado</th><th>Var. (R$)</th><th>Var. (%)</th><th>Status</th>
                    </tr>{linhas}</table></div>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- justificativas
def secao_justificativas(c, prof, df_mes, mes, is_ctrl, banda):
    if mes < get_cobranca(c):
        st.info(f"{MESES[mes]}/2026 não está sujeito à cobrança de justificativa.")
        return
    js = carregar_justificativas(2026, mes)
    jmap = {(j["uni_cod"], j["cr_cod"], j["conta_cod"]): j for j in js}
    itens = []
    for _, v in df_mes.iterrows():
        raw, pct = var_de(v["valor_planejado"], v["valor_realizado"])
        lab, _ = classifica(raw, pct, v["conta_cod"], banda)
        if lab != "Desfavorável":
            continue
        j = jmap.get((v["uni_cod"], v["cr_cod"], v["conta_cod"]), {"status": "PENDENTE", "texto": "", "comentario_controladoria": ""})
        itens.append((v, raw, pct, j))
    itens.sort(key=lambda t: t[1], reverse=True)
    st.caption(f"{len(itens)} desvio(s) desfavorável(is) a justificar em {MESES[mes]}/2026")
    st_cor = {"APROVADO": VERDE, "DEVOLVIDO": VERMELHO, "PENDENTE": CINZA_TXT, "JUSTIFICADO": AZUL_CORP, "EM_REVISAO": AZUL_CORP}
    raw_all = pd.DataFrame(carregar_orc(2026)) if itens else pd.DataFrame()
    oper_all = pd.DataFrame(carregar_operacional(2026, mes)) if itens else pd.DataFrame()
    for v, raw, pct, j in itens:
        status = j.get("status", "PENDENTE")
        titulo = f"{v['conta_cod']} · {v.get('conta_desc','')} — {v.get('cr_nome','')} ({v.get('unidade','')}) · {brl(raw)} · [{STATUS_LABEL.get(status)}]"
        with st.expander(titulo):
            a, b, d = st.columns(3)
            a.metric("Orçado", brl(v["valor_planejado"])); b.metric("Realizado", brl(v["valor_realizado"])); d.metric("Variação", brl(raw), pct_txt(pct))
            st.markdown("**Composição do realizado por empresa**")
            comp = ([] if raw_all.empty else raw_all[(raw_all["mes"] == mes) & (raw_all["cr_cod"] == v["cr_cod"]) & (raw_all["conta_cod"] == v["conta_cod"])]
                    [["uni_cod", "unidade", "valor_planejado", "valor_realizado"]].to_dict("records"))
            if comp:
                linhas = ""
                for e in sorted(comp, key=lambda x: x["uni_cod"]):
                    er, _ = var_de(e["valor_planejado"], e["valor_realizado"])
                    cor = VERMELHO if er > 0 else VERDE
                    nome = e.get("unidade") or ("PISA" if e["uni_cod"] == 1 else "KING" if e["uni_cod"] == 2 else f"Empresa {e['uni_cod']}")
                    linhas += (f"<tr><td>{e['uni_cod']} · {nome}</td><td>{brl(e['valor_planejado'])}</td>"
                               f"<td>{brl(e['valor_realizado'])}</td><td style='color:{cor}'>{brl(er)}</td></tr>")
                st.markdown(f"""<table class="lle"><tr><th>Empresa</th><th>Orçado</th><th>Realizado</th><th>Variação</th></tr>{linhas}
                    <tr class='total'><td>Net do CR</td><td>{brl(v['valor_planejado'])}</td><td>{brl(v['valor_realizado'])}</td><td>{brl(raw)}</td></tr></table>""", unsafe_allow_html=True)
            st.markdown("**Histórico das notas que compõem o realizado**")
            det = ([] if oper_all.empty else oper_all[(oper_all["cr_cod"] == v["cr_cod"]) & (oper_all["conta_cod"] == v["conta_cod"])]
                   [["uni_cod", "num_doc", "valor", "historico"]].to_dict("records"))
            if det:
                dfd = pd.DataFrame(det)
                dfd["uni_cod"] = dfd["uni_cod"].map(lambda u: f"{u} · {'PISA' if u == 1 else 'KING' if u == 2 else '—'}")
                dfd["valor"] = dfd["valor"].map(brl)
                dfd = dfd[["uni_cod", "num_doc", "valor", "historico"]]
                dfd.columns = ["Empresa", "NF/Doc", "Valor", "Histórico"]
                st.dataframe(dfd, use_container_width=True, hide_index=True)
            else:
                st.caption("Sem detalhe operacional para esta conta neste mês.")
            key = dict(ano=2026, mes=mes, uni_cod=int(v["uni_cod"]), cr_cod=int(v["cr_cod"]), conta_cod=int(v["conta_cod"]))
            kb = f"{mes}_{v['uni_cod']}_{v['cr_cod']}_{v['conta_cod']}"
            if status == "DEVOLVIDO" and j.get("comentario_controladoria"):
                st.warning(f"Controladoria: {j['comentario_controladoria']}")
            if not is_ctrl and status in ("PENDENTE", "DEVOLVIDO"):
                txt = st.text_area("Justificativa", value=j.get("texto", "") or "", key=f"txt_{kb}")
                c1, c2 = st.columns(2)
                if c1.button("Salvar rascunho", key=f"sv_{kb}"):
                    c.table("justificativa").upsert({**key, "texto": txt, "status": "PENDENTE", "atualizado_por": prof["nome"]}, on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute(); limpar_cache(); st.rerun()
                if c2.button("Enviar justificativa", key=f"en_{kb}", type="primary"):
                    if not txt.strip(): st.error("Escreva a justificativa antes de enviar.")
                    else:
                        c.table("justificativa").upsert({**key, "texto": txt, "status": "JUSTIFICADO", "atualizado_por": prof["nome"]}, on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute(); limpar_cache(); st.rerun()
            elif not is_ctrl:
                st.info(f"Justificativa: {j.get('texto') or '—'}"); st.caption("Aguardando a controladoria — não editável.")
            if is_ctrl:
                st.info(f"Justificativa do gestor: {j.get('texto') or '— (ainda não enviada)'}")
                if status in ("JUSTIFICADO", "EM_REVISAO"):
                    coment = st.text_input("Comentário (para devolução)", key=f"cm_{kb}")
                    c1, c2 = st.columns(2)
                    if c1.button("Aprovar", key=f"ap_{kb}", type="primary"):
                        c.table("justificativa").update({"status": "APROVADO"}).match(key).execute(); limpar_cache(); st.rerun()
                    if c2.button("Devolver", key=f"dv_{kb}"):
                        c.table("justificativa").update({"status": "DEVOLVIDO", "comentario_controladoria": coment}).match(key).execute(); limpar_cache(); st.rerun()
    if not itens:
        st.success("Nenhum desvio desfavorável a justificar com os filtros atuais.")

# ---------------------------------------------------------------- importar / config
def tela_importar(c):
    st.markdown("<div class='modtag'>Configuração e importação de dados</div>", unsafe_allow_html=True)
    st.markdown("<div class='modsub'>Regras de classificação, cobrança e carga das bases mensais</div>", unsafe_allow_html=True)
    st.subheader("Configuração")
    atual = get_faixa(c)
    nova = st.number_input("Faixa neutra (±%) — regra de classificação para todos os gestores",
                           value=float(atual), step=0.5, min_value=0.0)
    if st.button("Salvar faixa neutra"):
        set_faixa(c, nova); st.success(f"Faixa neutra atualizada para ±{nova:.1f}%.".replace(".", ","))
    st.markdown("**Cobrança de justificativas**")
    mc = get_cobranca(c)
    novo_mc = st.selectbox("Cobrar justificativas a partir do mês (2026)", list(range(1, 13)), index=mc - 1, format_func=lambda m: MESES[m])
    st.caption("Meses anteriores a este não geram pendência nem cobrança (continuam visíveis nas análises).")
    if st.button("Salvar início da cobrança"):
        set_cobranca(c, novo_mc); st.success(f"Cobrança a partir de {MESES[novo_mc]}/2026.")
    st.divider()
    st.subheader("Importar dados")
    st.caption("A base orçado x realizado atualiza os números; o arquivo operacional traz o histórico das notas.")
    ano = st.number_input("Ano", value=2026, step=1)
    def pick(row, cands):
        for cand in cands:
            for k in row.index:
                if norm(k) == cand or cand in norm(k): return row[k]
        return ""
    def num(x):
        if x is None or pd.isna(x): return 0.0
        if isinstance(x, (int, float)): return float(x)
        s = str(x).replace("R$", "").strip()
        if not s: return 0.0
        if "," in s: s = s.replace(".", "").replace(",", ".")
        try: return float(s)
        except ValueError: return 0.0
    def toint(x):
        if x is None or pd.isna(x): return None
        s = str(x).strip()
        try: return int(float(s)) if s else None
        except ValueError: return None
    f1 = st.file_uploader("1) Base orçado x realizado", type=["xlsx"], key="f1")
    if f1 and st.button("Importar orçado x realizado"):
        df = pd.read_excel(f1); recs = []
        for _, r in df.iterrows():
            uni, cr, ct = toint(pick(r, ["codigo unidade"])), toint(pick(r, ["codigo centro"])), toint(pick(r, ["codigo conta"]))
            if None in (uni, cr, ct): continue
            recs.append(dict(ano=toint(pick(r, ["ano"])) or int(ano), mes=toint(pick(r, ["mes"])) or 1, uni_cod=uni,
                unidade=str(pick(r, ["descricao unidade"]) or ""), cr_cod=cr, cr_nome=str(pick(r, ["descricao centro"]) or ""),
                cr_grupo=str(pick(r, ["cr grupo"]) or ""), conta_cod=ct, conta_desc=str(pick(r, ["descricao conta"]) or ""),
                valor_planejado=num(pick(r, ["valor planejado"])), valor_realizado=num(pick(r, ["valor realizado"])),
                tipo_conta=str(pick(r, ["tipo conta"]) or ""), classificacao=str(pick(r, ["classificacao"]) or "")))
        for ch in chunks(recs):
            c.table("orc_realizado").upsert(ch, on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute()
        st.success(f"{len(recs)} linhas de orçado x realizado importadas.")
        limpar_cache()
    f2 = st.file_uploader("2) Detalhe operacional (com COMPLHIST)", type=["xlsx"], key="f2")
    mes_op = st.number_input("Mês do arquivo operacional", value=6, min_value=1, max_value=12, step=1)
    if f2 and st.button("Importar histórico das notas"):
        raw = pd.read_excel(f2, header=None, nrows=10); hi = 0
        for i in range(min(10, len(raw))):
            if any("complhist" in norm(x) for x in raw.iloc[i].tolist()): hi = i; break
        df = pd.read_excel(f2, header=hi).dropna(how="all"); recs = []
        for _, r in df.iterrows():
            uni, cr, ct = toint(pick(r, ["codigo unidade", "codigo_unidade"])), toint(pick(r, ["codcencus", "codigo centro"])), toint(pick(r, ["codctactb", "codigo conta"]))
            if None in (uni, cr, ct): continue
            h = pick(r, ["complhist"])
            m = re.search(r"\b(?:NF|CFE)\s*0*([0-9]{3,})", str(h), re.I) if isinstance(h, str) else None
            hist = re.sub(r"\s*-?\d+\.\d+\s*$", "", str(h)).strip() if isinstance(h, str) else str(h or "")
            recs.append(dict(ano=toint(pick(r, ["ano"])) or int(ano), mes=toint(pick(r, ["mes"])) or int(mes_op),
                uni_cod=uni, cr_cod=cr, conta_cod=ct, num_doc=(m.group(1) if m else None), valor=num(pick(r, ["valor"])), historico=hist))
        c.table("operacional_detalhe").delete().eq("ano", int(ano)).eq("mes", int(mes_op)).execute()
        for ch in chunks(recs):
            c.table("operacional_detalhe").insert(ch).execute()
        st.success(f"{len(recs)} lançamentos com histórico importados.")
        limpar_cache()

# ---------------------------------------------------------------- painel de recebidas
def tela_painel(c, prof, banda, df_orc, cg):
    st.markdown("<div class='modtag'>Justificativas recebidas</div>", unsafe_allow_html=True)
    st.markdown("<div class='modsub'>Pendências por gestor, resumo e detalhe das respostas</div>", unsafe_allow_html=True)
    mes = st.selectbox("Mês", list(range(1, 13)), index=5, format_func=lambda m: f"{MESES[m]}/2026", key="painel_mes")

    if mes < get_cobranca(c):
        st.info(f"{MESES[mes]}/2026 não está sujeito à cobrança de justificativa.")
        return

    js = carregar_justificativas(2026, mes)

    # ----- Gerência de pendências (desvios desfavoráveis ainda não enviados) -----
    dfo = df_orc[df_orc["mes"] == mes] if not df_orc.empty else df_orc
    enviadas = {(int(j["uni_cod"]), int(j["cr_cod"]), int(j["conta_cod"])) for j in js if j.get("status") in ("JUSTIFICADO", "EM_REVISAO", "APROVADO")}
    devolvidas = {(int(j["uni_cod"]), int(j["cr_cod"]), int(j["conta_cod"])) for j in js if j.get("status") == "DEVOLVIDO"}
    pend = {}
    for _, v in dfo.iterrows():
        raw, pct = var_de(v["valor_planejado"], v["valor_realizado"])
        lab, _ = classifica(raw, pct, v["conta_cod"], banda)
        if lab != "Desfavorável":
            continue
        chave = (int(v["uni_cod"]), int(v["cr_cod"]), int(v["conta_cod"]))
        if chave in enviadas:
            continue
        gestor = cg.get((int(v["uni_cod"]), int(v["cr_cod"])), ("Sem gestor", ""))[0]
        d = pend.setdefault(gestor, {"n": 0, "valor": 0.0, "devolv": 0, "itens": []})
        d["n"] += 1; d["valor"] += raw
        dev = chave in devolvidas
        if dev: d["devolv"] += 1
        d["itens"].append({"cr": v.get("cr_nome", ""), "conta": f"{v['conta_cod']} · {v.get('conta_desc','')}", "raw": raw, "dev": dev})

    st.markdown("###### Pendências por gestor (desvios desfavoráveis ainda não justificados)")
    if not pend:
        st.success("Nenhuma pendência neste mês — todos os desvios desfavoráveis foram justificados.")
    else:
        total = sum(d["valor"] for d in pend.values())
        st.caption(f"Total: {sum(d['n'] for d in pend.values())} conta(s) · {brl(total)} de desvio parado")
        for gestor in sorted(pend, key=lambda g: -pend[g]["valor"]):
            d = pend[gestor]
            titulo = f"{gestor} — {d['n']} conta(s) · {brl(d['valor'])}" + (f" · {d['devolv']} devolvida(s)" if d["devolv"] else "")
            with st.expander(titulo):
                linhas = ""
                for it in sorted(d["itens"], key=lambda x: -x["raw"]):
                    tag = " (devolvida)" if it["dev"] else ""
                    linhas += (f"<tr><td>{it['cr']}</td><td>{it['conta']}{tag}</td>"
                               f"<td style='color:{VERMELHO}'>{brl(it['raw'])}</td></tr>")
                st.markdown(f"""<table class="lle"><tr><th>Centro de resultado</th><th>Conta</th>
                    <th>Desvio (R$)</th></tr>{linhas}</table>""", unsafe_allow_html=True)
    st.divider()

    if not js:
        st.info(f"Ainda não há justificativas enviadas em {MESES[mes]}/2026. As pendências acima mostram o que falta.")
        return

    linhas = []
    for j in js:
        gestor, cr_nome = cg.get((j["uni_cod"], j["cr_cod"]), ("—", str(j["cr_cod"])))
        linhas.append({"gestor": gestor, "cr": cr_nome, "conta": j["conta_cod"],
                       "status": j.get("status", "PENDENTE"), "texto": j.get("texto", "") or "",
                       "comentario": j.get("comentario_controladoria", "") or "",
                       "_k": (j["uni_cod"], j["cr_cod"], j["conta_cod"])})
    df = pd.DataFrame(linhas)

    st.markdown("###### Resumo por gestor")
    resumo = ""
    for gestor in sorted(df["gestor"].unique()):
        sub = df[df["gestor"] == gestor]
        cnt = {s: int((sub["status"] == s).sum()) for s in STATUS_LABEL}
        resumo += (f"<tr><td>{gestor}</td><td>{len(sub)}</td>"
                   f"<td style='color:{VERMELHO}'>{cnt['PENDENTE']+cnt['DEVOLVIDO']}</td>"
                   f"<td style='color:{AZUL_CORP}'>{cnt['JUSTIFICADO']+cnt['EM_REVISAO']}</td>"
                   f"<td style='color:{VERDE}'>{cnt['APROVADO']}</td></tr>")
    st.markdown(f"""<table class="lle"><tr><th>Gestor</th><th>Total</th>
        <th>A responder</th><th>Aguardando controladoria</th><th>Aprovadas</th></tr>{resumo}</table>""", unsafe_allow_html=True)

    st.markdown("###### Detalhe por gestor e centro de resultado")
    st_cor = {"APROVADO": VERDE, "DEVOLVIDO": VERMELHO, "PENDENTE": CINZA_TXT, "JUSTIFICADO": AZUL_CORP, "EM_REVISAO": AZUL_CORP}
    for gestor in sorted(df["gestor"].unique()):
        sub = df[df["gestor"] == gestor]
        with st.expander(f"{gestor} — {len(sub)} justificativa(s)"):
            for cr in sorted(sub["cr"].unique()):
                st.markdown(f"**{cr}**")
                linhas2 = ""
                for _, r in sub[sub["cr"] == cr].iterrows():
                    st_lab = STATUS_LABEL.get(r["status"])
                    cor = st_cor.get(r["status"], AZUL_CORP)
                    txt = (r["texto"][:120] + "…") if len(r["texto"]) > 120 else (r["texto"] or "—")
                    linhas2 += f"<tr><td>{r['conta']}</td><td>{chip(st_lab, cor)}</td><td style='text-align:left'>{txt}</td></tr>"
                st.markdown(f"""<table class="lle"><tr><th>Conta</th><th>Situação</th>
                    <th style='text-align:left'>Justificativa</th></tr>{linhas2}</table>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- administração
def tela_admin(c, prof):
    st.markdown("<div class='modtag'>Administração de acessos</div>", unsafe_allow_html=True)
    st.markdown("<div class='modsub'>Substituição por desligamento, ativação e transferência de centros</div>", unsafe_allow_html=True)
    st.caption("Substitua o e-mail de um gestor desligado, ative/desative acessos e transfira centros de resultado. "
               "O histórico de justificativas é sempre preservado.")

    gestores = c.table("gestor").select("codigo, nome, papel").order("nome").execute().data or []
    usuarios = c.table("gestor_usuario").select("email, gestor_codigo, papel_acesso, ativo").execute().data or []
    nome_por_cod = {g["codigo"]: g["nome"] for g in gestores}
    op_gestor = {f"{g['nome']} ({g['codigo']})": g["codigo"] for g in gestores}

    aba = st.radio("Ação", ["Substituir e-mail (desligamento)", "Ativar / desativar acesso", "Transferir centro de resultado"], horizontal=True)

    if aba.startswith("Substituir"):
        st.markdown("###### Substituir o e-mail de um gestor")
        st.caption("O e-mail antigo é desativado (mantém o histórico) e o novo passa a acessar os mesmos centros de resultado.")
        g_sel = st.selectbox("Gestor (função)", list(op_gestor.keys()))
        cod = op_gestor[g_sel]
        atuais = [u for u in usuarios if u["gestor_codigo"] == cod]
        if atuais:
            st.write("Acessos atuais desta função:")
            st.markdown("<table class='lle'><tr><th>E-mail</th><th>Tipo</th><th>Situação</th></tr>" +
                "".join(f"<tr><td>{u['email']}</td><td>{u['papel_acesso']}</td><td>{'Ativo' if u['ativo'] else 'Inativo'}</td></tr>" for u in atuais) +
                "</table>", unsafe_allow_html=True)
        email_antigo = st.selectbox("E-mail a desativar (do desligado)", ["(nenhum)"] + [u["email"] for u in atuais if u["ativo"]])
        email_novo = st.text_input("Novo e-mail (do substituto)", placeholder="novo.gestor@grupolle.com.br").strip().lower()
        st.info("Depois de salvar aqui, crie o login do novo e-mail no Supabase (Authentication → Users) com uma senha provisória.")
        if st.button("Aplicar substituição", type="primary"):
            if not email_novo or "@" not in email_novo:
                st.error("Informe um e-mail novo válido.")
            else:
                if email_antigo != "(nenhum)":
                    c.table("gestor_usuario").update({"ativo": False}).eq("email", email_antigo).execute()
                c.table("gestor_usuario").upsert({"email": email_novo, "gestor_codigo": cod, "papel_acesso": "titular", "ativo": True, "senha_provisoria": True}, on_conflict="email").execute(); limpar_cache()
                st.success(f"Feito: {email_novo} agora responde por {nome_por_cod.get(cod, cod)}." + (f" {email_antigo} foi desativado." if email_antigo != '(nenhum)' else ""))
                st.rerun()

    elif aba.startswith("Ativar"):
        st.markdown("###### Ativar ou desativar um acesso")
        if not usuarios:
            st.info("Nenhum usuário cadastrado.")
        else:
            for u in sorted(usuarios, key=lambda x: (nome_por_cod.get(x["gestor_codigo"], ""), x["email"])):
                col1, col2, col3 = st.columns([3, 2, 1.2])
                col1.write(f"**{u['email']}**")
                col2.write(f"{nome_por_cod.get(u['gestor_codigo'], u['gestor_codigo'])} · {'Ativo' if u['ativo'] else 'Inativo'}")
                novo_estado = not u["ativo"]
                rot = "Reativar" if not u["ativo"] else "Desativar"
                if col3.button(rot, key=f"tog_{u['email']}"):
                    c.table("gestor_usuario").update({"ativo": novo_estado}).eq("email", u["email"]).execute(); limpar_cache()
                    st.rerun()

    else:
        st.markdown("###### Transferir um centro de resultado para outro gestor")
        st.caption("Muda o dono do CR. As justificativas já feitas continuam registradas na conta/CR, sem alteração.")
        crs = c.table("cr_gestor").select("uni_cod, cr_cod, cr_nome, gestor_codigo").order("cr_nome").execute().data or []
        op_cr = {f"{r['cr_nome']} (uni {r['uni_cod']} · CR {r['cr_cod']}) — hoje: {nome_por_cod.get(r['gestor_codigo'], r['gestor_codigo'])}": (r["uni_cod"], r["cr_cod"]) for r in crs}
        cr_sel = st.selectbox("Centro de resultado", list(op_cr.keys()))
        destino = st.selectbox("Novo gestor responsável", list(op_gestor.keys()))
        if st.button("Transferir CR", type="primary"):
            uni, crc = op_cr[cr_sel]
            c.table("cr_gestor").update({"gestor_codigo": op_gestor[destino]}).eq("uni_cod", uni).eq("cr_cod", crc).execute(); limpar_cache()
            st.success(f"CR transferido para {destino.split(' (')[0]}.")
            st.rerun()

# ---------------------------------------------------------------- acompanhamento
def tela_acompanhamento(c, prof, banda, df_orc, cg, is_ctrl):
    st.markdown("<div class='modtag'>Módulo Acompanhamento de Despesas — Orçado x Realizado</div>", unsafe_allow_html=True)

    if df_orc.empty:
        st.info("Nenhum dado carregado ainda." + (" Use a aba 'Importar dados'." if is_ctrl else " Fale com a controladoria."))
        return
    df = df_orc.copy()
    if is_ctrl and cg:
        df["_resp"] = df.apply(lambda r: cg.get((int(r["uni_cod"]), int(r["cr_cod"])), ("—", ""))[0], axis=1)

    # ---------- filtros horizontais ----------
    ncols = 5 if is_ctrl else 4
    fcols = st.columns([1.1, 1.5, 1.2, 1.7, 1.7] if is_ctrl else [1.1, 1.2, 1.7, 1.7])
    i = 0
    mes = fcols[i].selectbox("Mês", list(range(1, 13)), index=5, format_func=lambda m: f"{MESES[m]}/2026", key="acomp_mes"); i += 1
    if is_ctrl:
        resps = ["Todos"] + sorted([r for r in df["_resp"].dropna().unique().tolist() if r])
        f_resp = fcols[i].selectbox("Gestor", resps, key="acomp_gestor"); i += 1
        if f_resp != "Todos": df = df[df["_resp"] == f_resp]
    unis = ["Todas"] + sorted(df["unidade"].dropna().unique().tolist())
    f_uni = fcols[i].selectbox("Unidade", unis, key="acomp_uni"); i += 1
    if f_uni != "Todas": df = df[df["unidade"] == f_uni]
    crs = ["Todos"] + sorted(df["cr_nome"].dropna().unique().tolist())
    f_cr = fcols[i].selectbox("Centro de resultado", crs, key="acomp_cr"); i += 1
    if f_cr != "Todos": df = df[df["cr_nome"] == f_cr]
    contas = ["Todas"] + sorted(df["conta_desc"].dropna().unique().tolist())
    f_conta = fcols[i].selectbox("Conta", contas, key="acomp_conta")
    if f_conta != "Todas": df = df[df["conta_desc"] == f_conta]

    # consolida empresas do mesmo CR+conta (net PISA + KING) por competência
    chaves = ["mes", "cr_cod", "conta_cod"]
    agg = {"valor_planejado": "sum", "valor_realizado": "sum",
           "unidade": "first", "uni_cod": "min", "cr_nome": "first",
           "conta_desc": "first", "cr_grupo": "first", "tipo_conta": "first", "classificacao": "first"}
    df = df.groupby(chaves, as_index=False).agg(agg)

    d_mes = df[df["mes"] == mes]
    d_ytd = df[df["mes"] <= mes]

    # aviso do gestor
    if not is_ctrl and mes < get_cobranca(c):
        st.info(f"{MESES[mes]}/2026 não está sujeito à cobrança de justificativa.")
    elif not is_ctrl:
        js = carregar_justificativas(2026, mes)
        enviadas = {(j["uni_cod"], j["cr_cod"], j["conta_cod"]) for j in js if j.get("status") in ("JUSTIFICADO", "EM_REVISAO", "APROVADO")}
        pend = 0
        for _, v in d_mes.iterrows():
            raw, pct = var_de(v["valor_planejado"], v["valor_realizado"])
            lab, _ = classifica(raw, pct, v["conta_cod"], banda)
            if lab == "Desfavorável" and (v["uni_cod"], v["cr_cod"], v["conta_cod"]) not in enviadas:
                pend += 1
        if pend:
            st.warning(f"Você tem {pend} conta(s) a justificar em {MESES[mes]}/2026 — veja a seção Justificativas ao final.")
        else:
            st.success(f"Nenhuma justificativa pendente em {MESES[mes]}/2026.")

    # ---------- resumo em 2 colunas ----------
    resumo_colunas(d_mes, d_ytd, banda, mes)
    contadores(d_mes, banda)
    st.caption("Convenção: contas de receita/dedução (código 3 ou 6) têm sinal invertido — a variação mede o IMPACTO no resultado. "
               "Verde = favorável, vermelho = desfavorável, cinza = dentro da faixa neutra (±"
               + f"{banda:.1f}".replace(".", ",") + "%).")

    st.divider()
    st.markdown("#### Evolução mensal — Jan a Dez (mês e acumulado YTD)")
    tabela_evolucao(df, banda, mes)

    st.divider()
    st.markdown(f"#### Desvios por centro de resultado — {MESES[mes]}/2026")
    drill_desvios(d_mes, banda, mes)

    st.divider()
    st.markdown(f"#### Justificativas · {MESES[mes]}/2026")
    secao_justificativas(c, prof, d_mes, mes, is_ctrl, banda)

# ---------------------------------------------------------------- main
def render_app(c, prof):
    """Renderiza o app autenticado (cabeçalho, barra, abas/telas e rodapé)."""
    is_ctrl = prof["papel"] == "controladoria"
    banda = get_faixa(c)

    header(prof)

    # barra superior: faixa vigente + sair
    tb = st.columns([6, 2.2, 1.1])
    tb[1].markdown(
        f"<div style='text-align:right; padding-top:6px; color:{CINZA_TXT}; font-size:12px;'>"
        f"Faixa neutra vigente: <b style='color:{AZUL_PROFUNDO}'>±{banda:.1f}%</b>".replace(".", ",")
        + (f" · Cobrança desde {MESES[get_cobranca(c)]}" if is_ctrl else "") + "</div>",
        unsafe_allow_html=True)
    if tb[2].button("Sair", key="sair_top", use_container_width=True):
        for k in ("access_token", "refresh_token", "email"): st.session_state.pop(k, None)
        st.rerun()

    # dados compartilhados (cacheados por token — releituras idênticas ficam instantâneas)
    orc = carregar_orc(2026)
    df_orc = pd.DataFrame(orc) if orc else pd.DataFrame()
    cg = {}
    try:
        for r in carregar_cr_gestor():
            cg[(int(r["uni_cod"]), int(r["cr_cod"]))] = ((r.get("gestor") or {}).get("nome", "—"), r.get("cr_nome", ""))
    except Exception:
        cg = {}

    if is_ctrl:
        t1, t2, t3, t4 = st.tabs(["📊 Acompanhamento", "📥 Justificativas recebidas",
                                  "🔑 Administração de acessos", "⬆️ Importar dados"])
        with t1: tela_acompanhamento(c, prof, banda, df_orc, cg, is_ctrl)
        with t2: tela_painel(c, prof, banda, df_orc, cg)
        with t3: tela_admin(c, prof)
        with t4: tela_importar(c)
    else:
        tela_acompanhamento(c, prof, banda, df_orc, cg, is_ctrl)

    rodape()

def main():
    inject_css()
    if not URL or not ANON:
        st.error("Faltam os segredos SUPABASE_URL e SUPABASE_ANON_KEY."); return

    # Âncora única: login e app ocupam o MESMO espaço na árvore.
    # Assim, ao autenticar, o app substitui a tela de login sem deixar "fantasma".
    slot = st.empty()

    if "access_token" not in st.session_state:
        with slot.container():
            tela_login()
        return

    c = client(); prof = perfil(c, st.session_state.get("email", ""))
    if not prof:
        with slot.container():
            st.error("Seu e-mail não está cadastrado. Fale com a controladoria.")
            if st.button("Sair"): st.session_state.clear(); st.rerun()
        return
    if prof.get("senha_provisoria"):
        with slot.container():
            tela_trocar_senha(c, st.session_state.get("email", ""))
        return

    with slot.container():
        render_app(c, prof)

main()
