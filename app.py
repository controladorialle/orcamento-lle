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

st.set_page_config(page_title="Sistema de Acompanhamento Orçamentário", page_icon="📊", layout="wide",
                   initial_sidebar_state="collapsed")
URL = st.secrets.get("SUPABASE_URL", ""); ANON = st.secrets.get("SUPABASE_ANON_KEY", "")
MESES = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
MABREV = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
STATUS_LABEL = {"PENDENTE": "Pendente", "JUSTIFICADO": "Justificado", "EM_REVISAO": "Em revisão", "DEVOLVIDO": "Devolvido", "APROVADO": "Aprovado"}
CATEGORIAS = [("SALARIOS", "Salários e ordenados"), ("ENCARGOS", "Encargos"), ("BENEFICIOS", "Benefícios"), ("OUTROS", "Outros de pessoal")]
CAT_LABEL = dict(CATEGORIAS)
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# Mapeamento das rubricas do template Treasy nas 4 categorias
HC_MAP = {
    "SALARIOS": ["SALARIO_TOTAL", "PRO_LABORE", "HE", "ADICIONAIS", "BONIFICACOES_GRATIFICACOES",
                 "BOLSA_ESTAGIO", "JOVEM_APRENDIZ", "QUEBRA_DE_CAIXA", "PROVISAO_FERIAS",
                 "PROVISAO_13_SALARIO", "AJUDA_DE_CUSTO", "AUXILIO_EDUCACAO", "COMISSAO"],
    "ENCARGOS": ["INSS_SALARIOS", "FGTS_SALARIOS", "FGTS_RESCISORIO", "INSS_SOBRE_PROVISAO_FERIAS",
                 "FGTS_SOBRE_PROVISAO_FERIAS", "INSS_SOBRE_PROVISAO_13_SALARIO", "FGTS_SOBRE_PROVISAO_13_SALARIO"],
    "BENEFICIOS": ["PROGRAMA_ALIMENTACAO", "ASSISTENCIA_MEDICA_ODONTO", "VALE_TRANSPORTE",
                   "VALE_COMBUSTIVEL", "QUALIDADE_DE_VIDA", "CESTA_BASICA"],
    "OUTROS": ["INDENIZACOES_AVISO", "TREINAMENTO", "UNIFORME_EPI", "EXAMES_MEDICOS", "CONTRIBUICAO_SINDICAL",
               "ABONO_PECUNIARIO_SINDICATO", "PLR", "EVENTOS_INTERNOS", "CONDUCOES_TRANSPORTES",
               "SEGURO_DE_VIDA", "DESPESAS_ENDOMARKETING"],
}
HC_COLS_DIM = ["ANO", "MES", "CODIGO_UNIDADE_NEGOCIO", "DESCRICAO_UNIDADE_NEGOCIO", "CODIGO_CENTRO_RESULTADO",
               "DESCRICAO_CENTRO_RESULTADO", "CODIGO_CARGO_FUNCIONARIO", "DESCRICAO_CARGO_FUNCIONARIO", "QUANTIDADE_FUNCIONARIOS"]

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
        padding:13px 16px; box-shadow:0 2px 8px rgba(7,22,57,.05); text-align:center; }}
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

def _toggle(k):
    st.session_state[k] = not st.session_state.get(k, False)

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

@st.cache_data(ttl=120, show_spinner=False)
def _q_justif_ano(tok, rtok, ano):
    cc = _cli_tok(tok, rtok)
    return cc.table("justificativa").select("*").eq("ano", ano).execute().data or []
def carregar_justificativas_ano(ano): return _q_justif_ano(*_tok(), ano)

@st.cache_data(ttl=300, show_spinner=False)
def _q_hc_quadro(tok, rtok, ano):
    cc = _cli_tok(tok, rtok)
    try: return cc.table("hc_quadro").select("*").eq("ano", ano).execute().data or []
    except Exception: return []

@st.cache_data(ttl=300, show_spinner=False)
def _q_hc_custo(tok, rtok, ano):
    cc = _cli_tok(tok, rtok)
    try: return cc.table("hc_custo").select("*").eq("ano", ano).execute().data or []
    except Exception: return []

def carregar_hc_quadro(ano): return _q_hc_quadro(*_tok(), ano)
def carregar_hc_custo(ano): return _q_hc_custo(*_tok(), ano)

@st.cache_data(ttl=60, show_spinner=False)
def _q_orc_log(tok, rtok, limite):
    cc = _cli_tok(tok, rtok)
    try: return cc.table("orc_log").select("*").order("alterado_em", desc=True).limit(limite).execute().data or []
    except Exception: return []
def carregar_orc_log(limite=200): return _q_orc_log(*_tok(), limite)

@st.cache_data(ttl=60, show_spinner=False)
def _q_hc_log(tok, rtok, limite):
    cc = _cli_tok(tok, rtok)
    try: return cc.table("hc_log").select("*").order("alterado_em", desc=True).limit(limite).execute().data or []
    except Exception: return []
def carregar_hc_log(limite=200): return _q_hc_log(*_tok(), limite)

@st.cache_data(ttl=120, show_spinner=False)
def _q_receita(tok, rtok, ano):
    cc = _cli_tok(tok, rtok)
    try: return cc.table("receita_venda").select("*").eq("ano", ano).execute().data or []
    except Exception: return []
def carregar_receita(ano=2026): return _q_receita(*_tok(), ano)

@st.cache_data(ttl=60, show_spinner=False)
def _q_receita_log(tok, rtok, limite):
    cc = _cli_tok(tok, rtok)
    try: return cc.table("receita_log").select("*").order("alterado_em", desc=True).limit(limite).execute().data or []
    except Exception: return []
def carregar_receita_log(limite=200): return _q_receita_log(*_tok(), limite)

@st.cache_data(ttl=120, show_spinner=False)
def _q_cmv(tok, rtok, ano):
    cc = _cli_tok(tok, rtok)
    try: return cc.table("cmv_valor").select("*").eq("ano", ano).execute().data or []
    except Exception: return []
def carregar_cmv(ano=2026): return _q_cmv(*_tok(), ano)

@st.cache_data(ttl=60, show_spinner=False)
def _q_cmv_log(tok, rtok, limite):
    cc = _cli_tok(tok, rtok)
    try: return cc.table("cmv_log").select("*").order("alterado_em", desc=True).limit(limite).execute().data or []
    except Exception: return []
def carregar_cmv_log(limite=200): return _q_cmv_log(*_tok(), limite)

def limpar_cache():
    """Limpeza total — usar em importações (mudam orçado e operacional)."""
    try: st.cache_data.clear()
    except Exception: pass

def limpar_cache_justif():
    """Limpeza cirúrgica após ações de justificativa: NÃO derruba orçado/operacional (que são pesados)."""
    for f in (_q_justif, _q_justif_ano):
        try: f.clear()
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
            <h1 style="color:{AZUL_PROFUNDO}; margin:12px 0 2px;">Sistema de Acompanhamento Orçamentário</h1>
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
          <h1>Sistema de Acompanhamento Orçamentário — LLE Ferragens</h1>
          <p>{prof['nome']}</p></div></div>
        <div class="lle-badge">GRUPO LLE <span>—</span> {papel}</div></div>""", unsafe_allow_html=True)

def rodape():
    st.markdown(f"""<div class="lle-foot">
        <div style="display:flex; align-items:center; gap:12px;">{LOGO}
          <div><div class="t">Sistema de Acompanhamento Orçamentário</div>
          <div class="s">Controladoria · Grupo LLE Ferragens</div></div></div>
        <div class="v"><div>Módulo Despesas · Orçado x Realizado</div>
          <div style="opacity:.7; margin-top:3px;">Todos os direitos reservados</div></div></div>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- blocos de insight
def resumo_colunas(d_mes, d_ytd, banda, mes, ano):
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
        + panel("MÊS", MESES[mes] + f"/{ano}", vp, vr, raw, pct, lab, cor)
        + panel("ACUMULADO YTD", f"Jan–{MABREV[mes]}/{ano}", yp, yr, yraw, ypct, ylab, ycor)
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
    st.caption("Clique em \u25b6 para abrir o orçado por conta e empresa do mês. YTD = acumulado até o mês.")
    ch1, ch2 = st.columns([0.05, 0.95])
    ch2.markdown("""<div class="drow head"><div class="nm">Mês</div>
        <div class="r">Orçado</div><div class="r">Realizado</div><div class="r">Var. (R$)</div>
        <div class="r">Var. (%)</div><div class="r">Status</div></div>""", unsafe_allow_html=True)

    for m in range(1, 13):
        vp = float(g.loc[m, "valor_planejado"]); vr = float(g.loc[m, "valor_realizado"])
        raw, pct = var_de(vp, vr)
        if vr == 0:
            cor = CINZA_TXT; vr_txt = "\u2014"; var_txt = "\u2014"; pctv = "\u2014"; status = chip("Sem realizado", CINZA_TXT)
        else:
            lab, cor = classifica(raw, pct, "5", banda)
            vr_txt = brl(vr); var_txt = brl(raw); pctv = pct_txt(pct); status = chip(lab, cor)
        key = f"evo_open_{m}"
        exp = st.session_state.get(key, False)
        cbtn, cbody = st.columns([0.05, 0.95])
        with cbtn:
            st.button("\u25bc" if exp else "\u25b6", key=f"btn_{key}", on_click=_toggle, args=(key,))
        exp = st.session_state.get(key, False)
        with cbody:
            bg = " style=\'background:#FFF7E6\'" if m == mes_sel else ""
            st.markdown(f"""<div class="drow"{bg}><div class="nm">{MESES[m]}</div>
                <div class="r">{brl(vp)}</div><div class="r">{vr_txt}</div>
                <div class="r" style="color:{cor};font-weight:600">{var_txt}</div>
                <div class="r" style="color:{cor}">{pctv}</div>
                <div class="r">{status}</div></div>""", unsafe_allow_html=True)
            if exp:
                yvp = float(cum.loc[m, "valor_planejado"]); yvr = float(cum.loc[m, "valor_realizado"])
                yraw, ypct = var_de(yvp, yvr)
                ycor = CINZA_TXT if yvr == 0 else (VERMELHO if yraw > 0 else VERDE)
                st.markdown(f"<div style='margin:2px 0 6px 8px;font-size:.85rem;color:{CINZA_TXT}'>"
                            f"Acumulado YTD (Jan\u2013{MESES[m]}): Orçado <b>{brl(yvp)}</b> · "
                            f"Realizado <b>{brl(yvr) if yvr else '\u2014'}</b> · "
                            f"Var. <b style='color:{ycor}'>{brl(yraw) if yvr else '\u2014'}</b> "
                            f"({pct_txt(ypct) if yvr else '\u2014'})</div>", unsafe_allow_html=True)
                det = df[df["mes"] == m].copy()
                if det.empty:
                    st.caption("Sem contas para os filtros selecionados neste mês.")
                else:
                    det = det.sort_values(["cr_nome", "conta_cod", "uni_cod"])
                    to = float(det["valor_planejado"].sum()); tr = float(det["valor_realizado"].sum()); tv = tr - to
                    linhas = ""
                    for r in det.itertuples():
                        p = float(r.valor_planejado or 0); q = float(r.valor_realizado or 0); rw = q - p
                        co = CINZA_TXT if q == 0 else (VERMELHO if rw > 0 else VERDE)
                        linhas += (f"<tr><td style='text-align:left'>{int(r.cr_cod)} · {r.cr_nome}</td>"
                                   f"<td style='text-align:left'>{int(r.conta_cod)} · {r.conta_desc}</td>"
                                   f"<td style='text-align:center'>{r.unidade}</td>"
                                   f"<td>{brl(p)}</td><td>{brl(q) if q else '\u2014'}</td>"
                                   f"<td style='color:{co}'>{brl(rw) if q else '\u2014'}</td></tr>")
                    tvcor = CINZA_TXT if tr == 0 else (VERMELHO if tv > 0 else VERDE)
                    total = (f"<tr class='mark'><td style='text-align:left' colspan='3'><b>Total ({len(det)} linha(s))</b></td>"
                             f"<td><b>{brl(to)}</b></td><td><b>{brl(tr) if tr else '\u2014'}</b></td>"
                             f"<td style='color:{tvcor}'><b>{brl(tv) if tr else '\u2014'}</b></td></tr>")
                    st.markdown(f"""<div class='scroll' style='margin:2px 0 12px 8px;'><table class="lle"><tr>
                        <th style='text-align:left'>Centro de resultado</th><th style='text-align:left'>Conta</th>
                        <th style='text-align:center'>Empresa</th><th>Orçado</th><th>Realizado</th><th>Var. (R$)</th></tr>
                        {linhas}{total}</table></div>""", unsafe_allow_html=True)

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
        exp = st.session_state.get(key, False)
        cbtn, cbody = st.columns([0.05, 0.95])
        with cbtn:
            st.button("▼" if exp else "▶", key=f"btn_{key}", on_click=_toggle, args=(key,))
        exp = st.session_state.get(key, False)
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
                    _emp = v.get("unidade", "") or ""
                    _cta = f"{v['conta_cod']} · {v.get('conta_desc','')}" + (f" ({_emp})" if _emp else "")
                    det.append((_cta, v["valor_planejado"], v["valor_realizado"], rw, pc, lb, co))
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
def secao_justificativas(c, prof, df_mes, mes, is_ctrl, banda, ano):
    if mes < get_cobranca(c):
        st.info(f"{MESES[mes]}/{ano} não está sujeito à cobrança de justificativa.")
        return
    js = carregar_justificativas(ano, mes)
    jmap = {(int(j["uni_cod"]), int(j["cr_cod"]), int(j["conta_cod"])): j for j in js}
    itens = []
    for _, v in df_mes.iterrows():
        raw, pct = var_de(v["valor_planejado"], v["valor_realizado"])
        lab, _ = classifica(raw, pct, v["conta_cod"], banda)
        if lab != "Desfavorável":
            continue
        j = jmap.get((int(v["uni_cod"]), int(v["cr_cod"]), int(v["conta_cod"])), {"status": "PENDENTE", "texto": "", "comentario_controladoria": ""})
        itens.append((v, raw, pct, j))
    itens.sort(key=lambda t: t[1], reverse=True)
    st_cor = {"APROVADO": VERDE, "DEVOLVIDO": VERMELHO, "PENDENTE": CINZA_TXT, "JUSTIFICADO": AZUL_CORP, "EM_REVISAO": AZUL_CORP}

    if not itens:
        st.success("Nenhum desvio desfavorável a justificar com os filtros atuais.")
        return

    # ----- filtro por situação (o gestor escolhe o que ver) -----
    GRUPOS = {"A responder": {"PENDENTE", "DEVOLVIDO"},
              "Aguardando controladoria": {"JUSTIFICADO", "EM_REVISAO"},
              "Aprovadas": {"APROVADO"}}
    def _st(it): return it[3].get("status", "PENDENTE")
    n_resp = sum(1 for it in itens if _st(it) in GRUPOS["A responder"])
    n_agu = sum(1 for it in itens if _st(it) in GRUPOS["Aguardando controladoria"])
    n_apr = sum(1 for it in itens if _st(it) in GRUPOS["Aprovadas"])
    opcoes = [f"Todas ({len(itens)})", f"A responder ({n_resp})",
              f"Aguardando controladoria ({n_agu})", f"Aprovadas ({n_apr})"]
    escolha = st.radio("Situação", opcoes, horizontal=True, key=f"just_filtro_{mes}")
    alvo = escolha.split(" (")[0]
    vis = itens if alvo == "Todas" else [it for it in itens if _st(it) in GRUPOS[alvo]]
    if not vis:
        st.caption(f"Exibindo 0 de {len(itens)} desvio(s) desfavorável(is) em {MESES[mes]}/{ano}")
        st.info("Nenhuma conta nesta situação.")
        return

    # paginação: desenhar centenas de expanders de uma vez deixava a página lenta
    PAGE = 25
    total = len(vis)
    npag = (total + PAGE - 1) // PAGE
    if npag > 1:
        pagina = int(st.number_input("Página", min_value=1, max_value=npag, value=1, step=1, key=f"just_pag_{mes}_{alvo}"))
        pagina = max(1, min(pagina, npag))
        ini = (pagina - 1) * PAGE
        page_vis = vis[ini:ini + PAGE]
        st.caption(f"Exibindo {ini + 1}–{ini + len(page_vis)} de {total} · página {pagina} de {npag} — dica: filtre por Gestor ou Centro de resultado acima para reduzir a lista.")
    else:
        page_vis = vis
        st.caption(f"Exibindo {total} de {len(itens)} desvio(s) desfavorável(is) em {MESES[mes]}/{ano}")

    # histórico só das contas desta página (evita carregar as ~50 mil notas do mês inteiro)
    contas_pg = sorted({int(v["conta_cod"]) for v, _, _, _ in page_vis})
    crs_pg = sorted({int(v["cr_cod"]) for v, _, _, _ in page_vis})
    try:
        _op = (c.table("operacional_detalhe").select("uni_cod,cr_cod,conta_cod,num_doc,valor,historico")
               .eq("ano", ano).eq("mes", mes).in_("cr_cod", crs_pg).in_("conta_cod", contas_pg).execute().data or [])
    except Exception:
        _op = []
    oper_all = pd.DataFrame(_op)
    for v, raw, pct, j in page_vis:
        status = j.get("status", "PENDENTE")
        titulo = f"{v['conta_cod']} · {v.get('conta_desc','')} — {v.get('cr_nome','')} ({v.get('unidade','')}) · {brl(raw)} · [{STATUS_LABEL.get(status)}]"
        with st.expander(titulo):
            a, b, d = st.columns(3)
            a.metric("Orçado", brl(v["valor_planejado"])); b.metric("Realizado", brl(v["valor_realizado"])); d.metric("Variação", brl(raw), pct_txt(pct))
            st.markdown(f"**Histórico das notas — {v.get('unidade', '')}**")
            det = ([] if oper_all.empty else oper_all[(oper_all["uni_cod"] == v["uni_cod"]) & (oper_all["cr_cod"] == v["cr_cod"]) & (oper_all["conta_cod"] == v["conta_cod"])]
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
            key = dict(ano=ano, mes=mes, uni_cod=int(v["uni_cod"]), cr_cod=int(v["cr_cod"]), conta_cod=int(v["conta_cod"]))
            kb = f"{mes}_{v['uni_cod']}_{v['cr_cod']}_{v['conta_cod']}"
            if status == "DEVOLVIDO" and j.get("comentario_controladoria"):
                st.warning(f"Controladoria: {j['comentario_controladoria']}")
            if not is_ctrl and status in ("PENDENTE", "DEVOLVIDO"):
                txt = st.text_area("Justificativa", value=j.get("texto", "") or "", key=f"txt_{kb}")
                c1, c2 = st.columns(2)
                if c1.button("Salvar rascunho", key=f"sv_{kb}"):
                    c.table("justificativa").upsert({**key, "texto": txt, "status": "PENDENTE", "atualizado_por": prof["nome"]}, on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute(); limpar_cache_justif(); st.rerun()
                if c2.button("Enviar justificativa", key=f"en_{kb}", type="primary"):
                    if not txt.strip(): st.error("Escreva a justificativa antes de enviar.")
                    else:
                        c.table("justificativa").upsert({**key, "texto": txt, "status": "JUSTIFICADO", "atualizado_por": prof["nome"]}, on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute(); limpar_cache_justif(); st.rerun()
            elif not is_ctrl:
                st.info(f"Justificativa: {j.get('texto') or '—'}"); st.caption("Aguardando a controladoria — não editável.")
            if is_ctrl:
                st.info(f"Justificativa do gestor: {j.get('texto') or '— (ainda não enviada)'}")
                if status in ("JUSTIFICADO", "EM_REVISAO"):
                    coment = st.text_input("Comentário (para devolução)", key=f"cm_{kb}")
                    c1, c2 = st.columns(2)
                    if c1.button("Aprovar", key=f"ap_{kb}", type="primary"):
                        c.table("justificativa").update({"status": "APROVADO"}).match(key).execute(); limpar_cache_justif(); st.rerun()
                    if c2.button("Devolver", key=f"dv_{kb}"):
                        c.table("justificativa").update({"status": "DEVOLVIDO", "comentario_controladoria": coment}).match(key).execute(); limpar_cache_justif(); st.rerun()

# ---------------------------------------------------------------- importar / config
def modelo_hc_treasy_xlsx():
    import io
    cols = HC_COLS_DIM + [col for cat in ("SALARIOS", "ENCARGOS", "BENEFICIOS", "OUTROS") for col in HC_MAP[cat]]
    ex1 = {c: 0 for c in cols}
    ex1.update({"ANO": 2026, "MES": 6, "CODIGO_UNIDADE_NEGOCIO": 1, "DESCRICAO_UNIDADE_NEGOCIO": "PISA",
                "CODIGO_CENTRO_RESULTADO": 102001, "DESCRICAO_CENTRO_RESULTADO": "CONTROLADORIA E AUDITORIA",
                "CODIGO_CARGO_FUNCIONARIO": 427, "DESCRICAO_CARGO_FUNCIONARIO": "ANALISTA DE CONTROLADORIA",
                "QUANTIDADE_FUNCIONARIOS": 1, "SALARIO_TOTAL": 5803, "INSS_SALARIOS": 1584.22, "FGTS_SALARIOS": 464.24})
    df = pd.DataFrame([ex1], columns=cols)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as x:
        df.to_excel(x, index=False, sheet_name="Planilha2")
    return buf.getvalue()

def tela_headcount(c, prof, ano, mes):
    st.markdown("<div class='modtag'>Gestão de Gastos com Pessoal</div>", unsafe_allow_html=True)
    st.markdown("<div class='modsub'>Quadro de pessoal e gastos com pessoal — orçado x realizado (Unidade × CR × Cargo)</div>", unsafe_allow_html=True)

    q = pd.DataFrame(carregar_hc_quadro(ano))
    k = pd.DataFrame(carregar_hc_custo(ano))
    if q.empty and k.empty:
        st.info("Nenhum dado de pessoal ainda. Vá à aba **Importar dados**, seção *Gestão de Gastos com Pessoal*, baixe o modelo e importe a planilha (padrão Treasy).")
        return
    for col in ("qtd_orcada", "qtd_realizada"):
        if not q.empty and col not in q.columns: q[col] = 0
    for col in ("valor_orcado", "valor_realizado"):
        if not k.empty and col not in k.columns: k[col] = 0

    base = pd.concat([q, k], ignore_index=True)
    def _opts_int(cod, nome):
        d = {}
        for _, r in base.iterrows():
            try: d[int(r[cod])] = str(r.get(nome, "") or int(r[cod]))
            except Exception: pass
        return d
    uni_map, cr_map = _opts_int("uni_cod", "unidade"), _opts_int("cr_cod", "cr_nome")
    cargo_map = {}
    for _, r in base.iterrows():
        cc = str(r.get("cargo_cod", "") or "")
        if cc: cargo_map[cc] = str(r.get("cargo_nome", "") or cc)

    fc = st.columns([1.4, 1.9, 1.6])
    uni_sel = fc[0].selectbox("Unidade", [0] + sorted(uni_map), format_func=lambda x: "Todas" if x == 0 else uni_map[x], key="hc_uni")
    cr_sel = fc[1].selectbox("Centro de resultado", [0] + sorted(cr_map), format_func=lambda x: "Todos" if x == 0 else f"{x} · {cr_map[x]}", key="hc_cr")
    cargo_sel = fc[2].selectbox("Cargo", [""] + sorted(cargo_map), format_func=lambda x: "Todos" if x == "" else f"{x} · {cargo_map[x]}", key="hc_cargo")

    def filtra(df, com_mes=True):
        if df.empty: return df
        d = df[df["mes"] == mes] if com_mes else df
        if uni_sel: d = d[d["uni_cod"] == uni_sel]
        if cr_sel: d = d[d["cr_cod"] == cr_sel]
        if cargo_sel: d = d[d["cargo_cod"].astype(str) == cargo_sel]
        return d
    qf, kf = filtra(q), filtra(k)

    hc_o = int(qf["qtd_orcada"].sum()) if not qf.empty else 0
    hc_r = int(qf["qtd_realizada"].sum()) if not qf.empty else 0
    custo_o = float(kf["valor_orcado"].sum()) if not kf.empty else 0.0
    custo_r = float(kf["valor_realizado"].sum()) if not kf.empty else 0.0
    d_hc = hc_r - hc_o
    d_custo_pct = (custo_r - custo_o) / custo_o * 100 if custo_o else 0.0
    medio = custo_r / hc_r if hc_r else 0.0
    cor_hc = CINZA_TXT if d_hc == 0 else (VERMELHO if d_hc > 0 else VERDE)
    cor_ct = CINZA_TXT if abs(d_custo_pct) < 0.05 else (VERMELHO if d_custo_pct > 0 else VERDE)

    cards = [("HC orçado", str(hc_o), AZUL_PROFUNDO), ("HC realizado", str(hc_r), AZUL_PROFUNDO),
             ("Δ headcount", f"{d_hc:+d}", cor_hc), ("Custo realizado", brl(custo_r), AZUL_PROFUNDO),
             ("Δ custo", pct_txt(d_custo_pct), cor_ct), ("Custo médio/func.", brl(medio), AZUL_PROFUNDO)]
    html = "<div class='cards'>"
    for lab, val, cor in cards:
        html += f"<div class='card'><div class='lab'>{lab}</div><div class='val' style='color:{cor}'>{val}</div></div>"
    st.markdown(html + "</div>", unsafe_allow_html=True)

    # quadro por cargo
    st.markdown("#### Quadro por cargo — orçado x realizado")
    hcg = (qf.groupby(["cargo_cod", "cargo_nome"], as_index=False).agg(o=("qtd_orcada", "sum"), r=("qtd_realizada", "sum"))
           if not qf.empty else pd.DataFrame(columns=["cargo_cod", "cargo_nome", "o", "r"]))
    custg = (kf.groupby("cargo_cod", as_index=False).agg(co=("valor_orcado", "sum"), cr=("valor_realizado", "sum"))
             if not kf.empty else pd.DataFrame(columns=["cargo_cod", "co", "cr"]))
    m = hcg.merge(custg, on="cargo_cod", how="outer") if not (hcg.empty and custg.empty) else pd.DataFrame()
    if not m.empty:
        m = m.fillna(0)
        m["cargo_nome"] = m["cargo_nome"].replace(0, "").astype(str)
        m = m.sort_values("cr", ascending=False)
        linhas = ""
        for _, r in m.iterrows():
            dhc = int(r["r"]) - int(r["o"])
            ch = CINZA_TXT if dhc == 0 else (VERMELHO if dhc > 0 else VERDE)
            dcp = (r["cr"] - r["co"]) / r["co"] * 100 if r["co"] else 0.0
            cc2 = CINZA_TXT if abs(dcp) < 0.05 else (VERMELHO if dcp > 0 else VERDE)
            nome = r["cargo_nome"] or f"Cargo {r['cargo_cod']}"
            linhas += (f"<tr><td style='text-align:left'>{r['cargo_cod']} · {nome}</td>"
                       f"<td style='text-align:center'>{int(r['o'])}</td><td style='text-align:center'>{int(r['r'])}</td>"
                       f"<td style='text-align:center; color:{ch}'>{dhc:+d}</td>"
                       f"<td>{brl(r['cr'])}</td><td style='text-align:center; color:{cc2}'>{pct_txt(dcp)}</td></tr>")
        st.markdown(f"""<table class="lle"><tr><th style='text-align:left'>Cargo</th>
            <th style='text-align:center'>HC orç.</th><th style='text-align:center'>HC real.</th>
            <th style='text-align:center'>Δ HC</th><th>Custo real.</th><th style='text-align:center'>Δ %</th></tr>{linhas}</table>""", unsafe_allow_html=True)
    else:
        st.caption("Sem lançamentos para os filtros selecionados.")

    # composição por categoria
    st.markdown("#### Composição do custo de pessoal por categoria")
    if not kf.empty:
        catg = kf.groupby("categoria", as_index=False).agg(o=("valor_orcado", "sum"), r=("valor_realizado", "sum"))
        cmap = {row["categoria"]: (row["o"], row["r"]) for _, row in catg.iterrows()}
        cols = st.columns(len(CATEGORIAS))
        for i, (cod, lab) in enumerate(CATEGORIAS):
            o, rr = cmap.get(cod, (0.0, 0.0))
            dcp = (rr - o) / o * 100 if o else 0.0
            cor = CINZA_TXT if abs(dcp) < 0.05 else (VERMELHO if dcp > 0 else VERDE)
            cols[i].markdown(
                f"<div style='border:1px solid {LINHA}; border-radius:12px; padding:12px'>"
                f"<div style='font-size:12px; color:{CINZA_TXT}'>{lab}</div>"
                f"<div style='font-size:16px; font-weight:700; color:{AZUL_PROFUNDO}'>{brl(rr)}</div>"
                f"<div style='font-size:11px; color:{cor}'>{pct_txt(dcp)} vs orçado</div></div>", unsafe_allow_html=True)
    else:
        st.caption("Sem custo de pessoal para os filtros selecionados.")

    # evolução mensal do headcount
    st.markdown("#### Evolução mensal do headcount")
    qy = filtra(q, com_mes=False)
    if not qy.empty:
        ev = qy.groupby("mes", as_index=False).agg(o=("qtd_orcada", "sum"), r=("qtd_realizada", "sum"))
        emap = {int(row["mes"]): (int(row["o"]), int(row["r"])) for _, row in ev.iterrows()}
        cab = "".join(f"<th style='text-align:center'>{MESES[mm][:3]}</th>" for mm in range(1, 13))
        lo = "".join(f"<td style='text-align:center'>{emap.get(mm,(0,0))[0]}</td>" for mm in range(1, 13))
        lr = "".join(f"<td style='text-align:center'>{emap.get(mm,(0,0))[1]}</td>" for mm in range(1, 13))
        st.markdown(f"""<table class="lle"><tr><th style='text-align:left'>&nbsp;</th>{cab}</tr>
            <tr><td style='text-align:left'>Orçado</td>{lo}</tr>
            <tr><td style='text-align:left'>Realizado</td>{lr}</tr></table>""", unsafe_allow_html=True)
    else:
        st.caption("Sem dados de quadro para os filtros selecionados.")

    # ---------- edição manual dos valores de pessoal (com log) ----------
    st.divider()
    st.markdown("#### Editar valores de pessoal")
    if not cr_sel:
        st.caption("Selecione uma **Unidade** e um **Centro de resultado** acima para editar quantidades e custos deste recorte.")
    elif qf.empty and kf.empty:
        st.caption("Nada para editar neste recorte.")
    else:
        st.caption("Ajuste headcount e custos por cargo. Toda alteração fica registrada no log.")
        # --- grade de headcount (quantidade) ---
        keys_q, disp_q, oqo, oqr = [], [], [], []
        if not qf.empty:
            qe = qf.sort_values("cargo_cod").reset_index(drop=True)
            for r in qe.itertuples():
                keys_q.append((int(r.ano), int(r.mes), int(r.uni_cod), int(r.cr_cod), str(r.cargo_cod)))
                oqo.append(round(float(getattr(r, "qtd_orcada", 0) or 0), 2))
                oqr.append(round(float(getattr(r, "qtd_realizada", 0) or 0), 2))
                disp_q.append(f"{r.cargo_cod} · {r.cargo_nome}")
            dq = pd.DataFrame({"Cargo": disp_q, "HC orçado": oqo, "HC realizado": oqr})
            st.markdown("**Quadro (quantidade de pessoas)**")
            edq = st.data_editor(dq, key=f"hce_q_{mes}_{uni_sel}_{cr_sel}", hide_index=True, use_container_width=True,
                                 num_rows="fixed", disabled=["Cargo"],
                                 column_config={"HC orçado": st.column_config.NumberColumn(format="%.0f", step=1),
                                                "HC realizado": st.column_config.NumberColumn(format="%.0f", step=1)})
        else:
            edq = None
        # --- grade de custo (por categoria) ---
        keys_k, disp_cargo, disp_cat, oko, okr = [], [], [], [], []
        if not kf.empty:
            ke = kf.sort_values(["cargo_cod", "categoria"]).reset_index(drop=True)
            for r in ke.itertuples():
                keys_k.append((int(r.ano), int(r.mes), int(r.uni_cod), int(r.cr_cod), str(r.cargo_cod), str(r.categoria)))
                oko.append(round(float(getattr(r, "valor_orcado", 0) or 0), 2))
                okr.append(round(float(getattr(r, "valor_realizado", 0) or 0), 2))
                disp_cargo.append(f"{r.cargo_cod} · {r.cargo_nome}")
                disp_cat.append(CAT_LABEL.get(str(r.categoria), str(r.categoria)))
            dk = pd.DataFrame({"Cargo": disp_cargo, "Categoria": disp_cat, "Orçado": oko, "Realizado": okr})
            st.markdown("**Custo por categoria**")
            edk = st.data_editor(dk, key=f"hce_k_{mes}_{uni_sel}_{cr_sel}", hide_index=True, use_container_width=True,
                                 num_rows="fixed", disabled=["Cargo", "Categoria"],
                                 column_config={"Orçado": st.column_config.NumberColumn(format="%.2f", step=0.01),
                                                "Realizado": st.column_config.NumberColumn(format="%.2f", step=0.01)})
        else:
            edk = None

        if st.button("Salvar alterações de pessoal", key="hce_save", type="primary"):
            mudou = 0
            if edq is not None:
                no, nr = list(edq["HC orçado"]), list(edq["HC realizado"])
                for i, kk in enumerate(keys_q):
                    an, me, uni, cr, cgo = kk
                    match = dict(ano=an, mes=me, uni_cod=uni, cr_cod=cr, cargo_cod=cgo)
                    for coluna, novos, orig in (("qtd_orcada", no, oqo), ("qtd_realizada", nr, oqr)):
                        try: nv = round(float(novos[i]), 2)
                        except (TypeError, ValueError): continue
                        if abs(nv - orig[i]) > 0.005:
                            c.table("hc_quadro").update({coluna: nv}).match(match).execute()
                            c.table("hc_log").insert(dict(**match, categoria=None, campo=coluna,
                                valor_antigo=orig[i], valor_novo=nv, alterado_por=prof.get("nome", ""))).execute()
                            mudou += 1
            if edk is not None:
                no, nr = list(edk["Orçado"]), list(edk["Realizado"])
                for i, kk in enumerate(keys_k):
                    an, me, uni, cr, cgo, cat = kk
                    match = dict(ano=an, mes=me, uni_cod=uni, cr_cod=cr, cargo_cod=cgo, categoria=cat)
                    for coluna, novos, orig in (("valor_orcado", no, oko), ("valor_realizado", nr, okr)):
                        try: nv = round(float(novos[i]), 2)
                        except (TypeError, ValueError): continue
                        if abs(nv - orig[i]) > 0.005:
                            c.table("hc_custo").update({coluna: nv}).match(match).execute()
                            c.table("hc_log").insert(dict(**match, campo=coluna,
                                valor_antigo=orig[i], valor_novo=nv, alterado_por=prof.get("nome", ""))).execute()
                            mudou += 1
            if mudou:
                limpar_cache()
                st.success(f"{mudou} alteração(ões) de pessoal salva(s) e registrada(s) no log.")
                st.rerun()
            else:
                st.info("Nenhuma alteração detectada.")

    # ---------- histórico de alterações de pessoal ----------
    log = carregar_hc_log(300)
    if log:
        st.markdown("###### Histórico de alterações de pessoal")
        CAMPO_LAB = {"qtd_orcada": "HC orçado", "qtd_realizada": "HC realizado",
                     "valor_orcado": "Custo orçado", "valor_realizado": "Custo realizado"}
        linhas = ""
        for g in log[:300]:
            quando = str(g.get("alterado_em", "") or "")[:16].replace("T", " ")
            va = float(g.get("valor_antigo") or 0); vn = float(g.get("valor_novo") or 0)
            seta = VERMELHO if vn > va else VERDE
            campo = g.get("campo", "")
            eh_qtd = campo in ("qtd_orcada", "qtd_realizada")
            de = (f"{va:.0f}" if eh_qtd else brl(va)); para = (f"{vn:.0f}" if eh_qtd else brl(vn))
            cat = CAT_LABEL.get(g.get("categoria") or "", "—") if not eh_qtd else "—"
            linhas += (f"<tr><td style='text-align:left'>{quando}</td><td style='text-align:left'>{g.get('alterado_por','') or '—'}</td>"
                       f"<td style='text-align:left'>{MESES[int(g.get('mes') or 1)]}</td>"
                       f"<td style='text-align:left'>{g.get('cargo_cod','')}</td>"
                       f"<td style='text-align:center'>{CAMPO_LAB.get(campo, campo)}</td>"
                       f"<td style='text-align:left'>{cat}</td>"
                       f"<td>{de}</td><td style='color:{seta}'>{para}</td></tr>")
        st.markdown(f"""<table class="lle"><tr>
            <th style='text-align:left'>Quando</th><th style='text-align:left'>Quem</th><th style='text-align:left'>Mês</th>
            <th style='text-align:left'>Cargo</th><th style='text-align:center'>Campo</th>
            <th style='text-align:left'>Categoria</th><th>De</th><th>Para</th></tr>{linhas}</table>""", unsafe_allow_html=True)

def tela_importar(c, ano):
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
    novo_mc = st.selectbox("Cobrar justificativas a partir do mês (vale para todos os anos)", list(range(1, 13)), index=mc - 1, format_func=lambda m: MESES[m])
    st.caption("Meses anteriores a este não geram pendência nem cobrança (continuam visíveis nas análises).")
    if st.button("Salvar início da cobrança"):
        set_cobranca(c, novo_mc); st.success(f"Cobrança a partir de {MESES[novo_mc]}.")
    st.divider()
    st.subheader("Importar dados")
    st.caption("A base orçado x realizado atualiza os números; o arquivo operacional traz o histórico das notas.")
    ano = st.number_input("Ano", value=int(ano), step=1)
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

    st.divider()
    st.subheader("Gestão de Gastos com Pessoal — planilha (padrão Treasy)")
    st.caption("Arquivo único no layout Treasy: uma linha por unidade × CR × cargo, com QUANTIDADE_FUNCIONARIOS e as "
               "rubricas de custo. O sistema soma as rubricas nas 4 categorias automaticamente. Importe o arquivo de "
               "Realizado e, se tiver, o de Orçado — cada tipo preenche a sua coluna, sem apagar o outro.")
    st.download_button("⬇️ Modelo — layout Treasy", data=modelo_hc_treasy_xlsx(),
                       file_name="modelo_headcount_treasy.xlsx", mime=XLSX_MIME, key="mdl_hc")

    def salvar_cargos(cargos):
        recs = [dict(cargo_cod=k, cargo_nome=v, ativo=True) for k, v in cargos.items() if k]
        for ch in chunks(recs):
            try: c.table("hc_cargo").upsert(ch, on_conflict="cargo_cod").execute()
            except Exception: pass

    def cod_cargo(x):
        if x is None or pd.isna(x): return None
        if isinstance(x, (int, float)):
            return str(int(x)) if float(x).is_integer() else str(x)
        return str(x).strip()

    tipo = st.radio("Este arquivo é", ["Realizado", "Orçado"], horizontal=True, key="hc_tipo")
    fh = st.file_uploader("Planilha de pessoal (layout Treasy)", type=["xlsx"], key="fh")
    if fh and st.button("Importar planilha de pessoal", key="imp_hc"):
        df = pd.read_excel(fh).dropna(how="all")
        cmap = {norm(col): col for col in df.columns}
        def cv(row, name):
            col = cmap.get(norm(name))
            return row[col] if col is not None else None
        realizado = tipo == "Realizado"
        qcol = "qtd_realizada" if realizado else "qtd_orcada"
        vcol = "valor_realizado" if realizado else "valor_orcado"
        quadro, custo, cargos = [], [], {}
        for _, r in df.iterrows():
            uni = toint(cv(r, "CODIGO_UNIDADE_NEGOCIO")); cr = toint(cv(r, "CODIGO_CENTRO_RESULTADO"))
            cgo = cod_cargo(cv(r, "CODIGO_CARGO_FUNCIONARIO"))
            if uni is None or cr is None or not cgo: continue
            an = toint(cv(r, "ANO")) or int(ano); me = toint(cv(r, "MES")) or 1
            cargo_nome = str(cv(r, "DESCRICAO_CARGO_FUNCIONARIO") or "")
            cargos[cgo] = cargo_nome
            dim = dict(ano=an, mes=me, uni_cod=uni, unidade=str(cv(r, "DESCRICAO_UNIDADE_NEGOCIO") or ""),
                       cr_cod=cr, cr_nome=str(cv(r, "DESCRICAO_CENTRO_RESULTADO") or ""),
                       cargo_cod=cgo, cargo_nome=cargo_nome)
            quadro.append({**dim, qcol: num(cv(r, "QUANTIDADE_FUNCIONARIOS"))})
            for cat, rubricas in HC_MAP.items():
                total = sum(num(cv(r, rub)) for rub in rubricas)
                custo.append({**dim, "categoria": cat, vcol: total})
        salvar_cargos(cargos)
        for ch in chunks(quadro):
            c.table("hc_quadro").upsert(ch, on_conflict="ano,mes,uni_cod,cr_cod,cargo_cod").execute()
        for ch in chunks(custo):
            c.table("hc_custo").upsert(ch, on_conflict="ano,mes,uni_cod,cr_cod,cargo_cod,categoria").execute()
        periodos = sorted({(x["ano"], x["mes"]) for x in quadro})
        per = ", ".join(f"{m:02d}/{a}" for a, m in periodos)
        st.success(f"{len(quadro)} cargo(s) importado(s) como {tipo}" + (f" — período(s): {per}." if per else "."))
        limpar_cache()

# ---------------------------------------------------------------- painel de recebidas
def relatorio_justificativas_xlsx(js_rows, df_orc, cg):
    """Monta um .xlsx das justificativas inseridas, enriquecido com orçado/realizado/variação."""
    import io
    fin = {}
    if df_orc is not None and not df_orc.empty:
        for _, r in df_orc.iterrows():
            fin[(int(r["mes"]), int(r["uni_cod"]), int(r["cr_cod"]), int(r["conta_cod"]))] = (
                float(r["valor_planejado"] or 0), float(r["valor_realizado"] or 0),
                r.get("unidade", "") or "", r.get("cr_nome", "") or "", r.get("conta_desc", "") or "")
    linhas = []
    for j in js_rows:
        mes = int(j["mes"]); uni = int(j["uni_cod"]); cr = int(j["cr_cod"]); ct = int(j["conta_cod"])
        vp, vr, unidade, cr_nome, conta_desc = fin.get((mes, uni, cr, ct), (0.0, 0.0, "", "", ""))
        raw, pct = var_de(vp, vr)
        gestor = cg.get((uni, cr), ("—", ""))[0]
        if not cr_nome:
            cr_nome = cg.get((uni, cr), ("", ""))[1]
        if not unidade:
            unidade = "PISA" if uni == 1 else "KING" if uni == 2 else f"Emp {uni}"
        linhas.append({
            "Ano": int(j.get("ano", 2026)), "Mês": MESES[mes], "Nº Mês": mes,
            "Gestor": gestor, "Cód. Unidade": uni, "Unidade": unidade,
            "Cód. CR": cr, "Centro de Resultado": cr_nome,
            "Cód. Conta": ct, "Descrição da Conta": conta_desc,
            "Orçado": round(vp, 2), "Realizado": round(vr, 2),
            "Variação (R$)": round(raw, 2), "Variação (%)": round(pct, 2),
            "Situação": STATUS_LABEL.get(j.get("status", "PENDENTE"), j.get("status", "")),
            "Justificativa": j.get("texto", "") or "",
            "Comentário Controladoria": j.get("comentario_controladoria", "") or "",
            "Atualizado por": j.get("atualizado_por", "") or "",
            "Atualizado em": str(j.get("atualizado_em", "") or ""),
        })
    df = pd.DataFrame(linhas)
    if not df.empty:
        df = df.sort_values(["Nº Mês", "Gestor", "Centro de Resultado", "Cód. Conta"], kind="stable").drop(columns=["Nº Mês"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Justificativas")
    return buf.getvalue(), len(df)

def tela_painel(c, prof, banda, df_orc, cg, ano, mes):
    st.markdown("<div class='modtag'>Justificativas recebidas</div>", unsafe_allow_html=True)
    st.markdown("<div class='modsub'>Pendências por gestor, resumo e detalhe das respostas</div>", unsafe_allow_html=True)

    # ----- Exportar relatório (.xlsx) -----
    with st.expander("📄 Exportar relatório de justificativas (.xlsx)"):
        escopo = st.radio("Período", [f"Mês selecionado ({MESES[mes]}/{ano})", f"Ano inteiro ({ano})"],
                          horizontal=True, key="rel_escopo")
        if st.button("Gerar relatório", key="rel_gen"):
            ano_todo = escopo.startswith("Ano")
            js_rel = carregar_justificativas_ano(ano) if ano_todo else carregar_justificativas(ano, mes)
            dados, nlin = relatorio_justificativas_xlsx(js_rel, df_orc, cg)
            st.session_state["rel_bytes"] = dados
            st.session_state["rel_nome"] = f"justificativas_{ano}.xlsx" if ano_todo else f"justificativas_{mes:02d}_{ano}.xlsx"
            st.session_state["rel_n"] = nlin
        if st.session_state.get("rel_bytes") is not None:
            if st.session_state.get("rel_n", 0) == 0:
                st.caption("Nenhuma justificativa inserida no período gerado.")
            else:
                st.download_button(f"⬇️ Baixar {st.session_state['rel_nome']} ({st.session_state['rel_n']} linha(s))",
                                   data=st.session_state["rel_bytes"], file_name=st.session_state["rel_nome"],
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="rel_dl")

    if mes < get_cobranca(c):
        st.info(f"{MESES[mes]}/{ano} não está sujeito à cobrança de justificativa.")
        return

    js = carregar_justificativas(ano, mes)

    # ----- Pendências por empresa (mesma régua do gestor: por unidade + CR + conta) -----
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
        d["itens"].append({"cr": v.get("cr_nome", ""), "conta": f"{v['conta_cod']} · {v.get('conta_desc','')} ({v.get('unidade','')})", "raw": raw, "dev": dev})

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
                    linhas += (f"<tr><td>{it['cr']}</td><td style='text-align:center'>{it['conta']}{tag}</td>"
                               f"<td style='text-align:center; color:{VERMELHO}'>{brl(it['raw'])}</td></tr>")
                st.markdown(f"""<table class="lle"><tr><th style='text-align:left'>Centro de resultado</th>
                    <th style='text-align:center'>Conta</th>
                    <th style='text-align:center'>Desvio (R$)</th></tr>{linhas}</table>""", unsafe_allow_html=True)
    st.divider()

    if not js:
        st.info(f"Ainda não há justificativas enviadas em {MESES[mes]}/{ano}. As pendências acima mostram o que falta.")
        return

    desc_map = dict(zip(df_orc["conta_cod"].astype(int), df_orc["conta_desc"])) if not df_orc.empty else {}
    linhas = []
    for j in js:
        gestor, cr_nome = cg.get((j["uni_cod"], j["cr_cod"]), ("—", str(j["cr_cod"])))
        linhas.append({"gestor": gestor, "cr": cr_nome, "cr_cod": int(j["cr_cod"]), "uni_cod": int(j["uni_cod"]),
                       "conta": j["conta_cod"], "conta_desc": desc_map.get(int(j["conta_cod"]), ""),
                       "status": j.get("status", "PENDENTE"), "texto": j.get("texto", "") or "",
                       "comentario": j.get("comentario_controladoria", "") or "",
                       "_k": (j["uni_cod"], j["cr_cod"], j["conta_cod"])})
    df = pd.DataFrame(linhas)

    st.markdown("###### Resumo por gestor")
    resumo = ""
    for gestor in sorted(df["gestor"].unique()):
        sub = df[df["gestor"] == gestor]
        cnt = {s: int((sub["status"] == s).sum()) for s in STATUS_LABEL}
        resumo += (f"<tr><td style='text-align:center'>{gestor}</td><td style='text-align:center'>{len(sub)}</td>"
                   f"<td style='text-align:center; color:{VERMELHO}'>{cnt['PENDENTE']+cnt['DEVOLVIDO']}</td>"
                   f"<td style='text-align:center; color:{AZUL_CORP}'>{cnt['JUSTIFICADO']+cnt['EM_REVISAO']}</td>"
                   f"<td style='text-align:center; color:{VERDE}'>{cnt['APROVADO']}</td></tr>")
    st.markdown(f"""<table class="lle"><tr><th style='text-align:center'>Gestor</th><th style='text-align:center'>Total</th>
        <th style='text-align:center'>A responder</th><th style='text-align:center'>Aguardando controladoria</th><th style='text-align:center'>Aprovadas</th></tr>{resumo}</table>""", unsafe_allow_html=True)

    st.markdown("###### Detalhe por gestor e centro de resultado")
    st_cor = {"APROVADO": VERDE, "DEVOLVIDO": VERMELHO, "PENDENTE": CINZA_TXT, "JUSTIFICADO": AZUL_CORP, "EM_REVISAO": AZUL_CORP}
    for gestor in sorted(df["gestor"].unique()):
        sub = df[df["gestor"] == gestor]
        with st.expander(f"{gestor} — {len(sub)} justificativa(s)"):
            for cr in sorted(sub["cr"].unique()):
                grp = sub[sub["cr"] == cr]
                crc = int(grp["cr_cod"].iloc[0])
                st.markdown(f"**{cr} · {crc}**")
                linhas2 = ""
                for _, r in grp.iterrows():
                    st_lab = STATUS_LABEL.get(r["status"])
                    cor = st_cor.get(r["status"], AZUL_CORP)
                    txt = (r["texto"][:120] + "…") if len(r["texto"]) > 120 else (r["texto"] or "—")
                    _emp = "PISA" if r["uni_cod"] == 1 else "KING" if r["uni_cod"] == 2 else f"Emp {r['uni_cod']}"
                    conta_disp = (f"{r['conta']} · {r['conta_desc']} ({_emp})" if r["conta_desc"] else f"{r['conta']} ({_emp})")
                    linhas2 += (f"<tr><td style='text-align:center'>{conta_disp}</td>"
                                f"<td style='text-align:center'>{chip(st_lab, cor)}</td>"
                                f"<td style='text-align:left'>{txt}</td></tr>")
                st.markdown(f"""<table class="lle"><tr><th style='text-align:center'>Conta</th>
                    <th style='text-align:center'>Situação</th>
                    <th style='text-align:left'>Justificativa</th></tr>{linhas2}</table>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- edição manual do orçado
def tela_editar_orcado(c, prof, df_orc, ano, mes):
    st.markdown("<div class='modtag'>Manutenção do Orçamento</div>", unsafe_allow_html=True)
    st.markdown("<div class='modsub'>Edite orçado e/ou realizado linha a linha — ideal para reclassificações no período, sem reimportar o arquivo. Toda alteração fica registrada no log.</div>", unsafe_allow_html=True)
    if df_orc is None or df_orc.empty:
        st.info("Base de orçado ainda não carregada. Importe na aba **Importar dados**.")
        return

    fc = st.columns([1.4, 1.9])
    unis = sorted({(int(r["uni_cod"]), r.get("unidade", "")) for _, r in df_orc.iterrows()})
    uni_sel = fc[0].selectbox("Unidade", [0] + [u[0] for u in unis],
                              format_func=lambda x: "Todas" if x == 0 else next((n for u, n in unis if u == x), str(x)), key="edo_uni")
    dcr = df_orc if not uni_sel else df_orc[df_orc["uni_cod"] == uni_sel]
    crs = sorted({(int(r["cr_cod"]), r.get("cr_nome", "")) for _, r in dcr.iterrows()})
    cr_sel = fc[1].selectbox("Centro de resultado", [0] + [x[0] for x in crs],
                             format_func=lambda x: "Todos" if x == 0 else f"{x} · {next((n for cc, n in crs if cc == x), '')}", key="edo_cr")

    view = df_orc[df_orc["mes"] == mes]
    if uni_sel: view = view[view["uni_cod"] == uni_sel]
    if cr_sel: view = view[view["cr_cod"] == cr_sel]
    if view.empty:
        st.info("Nenhuma conta para os filtros selecionados.")
    else:
        view = view.sort_values(["uni_cod", "cr_cod", "conta_cod"]).reset_index(drop=True)
        keys = [(int(r.ano), int(r.mes), int(r.uni_cod), int(r.cr_cod), int(r.conta_cod)) for r in view.itertuples()]
        orig_o = [round(float(r.valor_planejado or 0), 2) for r in view.itertuples()]
        orig_r = [round(float(r.valor_realizado or 0), 2) for r in view.itertuples()]
        disp = pd.DataFrame({
            "Unidade": [r.unidade for r in view.itertuples()],
            "Centro de resultado": [r.cr_nome for r in view.itertuples()],
            "Conta": [f"{int(r.conta_cod)} · {r.conta_desc}" for r in view.itertuples()],
            "Orçado": orig_o,
            "Realizado": orig_r,
        })
        st.caption(f"{len(disp)} conta(s) em {MESES[mes]}/{ano}. Edite **Orçado** e/ou **Realizado** e clique em salvar.")
        edited = st.data_editor(
            disp, key=f"edo_grid_{mes}_{uni_sel}_{cr_sel}", hide_index=True, use_container_width=True,
            num_rows="fixed", disabled=["Unidade", "Centro de resultado", "Conta"],
            column_config={"Orçado": st.column_config.NumberColumn(format="%.2f", step=0.01),
                           "Realizado": st.column_config.NumberColumn(format="%.2f", step=0.01)})
        if st.button("Salvar alterações", key="edo_save", type="primary"):
            novos_o = list(edited["Orçado"]); novos_r = list(edited["Realizado"])
            mudou = 0
            for i, kkey in enumerate(keys):
                an, me, uni, cr, ct = kkey
                match = dict(ano=an, mes=me, uni_cod=uni, cr_cod=cr, conta_cod=ct)
                for coluna, novos, orig in (("valor_planejado", novos_o, orig_o),
                                            ("valor_realizado", novos_r, orig_r)):
                    try: nv = round(float(novos[i]), 2)
                    except (TypeError, ValueError): continue
                    if abs(nv - orig[i]) > 0.005:
                        c.table("orc_realizado").update({coluna: nv}).match(match).execute()
                        c.table("orc_log").insert(dict(**match, campo=coluna, valor_antigo=orig[i],
                            valor_novo=nv, alterado_por=prof.get("nome", ""))).execute()
                        mudou += 1
            if mudou:
                limpar_cache()
                st.success(f"{mudou} alteração(ões) salva(s) e registrada(s) no log.")
                st.rerun()
            else:
                st.info("Nenhuma alteração detectada.")

    # ---------- histórico de alterações ----------
    st.divider()
    st.markdown("#### Histórico de alterações")
    log = carregar_orc_log(300)
    if not log:
        st.caption("Nenhuma alteração registrada ainda.")
        return
    nome = {}
    for _, r in df_orc.iterrows():
        nome[(int(r["uni_cod"]), int(r["cr_cod"]), int(r["conta_cod"]))] = (r.get("unidade", ""), r.get("cr_nome", ""), r.get("conta_desc", ""))
    linhas = ""
    for g in log[:300]:
        uni = int(g.get("uni_cod") or 0); cr = int(g.get("cr_cod") or 0); ct = int(g.get("conta_cod") or 0)
        un, crn, cd = nome.get((uni, cr, ct), ("", "", ""))
        quando = str(g.get("alterado_em", "") or "")[:16].replace("T", " ")
        va = float(g.get("valor_antigo") or 0); vn = float(g.get("valor_novo") or 0)
        seta = VERMELHO if vn > va else VERDE
        campo_lab = "Orçado" if g.get("campo") == "valor_planejado" else "Realizado"
        linhas += (f"<tr><td style='text-align:left'>{quando}</td><td style='text-align:left'>{g.get('alterado_por','') or '—'}</td>"
                   f"<td style='text-align:left'>{MESES[int(g.get('mes') or 1)]}</td><td style='text-align:left'>{un or uni}</td>"
                   f"<td style='text-align:left'>{cr} · {crn}</td><td style='text-align:left'>{ct} · {cd}</td>"
                   f"<td style='text-align:center'>{campo_lab}</td><td>{brl(va)}</td><td style='color:{seta}'>{brl(vn)}</td></tr>")
    st.markdown(f"""<table class="lle"><tr>
        <th style='text-align:left'>Quando</th><th style='text-align:left'>Quem</th><th style='text-align:left'>Mês</th>
        <th style='text-align:left'>Unidade</th><th style='text-align:left'>Centro de resultado</th>
        <th style='text-align:left'>Conta</th><th style='text-align:center'>Campo</th><th>De</th><th>Para</th></tr>{linhas}</table>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- receita bruta de vendas
def tela_receita(c, prof, ano):
    EMP = {1: "PISA", 2: "KING"}
    banda = get_faixa(c)
    st.markdown("<div class='modtag'>Receita Bruta de Vendas</div>", unsafe_allow_html=True)
    st.markdown("<div class='modsub'>Planejado x realizado por empresa e mês. Convenção de receita: realizado acima do planejado é favorável.</div>", unsafe_allow_html=True)

    rows = carregar_receita(ano)
    dfr = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["ano", "mes", "uni_cod", "unidade", "valor_planejado", "valor_realizado"])
    for col in ("valor_planejado", "valor_realizado"):
        if col not in dfr.columns: dfr[col] = 0.0

    emp = st.selectbox("Empresa", [0, 1, 2], format_func=lambda x: "Todas" if x == 0 else EMP[x], key="rec_emp")
    scope = dfr if not emp else dfr[dfr["uni_cod"] == emp]

    if scope.empty:
        g = pd.DataFrame(0.0, index=range(1, 13), columns=["valor_planejado", "valor_realizado"])
    else:
        g = scope.groupby("mes")[["valor_planejado", "valor_realizado"]].sum().reindex(range(1, 13), fill_value=0.0)
    tp = float(g["valor_planejado"].sum()); tr = float(g["valor_realizado"].sum())
    var = tr - tp; pct = (var / tp * 100) if tp else 0.0

    def cor_receita(v, pl):
        p = (v / pl * 100) if pl else 0.0
        if pl and abs(p) <= banda: return CINZA_TXT
        return VERDE if v >= 0 else VERMELHO

    # KPIs
    k = st.columns(4)
    kpi = [("Planejado (ano)", brl(tp), CINZA_TXT), ("Realizado (ano)", brl(tr), CINZA_TXT),
           ("Variação (R$)", brl(var), cor_receita(var, tp)), ("Variação (%)", pct_txt(pct), cor_receita(var, tp))]
    for col, (t, v, cr) in zip(k, kpi):
        col.markdown(f"<div class='card' style='text-align:center'><div style='font-size:.8rem;color:{CINZA_TXT}'>{t}</div>"
                     f"<div style='font-size:1.4rem;font-weight:700;color:{cr}'>{v}</div></div>", unsafe_allow_html=True)

    # evolução mensal
    st.markdown("#### Evolução mensal")
    linhas = ""
    for m in range(1, 13):
        vp = float(g.loc[m, "valor_planejado"]); vr = float(g.loc[m, "valor_realizado"])
        v = vr - vp; p = (v / vp * 100) if vp else 0.0
        if vr == 0 and vp == 0:
            vr_txt = "\u2014"; var_txt = "\u2014"; pctv = "\u2014"; status = chip("Sem dados", CINZA_TXT); cr = CINZA_TXT
        elif vr == 0:
            vr_txt = "\u2014"; var_txt = "\u2014"; pctv = "\u2014"; status = chip("Sem realizado", CINZA_TXT); cr = CINZA_TXT
        else:
            cr = cor_receita(v, vp)
            lab = "Neutro" if cr == CINZA_TXT else ("Favorável" if v >= 0 else "Desfavorável")
            vr_txt = brl(vr); var_txt = brl(v); pctv = pct_txt(p); status = chip(lab, cr)
        linhas += (f"<tr><td style='text-align:left'>{MESES[m]}</td><td>{brl(vp)}</td><td>{vr_txt}</td>"
                   f"<td style='color:{cr}'>{var_txt}</td><td style='color:{cr}'>{pctv}</td><td>{status}</td></tr>")
    tcor = cor_receita(var, tp)
    total = (f"<tr class='mark'><td style='text-align:left'><b>Total</b></td><td><b>{brl(tp)}</b></td>"
             f"<td><b>{brl(tr) if tr else '\u2014'}</b></td><td style='color:{tcor}'><b>{brl(var) if tr else '\u2014'}</b></td>"
             f"<td style='color:{tcor}'><b>{pct_txt(pct) if tr else '\u2014'}</b></td><td></td></tr>")
    st.markdown(f"""<div class='scroll'><table class="lle"><tr>
        <th style='text-align:left'>Mês</th><th>Planejado</th><th>Realizado</th>
        <th>Var. (R$)</th><th>Var. (%)</th><th>Status</th></tr>{linhas}{total}</table></div>""", unsafe_allow_html=True)

    # ---------- registrar / editar ----------
    st.divider()
    st.markdown("#### Registrar / editar receita")
    st.caption("Preencha Planejado e Realizado por empresa e mês. Toda alteração fica registrada no log.")
    emps = [1, 2] if not emp else [emp]
    idx = {(int(r["mes"]), int(r["uni_cod"])): r for _, r in dfr.iterrows()}
    keys, disp_mes, disp_emp, oplan, oreal = [], [], [], [], []
    for u in emps:
        for m in range(1, 13):
            r = idx.get((m, u))
            keys.append((m, u))
            disp_mes.append(MESES[m]); disp_emp.append(EMP[u])
            oplan.append(round(float(r["valor_planejado"]) if r is not None else 0.0, 2))
            oreal.append(round(float(r["valor_realizado"]) if r is not None else 0.0, 2))
    dedit = pd.DataFrame({"Mês": disp_mes, "Empresa": disp_emp, "Planejado": oplan, "Realizado": oreal})
    ed = st.data_editor(dedit, key=f"rec_grid_{emp}", hide_index=True, use_container_width=True,
                        num_rows="fixed", disabled=["Mês", "Empresa"],
                        column_config={"Planejado": st.column_config.NumberColumn(format="%.2f", step=0.01),
                                       "Realizado": st.column_config.NumberColumn(format="%.2f", step=0.01)})
    if st.button("Salvar receita", key="rec_save", type="primary"):
        np_, nr_ = list(ed["Planejado"]), list(ed["Realizado"])
        mudou = 0
        for i, (m, u) in enumerate(keys):
            try: p_novo = round(float(np_[i]), 2); r_novo = round(float(nr_[i]), 2)
            except (TypeError, ValueError): continue
            mud_p = abs(p_novo - oplan[i]) > 0.005
            mud_r = abs(r_novo - oreal[i]) > 0.005
            if not (mud_p or mud_r): continue
            c.table("receita_venda").upsert(dict(ano=ano, mes=m, uni_cod=u, unidade=EMP[u],
                valor_planejado=p_novo, valor_realizado=r_novo, atualizado_por=prof.get("nome", "")),
                on_conflict="ano,mes,uni_cod").execute()
            if mud_p:
                c.table("receita_log").insert(dict(ano=ano, mes=m, uni_cod=u, campo="valor_planejado",
                    valor_antigo=oplan[i], valor_novo=p_novo, alterado_por=prof.get("nome", ""))).execute(); mudou += 1
            if mud_r:
                c.table("receita_log").insert(dict(ano=ano, mes=m, uni_cod=u, campo="valor_realizado",
                    valor_antigo=oreal[i], valor_novo=r_novo, alterado_por=prof.get("nome", ""))).execute(); mudou += 1
        if mudou:
            limpar_cache()
            st.success(f"{mudou} alteração(ões) de receita salva(s) e registrada(s) no log.")
            st.rerun()
        else:
            st.info("Nenhuma alteração detectada.")

    # ---------- histórico ----------
    log = carregar_receita_log(300)
    if log:
        st.markdown("###### Histórico de alterações de receita")
        CL = {"valor_planejado": "Planejado", "valor_realizado": "Realizado"}
        ln = ""
        for glog in log[:300]:
            quando = str(glog.get("alterado_em", "") or "")[:16].replace("T", " ")
            va = float(glog.get("valor_antigo") or 0); vn = float(glog.get("valor_novo") or 0)
            seta = VERDE if vn > va else VERMELHO
            ln += (f"<tr><td style='text-align:left'>{quando}</td><td style='text-align:left'>{glog.get('alterado_por','') or '\u2014'}</td>"
                   f"<td style='text-align:left'>{MESES[int(glog.get('mes') or 1)]}</td>"
                   f"<td style='text-align:center'>{EMP.get(int(glog.get('uni_cod') or 0), glog.get('uni_cod'))}</td>"
                   f"<td style='text-align:center'>{CL.get(glog.get('campo'), glog.get('campo'))}</td>"
                   f"<td>{brl(va)}</td><td style='color:{seta}'>{brl(vn)}</td></tr>")
        st.markdown(f"""<table class="lle"><tr>
            <th style='text-align:left'>Quando</th><th style='text-align:left'>Quem</th><th style='text-align:left'>Mês</th>
            <th style='text-align:center'>Empresa</th><th style='text-align:center'>Campo</th><th>De</th><th>Para</th></tr>{ln}</table>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- CMV (custo da mercadoria vendida)
def tela_cmv(c, prof, ano):
    EMP = {1: "PISA", 2: "KING"}
    banda = get_faixa(c)
    st.markdown("<div class='modtag'>CMV — Custo da Mercadoria Vendida</div>", unsafe_allow_html=True)
    st.markdown("<div class='modsub'>Planejado x realizado por empresa e mês. Convenção de custo: realizado abaixo do planejado é favorável (gastou menos).</div>", unsafe_allow_html=True)

    rows = carregar_cmv(ano)
    dfr = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["ano", "mes", "uni_cod", "unidade", "valor_planejado", "valor_realizado"])
    for col in ("valor_planejado", "valor_realizado"):
        if col not in dfr.columns: dfr[col] = 0.0

    emp = st.selectbox("Empresa", [0, 1, 2], format_func=lambda x: "Todas" if x == 0 else EMP[x], key="cmv_emp")
    scope = dfr if not emp else dfr[dfr["uni_cod"] == emp]

    if scope.empty:
        g = pd.DataFrame(0.0, index=range(1, 13), columns=["valor_planejado", "valor_realizado"])
    else:
        g = scope.groupby("mes")[["valor_planejado", "valor_realizado"]].sum().reindex(range(1, 13), fill_value=0.0)
    tp = float(g["valor_planejado"].sum()); tr = float(g["valor_realizado"].sum())
    var = tr - tp; pct = (var / tp * 100) if tp else 0.0

    def cor_cmv(v, pl):  # custo: gastar menos que o previsto (v<=0) é favorável
        p = (v / pl * 100) if pl else 0.0
        if pl and abs(p) <= banda: return CINZA_TXT
        return VERDE if v <= 0 else VERMELHO

    # KPIs
    k = st.columns(4)
    kpi = [("Planejado (ano)", brl(tp), CINZA_TXT), ("Realizado (ano)", brl(tr), CINZA_TXT),
           ("Variação (R$)", brl(var), cor_cmv(var, tp)), ("Variação (%)", pct_txt(pct), cor_cmv(var, tp))]
    for col, (t, v, cr) in zip(k, kpi):
        col.markdown(f"<div class='card' style='text-align:center'><div style='font-size:.8rem;color:{CINZA_TXT}'>{t}</div>"
                     f"<div style='font-size:1.4rem;font-weight:700;color:{cr}'>{v}</div></div>", unsafe_allow_html=True)

    # evolução mensal
    st.markdown("#### Evolução mensal")
    linhas = ""
    for m in range(1, 13):
        vp = float(g.loc[m, "valor_planejado"]); vr = float(g.loc[m, "valor_realizado"])
        v = vr - vp; p = (v / vp * 100) if vp else 0.0
        if vr == 0 and vp == 0:
            vr_txt = "\u2014"; var_txt = "\u2014"; pctv = "\u2014"; status = chip("Sem dados", CINZA_TXT); cr = CINZA_TXT
        elif vr == 0:
            vr_txt = "\u2014"; var_txt = "\u2014"; pctv = "\u2014"; status = chip("Sem realizado", CINZA_TXT); cr = CINZA_TXT
        else:
            cr = cor_cmv(v, vp)
            lab = "Neutro" if cr == CINZA_TXT else ("Favorável" if v <= 0 else "Desfavorável")
            vr_txt = brl(vr); var_txt = brl(v); pctv = pct_txt(p); status = chip(lab, cr)
        linhas += (f"<tr><td style='text-align:left'>{MESES[m]}</td><td>{brl(vp)}</td><td>{vr_txt}</td>"
                   f"<td style='color:{cr}'>{var_txt}</td><td style='color:{cr}'>{pctv}</td><td>{status}</td></tr>")
    tcor = cor_cmv(var, tp)
    total = (f"<tr class='mark'><td style='text-align:left'><b>Total</b></td><td><b>{brl(tp)}</b></td>"
             f"<td><b>{brl(tr) if tr else '\u2014'}</b></td><td style='color:{tcor}'><b>{brl(var) if tr else '\u2014'}</b></td>"
             f"<td style='color:{tcor}'><b>{pct_txt(pct) if tr else '\u2014'}</b></td><td></td></tr>")
    st.markdown(f"""<div class='scroll'><table class="lle"><tr>
        <th style='text-align:left'>Mês</th><th>Planejado</th><th>Realizado</th>
        <th>Var. (R$)</th><th>Var. (%)</th><th>Status</th></tr>{linhas}{total}</table></div>""", unsafe_allow_html=True)

    # ---------- registrar / editar ----------
    st.divider()
    st.markdown("#### Registrar / editar CMV")
    st.caption("Preencha Planejado e Realizado por empresa e mês. Toda alteração fica registrada no log.")
    emps = [1, 2] if not emp else [emp]
    idx = {(int(r["mes"]), int(r["uni_cod"])): r for _, r in dfr.iterrows()}
    keys, disp_mes, disp_emp, oplan, oreal = [], [], [], [], []
    for u in emps:
        for m in range(1, 13):
            r = idx.get((m, u))
            keys.append((m, u))
            disp_mes.append(MESES[m]); disp_emp.append(EMP[u])
            oplan.append(round(float(r["valor_planejado"]) if r is not None else 0.0, 2))
            oreal.append(round(float(r["valor_realizado"]) if r is not None else 0.0, 2))
    dedit = pd.DataFrame({"Mês": disp_mes, "Empresa": disp_emp, "Planejado": oplan, "Realizado": oreal})
    ed = st.data_editor(dedit, key=f"cmv_grid_{emp}", hide_index=True, use_container_width=True,
                        num_rows="fixed", disabled=["Mês", "Empresa"],
                        column_config={"Planejado": st.column_config.NumberColumn(format="%.2f", step=0.01),
                                       "Realizado": st.column_config.NumberColumn(format="%.2f", step=0.01)})
    if st.button("Salvar CMV", key="cmv_save", type="primary"):
        np_, nr_ = list(ed["Planejado"]), list(ed["Realizado"])
        mudou = 0
        for i, (m, u) in enumerate(keys):
            try: p_novo = round(float(np_[i]), 2); r_novo = round(float(nr_[i]), 2)
            except (TypeError, ValueError): continue
            mud_p = abs(p_novo - oplan[i]) > 0.005
            mud_r = abs(r_novo - oreal[i]) > 0.005
            if not (mud_p or mud_r): continue
            c.table("cmv_valor").upsert(dict(ano=ano, mes=m, uni_cod=u, unidade=EMP[u],
                valor_planejado=p_novo, valor_realizado=r_novo, atualizado_por=prof.get("nome", "")),
                on_conflict="ano,mes,uni_cod").execute()
            if mud_p:
                c.table("cmv_log").insert(dict(ano=ano, mes=m, uni_cod=u, campo="valor_planejado",
                    valor_antigo=oplan[i], valor_novo=p_novo, alterado_por=prof.get("nome", ""))).execute(); mudou += 1
            if mud_r:
                c.table("cmv_log").insert(dict(ano=ano, mes=m, uni_cod=u, campo="valor_realizado",
                    valor_antigo=oreal[i], valor_novo=r_novo, alterado_por=prof.get("nome", ""))).execute(); mudou += 1
        if mudou:
            limpar_cache()
            st.success(f"{mudou} alteração(ões) de CMV salva(s) e registrada(s) no log.")
            st.rerun()
        else:
            st.info("Nenhuma alteração detectada.")

    # ---------- histórico ----------
    log = carregar_cmv_log(300)
    if log:
        st.markdown("###### Histórico de alterações de CMV")
        CL = {"valor_planejado": "Planejado", "valor_realizado": "Realizado"}
        ln = ""
        for glog in log[:300]:
            quando = str(glog.get("alterado_em", "") or "")[:16].replace("T", " ")
            va = float(glog.get("valor_antigo") or 0); vn = float(glog.get("valor_novo") or 0)
            seta = VERMELHO if vn > va else VERDE  # custo maior = vermelho
            ln += (f"<tr><td style='text-align:left'>{quando}</td><td style='text-align:left'>{glog.get('alterado_por','') or '\u2014'}</td>"
                   f"<td style='text-align:left'>{MESES[int(glog.get('mes') or 1)]}</td>"
                   f"<td style='text-align:center'>{EMP.get(int(glog.get('uni_cod') or 0), glog.get('uni_cod'))}</td>"
                   f"<td style='text-align:center'>{CL.get(glog.get('campo'), glog.get('campo'))}</td>"
                   f"<td>{brl(va)}</td><td style='color:{seta}'>{brl(vn)}</td></tr>")
        st.markdown(f"""<table class="lle"><tr>
            <th style='text-align:left'>Quando</th><th style='text-align:left'>Quem</th><th style='text-align:left'>Mês</th>
            <th style='text-align:center'>Empresa</th><th style='text-align:center'>Campo</th><th>De</th><th>Para</th></tr>{ln}</table>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- administração
def tela_admin(c, prof, ano):
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
def tela_acompanhamento(c, prof, banda, df_orc, cg, is_ctrl, ano, mes):
    st.markdown("<div class='modtag'>Módulo Acompanhamento de Despesas — Orçado x Realizado</div>", unsafe_allow_html=True)

    if df_orc.empty:
        st.info("Nenhum dado carregado ainda." + (" Use a aba 'Importar dados'." if is_ctrl else " Fale com a controladoria."))
        return
    df = df_orc.copy()
    if is_ctrl and cg:
        df["_resp"] = df.apply(lambda r: cg.get((int(r["uni_cod"]), int(r["cr_cod"])), ("—", ""))[0], axis=1)

    # ---------- filtros horizontais (o período vem do seletor global Ano/Mês) ----------
    fcols = st.columns([1.5, 1.2, 1.7, 1.7] if is_ctrl else [1.2, 1.7, 1.7])
    i = 0
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

    # SEM consolidação: cada empresa (PISA/KING) vira uma linha própria por CR+conta
    d_mes = df[df["mes"] == mes]
    d_ytd = df[df["mes"] <= mes]

    # aviso do gestor
    if not is_ctrl and mes < get_cobranca(c):
        st.info(f"{MESES[mes]}/{ano} não está sujeito à cobrança de justificativa.")
    elif not is_ctrl:
        js = carregar_justificativas(ano, mes)
        enviadas = {(int(j["uni_cod"]), int(j["cr_cod"]), int(j["conta_cod"])) for j in js if j.get("status") in ("JUSTIFICADO", "EM_REVISAO", "APROVADO")}
        pend = 0
        for _, v in d_mes.iterrows():
            raw, pct = var_de(v["valor_planejado"], v["valor_realizado"])
            lab, _ = classifica(raw, pct, v["conta_cod"], banda)
            if lab == "Desfavorável" and (int(v["uni_cod"]), int(v["cr_cod"]), int(v["conta_cod"])) not in enviadas:
                pend += 1
        if pend:
            st.warning(f"Você tem {pend} conta(s) a justificar em {MESES[mes]}/{ano} — veja a seção Justificativas ao final.")
        else:
            st.success(f"Nenhuma justificativa pendente em {MESES[mes]}/{ano}.")

    # ---------- submenu de seções (uma por vez -> menos rolagem, menos carga) ----------
    secao = st.radio("Seção", ["📌 Resumo", "📈 Evolução mensal", "🔎 Desvios por CR", "📝 Justificativas"],
                     horizontal=True, key="acomp_secao", label_visibility="collapsed")
    st.divider()

    if secao.endswith("Resumo"):
        resumo_colunas(d_mes, d_ytd, banda, mes, ano)
        contadores(d_mes, banda)
        st.caption("Convenção: contas de receita/dedução (código 3 ou 6) têm sinal invertido — a variação mede o IMPACTO no resultado. "
                   "Verde = favorável, vermelho = desfavorável, cinza = dentro da faixa neutra (±"
                   + f"{banda:.1f}".replace(".", ",") + "%).")
    elif "Evolução" in secao:
        st.markdown("#### Evolução mensal — Jan a Dez (mês e acumulado YTD)")
        tabela_evolucao(df, banda, mes)
    elif "Desvios" in secao:
        st.markdown(f"#### Desvios por centro de resultado — {MESES[mes]}/{ano}")
        drill_desvios(d_mes, banda, mes)
    else:
        st.markdown(f"#### Justificativas · {MESES[mes]}/{ano}")
        secao_justificativas(c, prof, d_mes, mes, is_ctrl, banda, ano)

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

    # ----- seletor global de período (Ano + Mês) — vale para todas as telas -----
    ANOS = list(range(2024, 2032))
    pc = st.columns([1.1, 1.4, 6])
    ano = int(pc[0].selectbox("Ano", ANOS, index=(ANOS.index(2026) if 2026 in ANOS else 0), key="g_ano"))
    mes = int(pc[1].selectbox("Mês", list(range(1, 13)), index=5, format_func=lambda m: MESES[m], key="g_mes"))

    # dados compartilhados (cacheados por token — releituras idênticas ficam instantâneas)
    orc = carregar_orc(ano)
    df_orc = pd.DataFrame(orc) if orc else pd.DataFrame()
    cg = {}
    try:
        for r in carregar_cr_gestor():
            cg[(int(r["uni_cod"]), int(r["cr_cod"]))] = ((r.get("gestor") or {}).get("nome", "—"), r.get("cr_nome", ""))
    except Exception:
        cg = {}

    if is_ctrl:
        NAV = ["📊 Acompanhamento", "📥 Justificativas recebidas", "💰 Receita de Vendas", "🧾 CMV",
               "👥 Gestão de Gastos com Pessoal", "✏️ Manutenção Orçamento",
               "🔑 Administração de acessos", "⬆️ Importar dados"]
        aba = st.radio("Navegação", NAV, horizontal=True, key="nav_ctrl", label_visibility="collapsed")
        st.markdown("<hr style='margin:.2rem 0 1rem; border:none; border-top:1px solid #E4E8F0'>", unsafe_allow_html=True)
        if aba == NAV[0]:   tela_acompanhamento(c, prof, banda, df_orc, cg, is_ctrl, ano, mes)
        elif aba == NAV[1]: tela_painel(c, prof, banda, df_orc, cg, ano, mes)
        elif aba == NAV[2]: tela_receita(c, prof, ano)
        elif aba == NAV[3]: tela_cmv(c, prof, ano)
        elif aba == NAV[4]: tela_headcount(c, prof, ano, mes)
        elif aba == NAV[5]: tela_editar_orcado(c, prof, df_orc, ano, mes)
        elif aba == NAV[6]: tela_admin(c, prof, ano)
        elif aba == NAV[7]: tela_importar(c, ano)
    else:
        tela_acompanhamento(c, prof, banda, df_orc, cg, is_ctrl, ano, mes)

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
