"""
LLE Orçamento — Acompanhamento orçado x realizado (Despesas)
Streamlit + Supabase. Segredos: SUPABASE_URL e SUPABASE_ANON_KEY.
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

st.set_page_config(page_title="LLE Orçamento", page_icon="📊", layout="wide")
URL = st.secrets.get("SUPABASE_URL", ""); ANON = st.secrets.get("SUPABASE_ANON_KEY", "")
MESES = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
MABREV = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
STATUS_LABEL = {"PENDENTE": "Pendente", "JUSTIFICADO": "Justificado", "EM_REVISAO": "Em revisão", "DEVOLVIDO": "Devolvido", "APROVADO": "Aprovado"}

LOGO = f"""<svg width="46" height="60" viewBox="0 0 32 44" xmlns="http://www.w3.org/2000/svg">
<polygon points="16,2 23,9 16,16 9,9" fill="{AMARELO}"/><polygon points="16,10 23,17 16,24 9,17" fill="{AMARELO}"/>
<polygon points="16,18 23,25 16,32 9,25" fill="{VERDE}"/><polygon points="16,26 23,33 16,40 9,33" fill="{AZUL_CORP}"/></svg>"""

def inject_css():
    st.markdown(f"""<style>
      .stApp {{ background:{CINZA_BG}; }}
      #MainMenu, footer {{ visibility:hidden; }}
      div.stButton>button[kind="primary"] {{ background:{AZUL_CORP}; border-color:{AZUL_CORP}; }}
      div.stButton>button[kind="primary"]:hover {{ background:{AZUL_PROFUNDO}; border-color:{AZUL_PROFUNDO}; }}
      .lle-header {{ background:linear-gradient(135deg,{AZUL_PROFUNDO} 0%,{AZUL_CORP} 100%);
        border-bottom:3px solid {AMARELO}; border-radius:10px; padding:14px 22px; display:flex;
        align-items:center; gap:16px; margin-bottom:8px; }}
      .lle-header h1 {{ color:#fff; font-size:20px; margin:0; }}
      .lle-header p {{ color:rgba(255,255,255,.7); font-size:12px; margin:2px 0 0; }}
      .cards {{ display:flex; gap:12px; flex-wrap:wrap; margin:6px 0 4px; }}
      .card {{ flex:1; min-width:150px; background:#fff; border:1px solid {LINHA}; border-radius:10px; padding:12px 14px; }}
      .card .lab {{ font-size:11px; color:{CINZA_TXT}; text-transform:uppercase; letter-spacing:.03em; }}
      .card .val {{ font-size:20px; color:{AZUL_PROFUNDO}; font-weight:600; margin-top:3px; }}
      table.lle {{ border-collapse:collapse; width:100%; font-size:13px; background:#fff; }}
      table.lle th {{ background:{AZUL_CORP}; color:#fff; padding:8px 10px; text-align:right; font-weight:600; }}
      table.lle th:first-child {{ text-align:left; }}
      table.lle td {{ padding:6px 10px; text-align:right; border-bottom:1px solid {LINHA}; }}
      table.lle td:first-child {{ text-align:left; }}
      table.lle tr.total td {{ font-weight:700; border-top:2px solid {AZUL_CORP}; background:#EEF2F8; }}
      .scroll {{ overflow-x:auto; }}
    </style>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- helpers
def norm(s): return unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode().lower().strip()
def brl(n): return "R$ " + f"{(n or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
def pct_txt(n): return f"{n:+.1f}".replace(".", ",") + "%"
def chunks(seq, n=500):
    for i in range(0, len(seq), n): yield seq[i:i + n]

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

# ---------------------------------------------------------------- login
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
            <h1 style="color:{AZUL_PROFUNDO}; margin:12px 0 2px;">LLE Ferragens</h1>
            <p style="color:{AZUL_CORP}; font-size:13px; margin:0;">Sistema de Orçamento · Controladoria</p></div>""", unsafe_allow_html=True)
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
    st.markdown(f"""<div class="lle-header">{LOGO}<div>
        <h1>Acompanhamento orçado x realizado · Despesas</h1>
        <p>{prof['nome']} · {papel}</p></div></div>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- blocos de insight
def cards(items):
    html = "<div class='cards'>"
    for lab, val, cor in items:
        c = f"color:{cor};" if cor else ""
        html += f"<div class='card'><div class='lab'>{lab}</div><div class='val' style='{c}'>{val}</div></div>"
    st.markdown(html + "</div>", unsafe_allow_html=True)

def kpis(df_mes, df_ytd, banda):
    vp, vr = df_mes["valor_planejado"].sum(), df_mes["valor_realizado"].sum()
    raw, pct = var_de(vp, vr); lab, cor = classifica(raw, pct, "5", banda)
    st.markdown("###### Resumo — mês selecionado")
    cards([("Orçado", brl(vp), None), ("Realizado", brl(vr), None),
           ("Variação (R$)", brl(raw), cor), ("Variação (%)", pct_txt(pct), cor), ("Status", lab, cor)])
    yp, yr = df_ytd["valor_planejado"].sum(), df_ytd["valor_realizado"].sum()
    yraw, ypct = var_de(yp, yr); ylab, ycor = classifica(yraw, ypct, "5", banda)
    st.markdown("###### Acumulado no ano (YTD até o mês)")
    cards([("Orçado YTD", brl(yp), None), ("Realizado YTD", brl(yr), None),
           ("Var. YTD (R$)", brl(yraw), ycor), ("Var. YTD (%)", pct_txt(ypct), ycor), ("Status YTD", ylab, ycor)])

def contadores(df_mes, banda):
    fav = desf = neu = 0
    for _, r in df_mes.iterrows():
        raw, pct = var_de(r["valor_planejado"], r["valor_realizado"])
        lab, _ = classifica(raw, pct, r["conta_cod"], banda)
        fav += lab == "Favorável"; desf += lab == "Desfavorável"; neu += lab == "Neutro"
    cards([("# Lançamentos", len(df_mes), None), ("# Favoráveis", fav, VERDE),
           ("# Desfavoráveis", desf, VERMELHO), ("# Neutros", neu, CINZA_TXT)])

def tabela_evolucao(df, banda):
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
            st_html = f"<td style='color:{cor}'>{brl(raw)}</td><td style='color:{cor}'>{pct_txt(pct)}</td><td style='color:{cor}'>{lab}</td>"
        yvp, yvr = cum.loc[m, "valor_planejado"], cum.loc[m, "valor_realizado"]
        yraw, ypct = var_de(yvp, yvr); ylab, ycor = classifica(yraw, ypct, "5", banda)
        yhtml = (f"<td>{brl(yvr)}</td><td style='color:{ycor}'>{brl(yraw)}</td>"
                 f"<td style='color:{ycor}'>{pct_txt(ypct)}</td>") if yvr else "<td>—</td><td>—</td><td>—</td>"
        linhas += f"<tr><td>{MESES[m]}</td><td>{brl(vp)}</td><td>{brl(vr) if vr else '—'}</td>{st_html}{yhtml}</tr>"
    st.markdown(f"""<div class='scroll'><table class="lle"><tr>
        <th>Mês</th><th>Orçado</th><th>Realizado</th><th>Var. (R$)</th><th>Var. (%)</th><th>Status</th>
        <th>Realizado YTD</th><th>Var. YTD (R$)</th><th>Var. YTD (%)</th></tr>{linhas}</table></div>""", unsafe_allow_html=True)

def tabela_cr(df_mes, banda):
    tot = df_mes["valor_realizado"].sum() or 1
    tot_o = df_mes["valor_planejado"].sum() or 1
    g = df_mes.groupby("cr_nome")[["valor_planejado", "valor_realizado"]].sum().reset_index()
    rows = []
    for _, r in g.iterrows():
        raw, pct = var_de(r["valor_planejado"], r["valor_realizado"])
        lab, cor = classifica(raw, pct, "5", banda)
        rows.append((r["cr_nome"], r["valor_planejado"], r["valor_realizado"], raw, pct, lab, cor))
    rows.sort(key=lambda x: x[3], reverse=True)
    linhas = ""
    for nome, vp, vr, raw, pct, lab, cor in rows:
        linhas += (f"<tr><td>{nome}</td><td>{brl(vp)}</td><td>{vp/tot_o*100:.1f}%</td><td>{brl(vr)}</td>"
                   f"<td>{vr/tot*100:.1f}%</td><td style='color:{cor}'>{brl(raw)}</td>"
                   f"<td style='color:{cor}'>{pct_txt(pct)}</td><td style='color:{cor}'>{lab}</td></tr>")
    tvp, tvr = df_mes["valor_planejado"].sum(), df_mes["valor_realizado"].sum()
    traw, tpct = var_de(tvp, tvr); _, tcor = classifica(traw, tpct, "5", banda)
    linhas += (f"<tr class='total'><td>Total</td><td>{brl(tvp)}</td><td>100%</td><td>{brl(tvr)}</td><td>100%</td>"
               f"<td style='color:{tcor}'>{brl(traw)}</td><td style='color:{tcor}'>{pct_txt(tpct)}</td><td></td></tr>")
    st.markdown(f"""<div class='scroll'><table class="lle"><tr>
        <th>Centro de resultado</th><th>Orçado</th><th>AV% Orç</th><th>Realizado</th><th>AV% Real</th>
        <th>Var. (R$)</th><th>Var. (%)</th><th>Status</th></tr>{linhas}</table></div>""", unsafe_allow_html=True)

def top5(df_mes, banda):
    recs = []
    for _, r in df_mes.iterrows():
        raw, pct = var_de(r["valor_planejado"], r["valor_realizado"])
        lab, _ = classifica(raw, pct, r["conta_cod"], banda)
        recs.append((r["conta_desc"], r["cr_nome"], r["valor_planejado"], r["valor_realizado"], raw, lab))
    desf = sorted([x for x in recs if x[5] == "Desfavorável"], key=lambda x: x[4], reverse=True)[:5]
    fav = sorted([x for x in recs if x[5] == "Favorável"], key=lambda x: x[4])[:5]
    def bloco(titulo, dados, cor):
        linhas = "".join(f"<tr><td>{d[0]}</td><td>{brl(d[3])}</td><td style='color:{cor}'>{brl(d[4])}</td></tr>" for d in dados) or "<tr><td colspan='3' style='color:#999'>—</td></tr>"
        return f"""<table class="lle"><tr><th style='background:{cor}'>{titulo}</th><th style='background:{cor}'>Realizado</th><th style='background:{cor}'>Var. (R$)</th></tr>{linhas}</table>"""
    a, b = st.columns(2)
    a.markdown(bloco("Top 5 desfavoráveis (conta)", desf, VERMELHO), unsafe_allow_html=True)
    b.markdown(bloco("Top 5 favoráveis (conta)", fav, VERDE), unsafe_allow_html=True)

# ---------------------------------------------------------------- justificativas
def secao_justificativas(c, prof, df_mes, mes, is_ctrl, banda):
    js = c.table("justificativa").select("*").eq("ano", 2026).eq("mes", mes).execute().data or []
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
    for v, raw, pct, j in itens:
        status = j.get("status", "PENDENTE")
        with st.expander(f"{v['conta_cod']} · {v.get('conta_desc','')} — {v.get('cr_nome','')} ({v.get('unidade','')}) · {brl(raw)} · [{STATUS_LABEL.get(status)}]"):
            a, b, d = st.columns(3)
            a.metric("Orçado", brl(v["valor_planejado"])); b.metric("Realizado", brl(v["valor_realizado"])); d.metric("Variação", brl(raw), pct_txt(pct))
            st.markdown("**Histórico das notas que compõem o realizado**")
            det = (c.table("operacional_detalhe").select("num_doc, valor, historico").eq("ano", 2026).eq("mes", mes)
                   .eq("uni_cod", v["uni_cod"]).eq("cr_cod", v["cr_cod"]).eq("conta_cod", v["conta_cod"]).execute().data or [])
            if det:
                dfd = pd.DataFrame(det); dfd["valor"] = dfd["valor"].map(brl); dfd.columns = ["NF/Doc", "Valor", "Histórico"]
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
                    c.table("justificativa").upsert({**key, "texto": txt, "status": "PENDENTE", "atualizado_por": prof["nome"]}, on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute(); st.rerun()
                if c2.button("Enviar justificativa", key=f"en_{kb}", type="primary"):
                    if not txt.strip(): st.error("Escreva a justificativa antes de enviar.")
                    else:
                        c.table("justificativa").upsert({**key, "texto": txt, "status": "JUSTIFICADO", "atualizado_por": prof["nome"]}, on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute(); st.rerun()
            elif not is_ctrl:
                st.info(f"Justificativa: {j.get('texto') or '—'}"); st.caption("Aguardando a controladoria — não editável.")
            if is_ctrl:
                st.info(f"Justificativa do gestor: {j.get('texto') or '— (ainda não enviada)'}")
                if status in ("JUSTIFICADO", "EM_REVISAO"):
                    coment = st.text_input("Comentário (para devolução)", key=f"cm_{kb}")
                    c1, c2 = st.columns(2)
                    if c1.button("Aprovar", key=f"ap_{kb}", type="primary"):
                        c.table("justificativa").update({"status": "APROVADO"}).match(key).execute(); st.rerun()
                    if c2.button("Devolver", key=f"dv_{kb}"):
                        c.table("justificativa").update({"status": "DEVOLVIDO", "comentario_controladoria": coment}).match(key).execute(); st.rerun()
    if not itens:
        st.success("Nenhum desvio desfavorável a justificar com os filtros atuais.")

# ---------------------------------------------------------------- importar
def tela_importar(c):
    st.subheader("Configuração")
    atual = get_faixa(c)
    nova = st.number_input("Faixa neutra (±%) — regra de classificação para todos os gestores",
                           value=float(atual), step=0.5, min_value=0.0)
    if st.button("Salvar faixa neutra"):
        set_faixa(c, nova); st.success(f"Faixa neutra atualizada para ±{nova:.1f}%.".replace(".", ","))
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

def tela_painel(c, prof):
    header(prof)
    st.subheader("Justificativas recebidas")
    mes = st.selectbox("Mês", list(range(1, 13)), index=5, format_func=lambda m: f"{MESES[m]}/2026")

    js = c.table("justificativa").select("*").eq("ano", 2026).eq("mes", mes).execute().data or []

    # mapa CR -> (gestor, nome do CR)  [chave sempre inteira, para casar entre tabelas]
    cg = {}
    for r in (c.table("cr_gestor").select("uni_cod, cr_cod, cr_nome, gestor(nome)").execute().data or []):
        cg[(int(r["uni_cod"]), int(r["cr_cod"]))] = ((r.get("gestor") or {}).get("nome", "—"), r.get("cr_nome", ""))

    # ----- Gerência de pendências (desvios desfavoráveis ainda não enviados) -----
    dfo = pd.DataFrame(ler_tudo(c, "orc_realizado", 2026))
    dfo = dfo[dfo["mes"] == mes]
    banda = get_faixa(c)
    enviadas = {(int(j["uni_cod"]), int(j["cr_cod"]), int(j["conta_cod"])) for j in js if j.get("status") in ("JUSTIFICADO", "EM_REVISAO", "APROVADO")}
    devolvidas = {(int(j["uni_cod"]), int(j["cr_cod"]), int(j["conta_cod"])) for j in js if j.get("status") == "DEVOLVIDO"}
    # agrega pendências por gestor -> lista de contas
    pend = {}   # gestor -> {"n","valor","devolv","itens":[...]}
    for _, v in dfo.iterrows():
        raw, pct = var_de(v["valor_planejado"], v["valor_realizado"])
        lab, _ = classifica(raw, pct, v["conta_cod"], banda)
        if lab != "Desfavorável":
            continue
        chave = (int(v["uni_cod"]), int(v["cr_cod"]), int(v["conta_cod"]))
        if chave in enviadas:
            continue
        gestor, _crn = cg.get((int(v["uni_cod"]), int(v["cr_cod"])), ("Sem gestor", ""))
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

    # resumo por gestor
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

    # detalhe por gestor -> centro de resultado
    st.markdown("###### Detalhe por gestor e centro de resultado")
    for gestor in sorted(df["gestor"].unique()):
        sub = df[df["gestor"] == gestor]
        with st.expander(f"{gestor} — {len(sub)} justificativa(s)"):
            for cr in sorted(sub["cr"].unique()):
                st.markdown(f"**{cr}**")
                linhas2 = ""
                for _, r in sub[sub["cr"] == cr].iterrows():
                    st_lab = STATUS_LABEL.get(r["status"])
                    cor = {"APROVADO": VERDE, "DEVOLVIDO": VERMELHO, "PENDENTE": CINZA_TXT}.get(r["status"], AZUL_CORP)
                    txt = (r["texto"][:120] + "…") if len(r["texto"]) > 120 else (r["texto"] or "—")
                    linhas2 += f"<tr><td>{r['conta']}</td><td style='color:{cor}'>{st_lab}</td><td style='text-align:left'>{txt}</td></tr>"
                st.markdown(f"""<table class="lle"><tr><th>Conta</th><th>Situação</th>
                    <th style='text-align:left'>Justificativa</th></tr>{linhas2}</table>""", unsafe_allow_html=True)

def tela_admin(c, prof):
    header(prof)
    st.subheader("Administração de acessos")
    st.caption("Substitua o e-mail de um gestor desligado, ative/desative acessos e transfira centros de resultado. "
               "O histórico de justificativas é sempre preservado.")

    gestores = c.table("gestor").select("codigo, nome, papel").order("nome").execute().data or []
    usuarios = c.table("gestor_usuario").select("email, gestor_codigo, papel_acesso, ativo").execute().data or []
    nome_por_cod = {g["codigo"]: g["nome"] for g in gestores}
    op_gestor = {f"{g['nome']} ({g['codigo']})": g["codigo"] for g in gestores}

    aba = st.radio("Ação", ["Substituir e-mail (desligamento)", "Ativar / desativar acesso", "Transferir centro de resultado"], horizontal=False)

    # 1) SUBSTITUIR E-MAIL --------------------------------------------------
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
                c.table("gestor_usuario").upsert({"email": email_novo, "gestor_codigo": cod, "papel_acesso": "titular", "ativo": True, "senha_provisoria": True}, on_conflict="email").execute()
                st.success(f"Feito: {email_novo} agora responde por {nome_por_cod.get(cod, cod)}." + (f" {email_antigo} foi desativado." if email_antigo != '(nenhum)' else ""))
                st.rerun()

    # 2) ATIVAR / DESATIVAR -------------------------------------------------
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
                    c.table("gestor_usuario").update({"ativo": novo_estado}).eq("email", u["email"]).execute()
                    st.rerun()

    # 3) TRANSFERIR CR ------------------------------------------------------
    else:
        st.markdown("###### Transferir um centro de resultado para outro gestor")
        st.caption("Muda o dono do CR. As justificativas já feitas continuam registradas na conta/CR, sem alteração.")
        crs = c.table("cr_gestor").select("uni_cod, cr_cod, cr_nome, gestor_codigo").order("cr_nome").execute().data or []
        op_cr = {f"{r['cr_nome']} (uni {r['uni_cod']} · CR {r['cr_cod']}) — hoje: {nome_por_cod.get(r['gestor_codigo'], r['gestor_codigo'])}": (r["uni_cod"], r["cr_cod"]) for r in crs}
        cr_sel = st.selectbox("Centro de resultado", list(op_cr.keys()))
        destino = st.selectbox("Novo gestor responsável", list(op_gestor.keys()))
        if st.button("Transferir CR", type="primary"):
            uni, crc = op_cr[cr_sel]
            c.table("cr_gestor").update({"gestor_codigo": op_gestor[destino]}).eq("uni_cod", uni).eq("cr_cod", crc).execute()
            st.success(f"CR transferido para {destino.split(' (')[0]}.")
            st.rerun()

# ---------------------------------------------------------------- acompanhamento
def tela_acompanhamento(c, prof):
    is_ctrl = prof["papel"] == "controladoria"
    header(prof)
    banda = get_faixa(c)
    orc = ler_tudo(c, "orc_realizado", 2026)
    if not orc:
        st.info("Nenhum dado carregado ainda." + (" Use 'Importar dados'." if is_ctrl else " Fale com a controladoria.")); return
    df = pd.DataFrame(orc)
    resp_map = {}
    if is_ctrl:
        try:
            for r in (c.table("cr_gestor").select("uni_cod, cr_cod, gestor(nome)").execute().data or []):
                resp_map[(r["uni_cod"], r["cr_cod"])] = (r.get("gestor") or {}).get("nome", "—")
        except Exception:
            resp_map = {}

    with st.sidebar:
        st.markdown(f"<div style='text-align:center'>{LOGO}</div>", unsafe_allow_html=True)
        st.markdown(f"**{prof['nome']}**"); st.caption("Controladoria" if is_ctrl else "Gestor")
        st.caption(f"Faixa neutra vigente: ±{banda:.1f}%".replace(".", ","))
        st.divider()
        mes = st.selectbox("Mês", list(range(1, 13)), index=5, format_func=lambda m: f"{MESES[m]}/2026")
        st.markdown("**Filtros**")
        if is_ctrl and resp_map:
            df = df.copy(); df["_resp"] = df.apply(lambda r: resp_map.get((r["uni_cod"], r["cr_cod"]), "—"), axis=1)
            resps = ["Todos"] + sorted(df["_resp"].dropna().unique().tolist())
            f_resp = st.selectbox("Gestor", resps)
            if f_resp != "Todos": df = df[df["_resp"] == f_resp]
        unis = ["Todas"] + sorted(df["unidade"].dropna().unique().tolist())
        f_uni = st.selectbox("Unidade", unis)
        if f_uni != "Todas": df = df[df["unidade"] == f_uni]
        crs = ["Todos"] + sorted(df["cr_nome"].dropna().unique().tolist())
        f_cr = st.selectbox("Centro de resultado", crs)
        if f_cr != "Todos": df = df[df["cr_nome"] == f_cr]
        contas = ["Todas"] + sorted(df["conta_desc"].dropna().unique().tolist())
        f_conta = st.selectbox("Conta", contas)
        if f_conta != "Todas": df = df[df["conta_desc"] == f_conta]
        st.divider()
        if st.button("Sair"):
            for k in ("access_token", "refresh_token", "email"): st.session_state.pop(k, None)
            st.rerun()

    d_mes = df[df["mes"] == mes]
    d_ytd = df[df["mes"] <= mes]

    # pendências do mês (para o aviso do gestor): desvio desfavorável ainda não enviado
    if not is_ctrl:
        js = c.table("justificativa").select("uni_cod,cr_cod,conta_cod,status").eq("ano", 2026).eq("mes", mes).execute().data or []
        enviadas = {(j["uni_cod"], j["cr_cod"], j["conta_cod"]) for j in js if j.get("status") in ("JUSTIFICADO", "EM_REVISAO", "APROVADO")}
        pend = 0
        for _, v in d_mes.iterrows():
            raw, pct = var_de(v["valor_planejado"], v["valor_realizado"])
            lab, _ = classifica(raw, pct, v["conta_cod"], banda)
            if lab == "Desfavorável" and (v["uni_cod"], v["cr_cod"], v["conta_cod"]) not in enviadas:
                pend += 1
        if pend:
            st.warning(f"Você tem {pend} conta(s) a justificar em {MESES[mes]}/2026. Elas estão na seção Justificativas, ao final da página.")
        else:
            st.success(f"Nenhuma justificativa pendente em {MESES[mes]}/2026.")
    st.caption("Convenção: contas de receita/dedução (código 3 ou 6) têm sinal invertido — a variação mede o IMPACTO no resultado. "
               "Verde = favorável (gastou/deduziu menos que o previsto), vermelho = desfavorável, cinza = dentro da faixa neutra.")

    kpis(d_mes, d_ytd, banda)
    st.markdown("###### Contadores do mês")
    contadores(d_mes, banda)
    st.divider()
    st.markdown("#### Evolução mensal — Jan a Dez")
    tabela_evolucao(df, banda)
    st.divider()
    st.markdown(f"#### Análise por centro de resultado — {MESES[mes]}")
    tabela_cr(d_mes, banda)
    st.markdown("#### Maiores desvios do mês (por conta)")
    top5(d_mes, banda)
    st.divider()
    st.markdown(f"#### Justificativas · {MESES[mes]}/2026")
    secao_justificativas(c, prof, d_mes, mes, is_ctrl, banda)

# ---------------------------------------------------------------- main
def main():
    inject_css()
    if not URL or not ANON:
        st.error("Faltam os segredos SUPABASE_URL e SUPABASE_ANON_KEY."); return
    if "access_token" not in st.session_state:
        tela_login(); return
    c = client(); prof = perfil(c, st.session_state.get("email", ""))
    if not prof:
        st.error("Seu e-mail não está cadastrado. Fale com a controladoria.")
        if st.button("Sair"): st.session_state.clear(); st.rerun()
        return
    if prof.get("senha_provisoria"):
        tela_trocar_senha(c, st.session_state.get("email", "")); return
    if prof["papel"] == "controladoria":
        aba = st.sidebar.radio("Menu", ["Acompanhamento", "Justificativas recebidas", "Administração de acessos", "Importar dados"])
        if aba == "Importar dados": header(prof); tela_importar(c); return
        if aba == "Justificativas recebidas": tela_painel(c, prof); return
        if aba == "Administração de acessos": tela_admin(c, prof); return
    tela_acompanhamento(c, prof)

main()
