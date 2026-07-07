"""
LLE Orçamento — Acompanhamento orçado x realizado (módulo de Justificativas)
Arquivo único. Requer os segredos SUPABASE_URL e SUPABASE_ANON_KEY em .streamlit/secrets.toml
(ou na aba Secrets do Streamlit Cloud).
"""
import re
import unicodedata
import pandas as pd
import streamlit as st
from supabase import create_client

# ---------------------------------------------------------------- config
st.set_page_config(page_title="LLE Orçamento", page_icon="📊", layout="wide")
URL = st.secrets.get("SUPABASE_URL", "")
ANON = st.secrets.get("SUPABASE_ANON_KEY", "")
MESES = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
STATUS_LABEL = {"PENDENTE": "Pendente", "JUSTIFICADO": "Justificado", "EM_REVISAO": "Em revisão",
                "DEVOLVIDO": "Devolvido", "APROVADO": "Aprovado"}
TH_RS, TH_PCT = 3000, 10          # gatilho: desvio desfavorável acima de R$ 3.000 ou 10%
NAVY = "#13194D"

# ---------------------------------------------------------------- helpers
def brl(n):
    return "R$ " + f"{(n or 0):,.0f}".replace(",", ".")

def norm(s):
    return unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode().lower().strip()

def num(x):
    if x is None or pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace("R$", "").strip()
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0

def to_int(x):
    """Converte para inteiro; devolve None se vazio/invalido (linha e ignorada)."""
    if x is None or pd.isna(x):
        return None
    s = str(x).strip()
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None

def extrai_hist(h):
    if not isinstance(h, str):
        return None, str(h or "")
    m = re.search(r"\b(?:NF|CFE)\s*0*([0-9]{3,})", h, re.I)
    return (m.group(1) if m else None), re.sub(r"\s*-?\d+\.\d+\s*$", "", h).strip()

def desvio(v):
    va = float(v["valor_realizado"]) - float(v["valor_planejado"])
    p = va / float(v["valor_planejado"]) * 100 if v["valor_planejado"] else (100 if v["valor_realizado"] else 0)
    desfav = norm(v.get("classificacao", "")).startswith("desfav") if v.get("classificacao") else va > 0
    flag = desfav and (abs(va) >= TH_RS or abs(p) >= TH_PCT)
    return va, p, desfav, flag

# ---------------------------------------------------------------- supabase
def client():
    c = create_client(URL, ANON)
    tok = st.session_state.get("access_token")
    rtok = st.session_state.get("refresh_token")
    if tok and rtok:
        try:
            c.auth.set_session(tok, rtok)
        except Exception:
            pass
    return c

def do_login(email, senha):
    try:
        c = create_client(URL, ANON)
        res = c.auth.sign_in_with_password({"email": email.strip(), "password": senha})
        st.session_state.access_token = res.session.access_token
        st.session_state.refresh_token = res.session.refresh_token
        st.session_state.email = email.strip()
        return True, ""
    except Exception:
        return False, "E-mail ou senha inválidos."

def perfil(c, email):
    r = c.table("gestor_usuario").select("gestor_codigo, papel_acesso, gestor(nome, papel)").eq("email", email).execute()
    if not r.data:
        return None
    row = r.data[0]
    g = row.get("gestor") or {}
    return {"gestor_codigo": row["gestor_codigo"], "nome": g.get("nome", email),
            "papel": g.get("papel", "gestor"), "acesso": row["papel_acesso"]}

def chunks(seq, n=500):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]

# ---------------------------------------------------------------- telas
def tela_login():
    st.markdown(f"<h2 style='color:{NAVY}'>Grupo LLE · Orçamento & Controle</h2>", unsafe_allow_html=True)
    st.caption("Acompanhamento orçado x realizado")
    with st.form("login"):
        email = st.text_input("E-mail")
        senha = st.text_input("Senha", type="password")
        ok = st.form_submit_button("Entrar", type="primary")
    if ok:
        sucesso, msg = do_login(email, senha)
        if sucesso:
            st.rerun()
        else:
            st.error(msg)

def tela_justificativas(c, prof):
    is_ctrl = prof["papel"] == "controladoria"
    st.sidebar.markdown(f"**{prof['nome']}**")
    st.sidebar.caption("Controladoria" if is_ctrl else f"Gestor · {prof['acesso']}")
    mes = st.sidebar.selectbox("Mês", list(range(1, 13)), index=5, format_func=lambda m: f"{MESES[m]}/2026")
    ver_todas = st.sidebar.checkbox("Mostrar todas as contas", value=False)
    filtro_status = st.sidebar.selectbox("Situação", ["Todas"] + list(STATUS_LABEL.keys()),
                                         format_func=lambda s: STATUS_LABEL.get(s, s))
    if st.sidebar.button("Sair"):
        for k in ("access_token", "refresh_token", "email"):
            st.session_state.pop(k, None)
        st.rerun()

    # dados (RLS decide o que cada um enxerga)
    rows = c.table("orc_realizado").select("*").eq("ano", 2026).eq("mes", mes).execute().data or []
    js = c.table("justificativa").select("*").eq("ano", 2026).eq("mes", mes).execute().data or []
    jmap = {(j["uni_cod"], j["cr_cod"], j["conta_cod"]): j for j in js}

    itens = []
    for v in rows:
        va, p, desfav, flag = desvio(v)
        if not ver_todas and not flag:
            continue
        j = jmap.get((v["uni_cod"], v["cr_cod"], v["conta_cod"]), {"status": "PENDENTE", "texto": "", "comentario_controladoria": ""})
        if filtro_status != "Todas" and j.get("status", "PENDENTE") != filtro_status:
            continue
        itens.append((v, va, p, desfav, j))
    itens.sort(key=lambda t: t[1], reverse=True)

    total_flag = sum(1 for v in rows if desvio(v)[3])
    st.markdown(f"### Justificativas · {MESES[mes]}/2026")
    st.caption(f"{len(itens)} conta(s) em tela · {total_flag} com desvio desfavorável · "
               f"gatilho acima de {brl(TH_RS)} ou {TH_PCT}%")

    for v, va, p, desfav, j in itens:
        status = j.get("status", "PENDENTE")
        titulo = f"{v['conta_cod']} · {v.get('conta_desc','')}  —  {v.get('cr_nome','')} ({v.get('unidade','')})  ·  [{STATUS_LABEL.get(status)}]"
        with st.expander(titulo):
            a, b, d = st.columns(3)
            a.metric("Orçado", brl(v["valor_planejado"]))
            b.metric("Realizado", brl(v["valor_realizado"]))
            d.metric("Variação", brl(va), f"{p:+.1f}%")

            st.markdown("**Histórico das notas que compõem o realizado**")
            det = (c.table("operacional_detalhe").select("num_doc, valor, historico")
                   .eq("ano", 2026).eq("mes", mes).eq("uni_cod", v["uni_cod"])
                   .eq("cr_cod", v["cr_cod"]).eq("conta_cod", v["conta_cod"]).execute().data or [])
            if det:
                df = pd.DataFrame(det)
                df["valor"] = df["valor"].map(brl)
                df.columns = ["NF/Doc", "Valor", "Histórico"]
                st.dataframe(df, use_container_width=True, hide_index=True)
            else:
                st.caption("Sem detalhe operacional para esta conta neste mês.")

            key = dict(ano=2026, mes=mes, uni_cod=v["uni_cod"], cr_cod=v["cr_cod"], conta_cod=v["conta_cod"])

            if status == "DEVOLVIDO" and j.get("comentario_controladoria"):
                st.warning(f"Controladoria: {j['comentario_controladoria']}")

            # ---- gestor
            if not is_ctrl and status in ("PENDENTE", "DEVOLVIDO"):
                _kb = f"{mes}_{v['uni_cod']}_{v['cr_cod']}_{v['conta_cod']}"
                txt = st.text_area("Justificativa", value=j.get("texto", "") or "", key=f"txt_{_kb}")
                c1, c2 = st.columns(2)
                if c1.button("Salvar rascunho", key=f"sv_{_kb}"):
                    c.table("justificativa").upsert({**key, "texto": txt, "status": "PENDENTE",
                                                     "atualizado_por": prof["nome"]},
                                                    on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute()
                    st.rerun()
                if c2.button("Enviar justificativa", key=f"en_{_kb}", type="primary"):
                    if not txt.strip():
                        st.error("Escreva a justificativa antes de enviar.")
                    else:
                        c.table("justificativa").upsert({**key, "texto": txt, "status": "JUSTIFICADO",
                                                         "atualizado_por": prof["nome"]},
                                                        on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute()
                        st.rerun()
            elif not is_ctrl:
                st.info(f"Justificativa: {j.get('texto') or '—'}")
                st.caption("Aguardando a controladoria — não editável.")

            # ---- controladoria
            if is_ctrl:
                _kb = f"{mes}_{v['uni_cod']}_{v['cr_cod']}_{v['conta_cod']}"
                st.info(f"Justificativa do gestor: {j.get('texto') or '— (ainda não enviada)'}")
                if status in ("JUSTIFICADO", "EM_REVISAO"):
                    coment = st.text_input("Comentário (para devolução)", key=f"cm_{_kb}")
                    c1, c2 = st.columns(2)
                    if c1.button("Aprovar", key=f"ap_{_kb}", type="primary"):
                        c.table("justificativa").update({"status": "APROVADO"}).match(key).execute()
                        st.rerun()
                    if c2.button("Devolver", key=f"dv_{_kb}"):
                        c.table("justificativa").update({"status": "DEVOLVIDO",
                                                         "comentario_controladoria": coment}).match(key).execute()
                        st.rerun()

    if not itens:
        st.success("Nada pendente com os filtros atuais.")

def tela_importar(c):
    st.markdown("### Importar dados (controladoria)")
    st.caption("Suba as duas planilhas do mês. A base orçado x realizado atualiza os números; "
               "o arquivo operacional traz o histórico das notas.")
    ano = st.number_input("Ano", value=2026, step=1)

    f1 = st.file_uploader("1) Base orçado x realizado (Pasta)", type=["xlsx"], key="f1")
    if f1 and st.button("Importar orçado x realizado"):
        df = pd.read_excel(f1)
        def pick(row, cands):
            for cand in cands:
                for k in row.index:
                    if norm(k) == cand or cand in norm(k):
                        return row[k]
            return ""
        recs = []
        for _, r in df.iterrows():
            conta = pick(r, ["codigo conta"])
            uni_cod = to_int(pick(r, ["codigo unidade"]))
            cr_cod = to_int(pick(r, ["codigo centro"]))
            conta_cod = to_int(conta)
            if None in (uni_cod, cr_cod, conta_cod):
                continue
            recs.append(dict(
                ano=to_int(pick(r, ["ano"])) or int(ano), mes=to_int(pick(r, ["mes"])) or 1,
                uni_cod=uni_cod, unidade=str(pick(r, ["descricao unidade"]) or ""),
                cr_cod=cr_cod, cr_nome=str(pick(r, ["descricao centro"]) or ""),
                cr_grupo=str(pick(r, ["cr grupo"]) or ""), conta_cod=conta_cod,
                conta_desc=str(pick(r, ["descricao conta"]) or ""),
                valor_planejado=num(pick(r, ["valor planejado"])), valor_realizado=num(pick(r, ["valor realizado"])),
                tipo_conta=str(pick(r, ["tipo conta"]) or ""), classificacao=str(pick(r, ["classificacao"]) or "")))
        for ch in chunks(recs):
            c.table("orc_realizado").upsert(ch, on_conflict="ano,mes,uni_cod,cr_cod,conta_cod").execute()
        st.success(f"{len(recs)} linhas de orçado x realizado importadas.")

    f2 = st.file_uploader("2) Detalhe operacional (com COMPLHIST)", type=["xlsx"], key="f2")
    mes_op = st.number_input("Mês do arquivo operacional", value=6, min_value=1, max_value=12, step=1)
    if f2 and st.button("Importar histórico das notas"):
        raw = pd.read_excel(f2, header=None, nrows=10)
        hi = 0
        for i in range(min(10, len(raw))):
            if any("complhist" in norm(x) for x in raw.iloc[i].tolist()):
                hi = i
                break
        df = pd.read_excel(f2, header=hi).dropna(how="all")
        def pick(row, cands):
            for cand in cands:
                for k in row.index:
                    if norm(k) == cand or cand in norm(k):
                        return row[k]
            return ""
        recs = []
        for _, r in df.iterrows():
            conta = pick(r, ["codctactb", "codigo conta"])
            uni_cod = to_int(pick(r, ["codigo unidade", "codigo_unidade"]))
            cr_cod = to_int(pick(r, ["codcencus", "codigo centro"]))
            conta_cod = to_int(conta)
            if None in (uni_cod, cr_cod, conta_cod):
                continue
            nf, hist = extrai_hist(pick(r, ["complhist"]))
            recs.append(dict(ano=to_int(pick(r, ["ano"])) or int(ano),
                             mes=to_int(pick(r, ["mes"])) or int(mes_op),
                             uni_cod=uni_cod, cr_cod=cr_cod, conta_cod=conta_cod,
                             num_doc=nf, valor=num(pick(r, ["valor"])), historico=hist))
        c.table("operacional_detalhe").delete().eq("ano", int(ano)).eq("mes", int(mes_op)).execute()
        for ch in chunks(recs):
            c.table("operacional_detalhe").insert(ch).execute()
        st.success(f"{len(recs)} lançamentos com histórico importados para {MESES[int(mes_op)]}/{int(ano)}.")

# ---------------------------------------------------------------- main
def main():
    if not URL or not ANON:
        st.error("Faltam os segredos SUPABASE_URL e SUPABASE_ANON_KEY. Configure em Secrets.")
        return
    if "access_token" not in st.session_state:
        tela_login()
        return
    c = client()
    prof = perfil(c, st.session_state.get("email", ""))
    if not prof:
        st.error("Seu e-mail não está cadastrado como gestor. Fale com a controladoria.")
        if st.button("Sair"):
            st.session_state.clear()
            st.rerun()
        return
    if prof["papel"] == "controladoria":
        aba = st.sidebar.radio("Menu", ["Justificativas", "Importar dados"])
        st.sidebar.divider()
        if aba == "Importar dados":
            tela_importar(c)
            return
    tela_justificativas(c, prof)

main()
