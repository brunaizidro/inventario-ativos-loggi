import re
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

APP_TZ = ZoneInfo("America/Sao_Paulo")
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

LOGGI_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], button, input, select, textarea { font-family: 'Montserrat', Arial, sans-serif; }
[data-testid="stAppViewContainer"] { background:#f7fafc; }
[data-testid="stHeader"] { background:rgba(255,255,255,.88); }
.block-container { max-width:1440px; padding-top:1.2rem; }
.logo-title { color:#002766; font-weight:800; font-size:1.35rem; margin:0; }
.logo-subtitle { color:#667085; font-size:.78rem; margin-top:.1rem; }
.card-title { color:#002766; font-weight:800; font-size:1.05rem; margin-bottom:.1rem; }
.card-subtitle { color:#667085; font-size:.72rem; }
.kpi { background:#fff; border:1px solid #e5eaf0; border-radius:14px; padding:16px 18px; box-shadow:0 8px 24px rgba(0,39,102,.05); }
.kpi-blue { background:linear-gradient(135deg,#00BAFF,#006AFF); border:0; color:#fff; }
.kpi-label { font-size:.68rem; font-weight:600; color:#667085; }
.kpi-blue .kpi-label { color:#fff; }
.kpi-value { font-size:1.45rem; font-weight:800; color:#002766; margin-top:4px; }
.kpi-blue .kpi-value { color:#fff; }
[data-testid="stSidebar"] { background:#002766; }
[data-testid="stSidebar"] * { color:#fff !important; }
hr { border-color:#e5eaf0 !important; }
</style>
"""

REQUIRED_HEADERS = {
    "BASES": ["NOME_BASE", "DESCRICAO", "STATUS", "DATA_CADASTRO"],
    "USUARIOS": ["EMAIL", "NOME", "PERFIL", "STATUS", "DATA_CADASTRO"],
    "PERMISSOES": ["EMAIL", "NOME_BASE", "STATUS", "DATA_CADASTRO"],
    "ATIVOS_ESPERADOS": ["CODIGO_ATIVO", "TIPO_ATIVO", "NOME_BASE", "STATUS", "ORIGEM", "DATA_CADASTRO", "ULTIMA_ATUALIZACAO"],
    "INVENTARIOS": ["NOME_BASE", "TIPO_INVENTARIO", "SEMANA_REFERENCIA", "DATA_INICIO", "DATA_FIM", "RESPONSAVEL", "STATUS", "TOTAL_BIPS", "OBSERVACAO"],
    "BIPS": ["NOME_BASE", "TIPO_INVENTARIO", "SEMANA_REFERENCIA", "CODIGO_ATIVO", "TIPO_ATIVO", "STATUS_ATIVO", "DATA_HORA", "RESPONSAVEL", "RESULTADO"],
}


def now_local():
    return datetime.now(APP_TZ)


def normalize(value: object) -> str:
    return str(value or "").strip().upper()


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    url = str(st.secrets.get("SPREADSHEET_URL", "")).strip()
    if not url:
        st.error("SPREADSHEET_URL não configurada nos Secrets do Streamlit.")
        st.stop()
    block = st.secrets.get("google_service_account")
    if not block:
        st.error("[google_service_account] não configurado nos Secrets do Streamlit.")
        st.stop()
    info = block.to_dict() if hasattr(block, "to_dict") else dict(block)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds).open_by_url(url)


def ws(name: str):
    try:
        return get_spreadsheet().worksheet(name)
    except gspread.WorksheetNotFound:
        st.error(f"A aba {name} não foi encontrada na planilha.")
        st.stop()


def read_df(name: str) -> pd.DataFrame:
    values = ws(name).get_all_values()
    if not values:
        return pd.DataFrame(columns=REQUIRED_HEADERS.get(name, []))
    return pd.DataFrame(values[1:], columns=values[0])


def ensure_sheet_structure():
    ss = get_spreadsheet()
    for name, headers in REQUIRED_HEADERS.items():
        try:
            sheet = ss.worksheet(name)
        except gspread.WorksheetNotFound:
            sheet = ss.add_worksheet(title=name, rows=1000, cols=max(12, len(headers)))
        current = sheet.row_values(1)
        if current[:len(headers)] != headers:
            sheet.update("A1", [headers])


def current_week() -> str:
    return f"Semana {now_local().isocalendar().week}"


def auth_enabled() -> bool:
    return bool(st.secrets.get("auth")) and hasattr(st, "user")


def current_user_email() -> str:
    if auth_enabled() and st.user.is_logged_in:
        return str(st.user.email or "").lower().strip()
    return str(st.secrets.get("DEV_EMAIL", "")).lower().strip()


def login_gate():
    if auth_enabled():
        if not st.user.is_logged_in:
            st.markdown("# Inventário de Ativos")
            st.write("Sistema interno. Entre com sua conta Google para continuar.")
            st.button("Entrar com Google", on_click=st.login, use_container_width=False)
            st.stop()
    if not current_user_email():
        st.warning("Para o primeiro teste, configure DEV_EMAIL nos Secrets ou o login Google/OIDC de produção.")
        st.stop()


def user_record(email: str):
    df = read_df("USUARIOS")
    if df.empty or "EMAIL" not in df.columns:
        return None
    hit = df[df["EMAIL"].astype(str).str.strip().str.lower() == email.lower()]
    if hit.empty:
        return None
    row = hit.iloc[0].to_dict()
    if normalize(row.get("STATUS")) not in {"ATIVO", "ATIVA"}:
        return None
    return {"email": email, "nome": str(row.get("NOME", "")), "perfil": str(row.get("PERFIL", ""))}


def authorized_bases(email: str):
    user = user_record(email)
    if not user:
        return []
    bases_df = read_df("BASES")
    active = []
    if not bases_df.empty:
        for _, row in bases_df.iterrows():
            if str(row.get("NOME_BASE", "")).strip() and normalize(row.get("STATUS")) in {"ATIVO", "ATIVA"}:
                active.append(str(row["NOME_BASE"]).strip())
    if normalize(user.get("PERFIL")) in {"ADMIN", "ADMINISTRADOR"}:
        return active
    perm = read_df("PERMISSOES")
    allowed = set()
    if not perm.empty:
        for _, row in perm.iterrows():
            if str(row.get("EMAIL", "")).strip().lower() == email.lower() and normalize(row.get("STATUS")) in {"ATIVO", "ATIVA"}:
                if str(row.get("NOME_BASE", "")).strip():
                    allowed.add(str(row["NOME_BASE"]).strip())
    return [b for b in active if b in allowed]


def get_config(base: str, email: str):
    if base not in authorized_bases(email):
        raise ValueError(f"Você não possui acesso à base: {base}")
    inv = read_df("INVENTARIOS")
    week = current_week()
    has_history = False
    ongoing_type = None
    if not inv.empty:
        for _, row in inv.iterrows():
            if normalize(row.get("NOME_BASE")) != normalize(base):
                continue
            has_history = True
            if str(row.get("SEMANA_REFERENCIA", "")).strip() == week and normalize(row.get("STATUS")) == "EM ANDAMENTO":
                ongoing_type = normalize(row.get("TIPO_INVENTARIO"))
    tipo = ongoing_type or ("SEMANAL" if has_history else "INICIAL")
    return {"tipo": tipo, "descricao": "Inventário Inicial" if tipo == "INICIAL" else "Inventário Semanal", "semana": week, "continuar": bool(ongoing_type)}


def normalize_type(tipo: str):
    return {
        "TABLET":"TABLET", "TAB":"TABLET", "BIP":"BIP",
        "TERMINAL SIMPLES":"TERMINAL SIMPLES", "TERM-SIMP":"TERMINAL SIMPLES", "TERM-SMP":"TERMINAL SIMPLES",
        "CELULAR":"CELULAR", "NOTEBOOK":"NOTEBOOK", "OUTRO":"OUTRO",
    }.get(normalize(tipo))


def type_by_code(code: str, base: str):
    code = normalize(code)
    if not code:
        return None
    exp = read_df("ATIVOS_ESPERADOS")
    if not exp.empty:
        for _, row in exp.iterrows():
            if normalize(row.get("CODIGO_ATIVO")) == code and (not normalize(row.get("NOME_BASE")) or normalize(row.get("NOME_BASE")) == normalize(base)):
                normalized = normalize_type(row.get("TIPO_ATIVO", ""))
                if normalized:
                    return normalized
    for prefixes, tipo in [
        (("TERM-SIMP", "TERM-SMP"), "TERMINAL SIMPLES"), (("BIP",), "BIP"),
        (("TAB",), "TABLET"), (("CEL",), "CELULAR"), (("NOTE", "NTB"), "NOTEBOOK")
    ]:
        if any(code == p or code.startswith(p+"-") or code.startswith(p+"_") for p in prefixes):
            return tipo
    return None


def start_inventory(base: str, tipo: str, responsavel: str, email: str):
    cfg = get_config(base, email)
    if tipo != cfg["tipo"]:
        raise ValueError(f"O tipo de inventário é definido automaticamente para esta base: {cfg['descricao']}.")
    inv = read_df("INVENTARIOS")
    for _, row in inv.iterrows():
        if normalize(row.get("NOME_BASE")) == normalize(base) and str(row.get("SEMANA_REFERENCIA", "")).strip() == cfg["semana"]:
            status = normalize(row.get("STATUS"))
            if status == "FINALIZADO":
                raise ValueError(f"Esta base já realizou o inventário da {cfg['semana']}.")
            if status == "EM ANDAMENTO":
                return True
    ws("INVENTARIOS").append_row([base, cfg["tipo"], cfg["semana"], now_local().strftime("%d/%m/%Y %H:%M:%S"), "", responsavel, "EM ANDAMENTO", 0, ""])
    return False


def register_batch(items, email: str):
    if not items:
        raise ValueError("Nenhum ativo para registrar.")
    base = str(items[0]["nomeBase"]).strip()
    cfg = get_config(base, email)
    bips = read_df("BIPS")
    keys = {(normalize(r.get("NOME_BASE")), str(r.get("SEMANA_REFERENCIA", "")).strip(), normalize(r.get("CODIGO_ATIVO"))) for _, r in bips.iterrows()} if not bips.empty else set()
    rows, duplicated, next_num = [], 0, 0
    for item in items:
        if normalize(item.get("nomeBase")) != normalize(base) or str(item.get("semanaReferencia", "")).strip() != cfg["semana"]:
            raise ValueError("Todos os ativos do lote precisam pertencer ao mesmo inventário.")
        status = normalize(item.get("statusAtivo"))
        if status not in {"FUNCIONANDO", "ESTRAGADO"}:
            raise ValueError("Existe ativo sem condição válida.")
        if item.get("semTag"):
            tipo = normalize_type(item.get("tipoAtivo"))
            if tipo not in {"TABLET","BIP","TERMINAL SIMPLES","CELULAR","NOTEBOOK","OUTRO"}:
                raise ValueError("Tipo inválido para SEM TAG.")
            while True:
                next_num += 1
                code = f"SEM TAG-{next_num:03d}"
                if (normalize(base), cfg["semana"], code) not in keys:
                    break
        else:
            code = normalize(item.get("codigoAtivo"))
            if not code:
                raise ValueError("Bipe ou informe o código do ativo.")
            tipo = type_by_code(code, base)
            if not tipo:
                raise ValueError(f"Não foi possível identificar o tipo da tag {code}.")
            informed = normalize_type(item.get("tipoAtivo"))
            if informed and informed != tipo:
                raise ValueError(f"O tipo do ativo {code} não corresponde à tag.")
        key = (normalize(base), cfg["semana"], code)
        if key in keys:
            duplicated += 1
            continue
        keys.add(key)
        rows.append([base, cfg["tipo"], cfg["semana"], code, tipo, status, now_local().strftime("%d/%m/%Y %H:%M:%S"), item.get("responsavel", ""), "REGISTRADO"])
    if rows:
        ws("BIPS").append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows), duplicated


def get_registered(base: str, week: str, email: str):
    if base not in authorized_bases(email):
        raise ValueError("Acesso não autorizado à base.")
    bips = read_df("BIPS")
    if bips.empty:
        return []
    return [{"row":i+2,"codigo":normalize(r.get("CODIGO_ATIVO")),"tipo":normalize(r.get("TIPO_ATIVO")),"status":normalize(r.get("STATUS_ATIVO"))} for i,r in bips.iterrows() if normalize(r.get("NOME_BASE"))==normalize(base) and str(r.get("SEMANA_REFERENCIA", "")).strip()==week]


def update_status(base: str, week: str, code: str, status: str, email: str):
    if normalize(status) not in {"FUNCIONANDO", "ESTRAGADO"}:
        raise ValueError("A condição deve ser FUNCIONANDO ou ESTRAGADO.")
    if base not in authorized_bases(email):
        raise ValueError("Acesso não autorizado à base.")
    sh = ws("BIPS")
    for idx, row in enumerate(sh.get_all_values()[1:], start=2):
        if normalize(row[0])==normalize(base) and str(row[2]).strip()==week and normalize(row[3])==normalize(code):
            sh.update_cell(idx, 6, normalize(status))
            return
    raise ValueError("Ativo não encontrado no inventário.")


def finish_inventory(base: str, week: str, email: str):
    if base not in authorized_bases(email):
        raise ValueError("Acesso não autorizado à base.")
    sh = ws("INVENTARIOS")
    target = None
    for idx, row in enumerate(sh.get_all_values()[1:], start=2):
        if normalize(row[0])==normalize(base) and str(row[2]).strip()==week and normalize(row[6])=="EM ANDAMENTO":
            target=idx; break
    if not target:
        raise ValueError("Inventário em andamento não encontrado.")
    bips = read_df("BIPS")
    total = 0 if bips.empty else int(((bips["NOME_BASE"].astype(str).str.upper().str.strip()==normalize(base)) & (bips["SEMANA_REFERENCIA"].astype(str).str.strip()==week)).sum())
    sh.update_cell(target, 5, now_local().strftime("%d/%m/%Y %H:%M:%S"))
    sh.update_cell(target, 7, "FINALIZADO")
    sh.update_cell(target, 8, total)
    return total


def latest_inventory_by_base(email: str):
    allowed=set(authorized_bases(email)); inv=read_df("INVENTARIOS")
    rows=[]
    if not inv.empty:
        for _,r in inv.iterrows():
            base=str(r.get("NOME_BASE", "")).strip()
            if base in allowed and normalize(r.get("STATUS"))=="FINALIZADO":
                rows.append({"Base":base,"Semana":str(r.get("SEMANA_REFERENCIA", "")),"Data fim":str(r.get("DATA_FIM", "")),"Status":"REALIZADO","Total":r.get("TOTAL_BIPS",0)})
    df=pd.DataFrame(rows)
    return df.drop_duplicates(subset=["Base"],keep="last").sort_values("Base") if not df.empty else df


def render_header(user):
    a,b=st.columns([4,1])
    with a: st.markdown('<div class="logo-title">Inventário de Ativos</div><div class="logo-subtitle">Controle operacional</div>',unsafe_allow_html=True)
    with b: st.markdown(f"**{user.get('nome') or user.get('email','Usuário')}**<br><span style='font-size:.7rem;color:#667085'>{user.get('perfil','')}</span>",unsafe_allow_html=True)


def render_dashboard(email):
    user=user_record(email); render_header(user)
    st.markdown("## Dashboard"); st.caption("Visão consolidada dos inventários realizados por base")
    df=latest_inventory_by_base(email); bases=len(df); latest_week=df["Semana"].iloc[-1] if not df.empty else "-"
    c1,c2,c3,c4=st.columns(4)
    for c,label,val,blue in [(c1,"Bases com inventário",bases,False),(c2,"Inventários realizados",bases,True),(c3,"Semana mais recente",latest_week,False),(c4,"Atualização",now_local().strftime('%H:%M'),False)]:
        cls='kpi kpi-blue' if blue else 'kpi'; c.markdown(f'<div class="{cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div></div>',unsafe_allow_html=True)
    st.divider(); st.markdown('<div class="card-title">Inventários por base</div><div class="card-subtitle">Último inventário finalizado de cada base</div>',unsafe_allow_html=True)
    if df.empty: st.info("Nenhum inventário finalizado ainda.")
    else: st.dataframe(df,hide_index=True,use_container_width=True)


def render_inventory(email):
    user=user_record(email); render_header(user)
    st.markdown("## Inventário"); st.caption("Registre os ativos da base e acompanhe a execução do inventário.")
    bases=authorized_bases(email)
    if not bases: st.warning("Nenhuma base autorizada para este usuário."); return
    base=st.selectbox("Base",bases,key="base_select")
    cfg=get_config(base,email); registered=get_registered(base,cfg["semana"],email)
    c1,c2,c3,c4=st.columns(4); c1.metric("Tipo",cfg["tipo"]); c2.metric("Semana",cfg["semana"]); c3.metric("Status","Em andamento" if cfg["continuar"] or st.session_state.get("inventario_iniciado") else "Não iniciado"); c4.metric("Ativos registrados",len(registered))
    if st.button("Iniciar / continuar inventário",type="primary",use_container_width=True):
        try:
            continued=start_inventory(base,cfg["tipo"],user["nome"],email); st.session_state["inventario_iniciado"]=True; st.session_state["base_ativa"]=base
            st.success("Inventário em andamento localizado." if continued else f"{cfg['descricao']} iniciado com sucesso."); st.rerun()
        except Exception as exc: st.error(str(exc))
    if not st.session_state.get("inventario_iniciado") and not cfg["continuar"]:
        st.info("Inicie o inventário para liberar a bipagem de ativos."); return
    st.session_state["inventario_iniciado"]=True
    st.divider(); st.markdown('<div class="card-title">Registro rápido</div><div class="card-subtitle">Bipe uma tag ou selecione SEM TAG.</div>',unsafe_allow_html=True)
    no_tag=st.checkbox("SEM TAG",key="sem_tag")
    if no_tag:
        tipo=st.selectbox("Tipo do ativo",["BIP","TERMINAL SIMPLES","TABLET","CELULAR","NOTEBOOK","OUTRO"],key="manual_type"); code=""
    else:
        code=st.text_input("Código do ativo",placeholder="Ex.: BIP-12.5415 / TERM-SMP-10.2301 / TAB-13.0096",key="asset_code")
        tipo=type_by_code(code,base) if code else ""
        if code: st.caption(f"Tipo identificado: **{tipo or 'não identificado'}**")
    status=st.radio("Condição",["FUNCIONANDO","ESTRAGADO"],horizontal=True,key="asset_status")
    if st.button("Registrar ativo",use_container_width=True):
        try:
            n,d=register_batch([{ "nomeBase":base,"semTag":no_tag,"codigoAtivo":code,"tipoAtivo":tipo,"statusAtivo":status,"semanaReferencia":cfg["semana"],"responsavel":user["nome"] }],email)
            if n: st.success("Ativo registrado."); st.rerun()
            else: st.warning("Ativo duplicado nesta semana.")
        except Exception as exc: st.error(str(exc))
    registered=get_registered(base,cfg["semana"],email); st.divider(); st.markdown('<div class="card-title">Ativos registrados</div>',unsafe_allow_html=True)
    if registered:
        st.dataframe(pd.DataFrame([{ "Ativo":x["codigo"],"Tipo":x["tipo"],"Condição":x["status"] } for x in registered]),hide_index=True,use_container_width=True)
        selected=st.selectbox("Ativo",[x["codigo"] for x in registered],key="edit_asset"); new_status=st.selectbox("Nova condição",["FUNCIONANDO","ESTRAGADO"],key="edit_status")
        if st.button("Salvar condição"):
            try: update_status(base,cfg["semana"],selected,new_status,email); st.success("Condição atualizada."); st.rerun()
            except Exception as exc: st.error(str(exc))
    else: st.info("Nenhum ativo registrado ainda.")
    st.divider(); st.markdown(f"**Total de ativos nesta semana: {len(registered)}**")
    if st.button("Finalizar inventário",type="primary",use_container_width=True):
        try:
            total=finish_inventory(base,cfg["semana"],email); st.session_state["inventario_iniciado"]=False; st.success(f"Inventário finalizado com sucesso. Total de ativos: {total}."); st.rerun()
        except Exception as exc: st.error(str(exc))


def main():
    st.set_page_config(page_title="Inventário de Ativos",page_icon="▣",layout="wide",initial_sidebar_state="expanded")
    st.markdown(LOGGI_CSS,unsafe_allow_html=True); login_gate(); email=current_user_email(); user=user_record(email)
    if not user:
        st.error("Seu e-mail não está autorizado ou seu usuário está inativo na aba USUARIOS.")
        if auth_enabled() and st.user.is_logged_in: st.sidebar.button("Sair",on_click=st.logout)
        st.stop()
    try: ensure_sheet_structure()
    except Exception as exc: st.error(f"Falha ao conectar/configurar a planilha: {exc}"); st.stop()
    with st.sidebar:
        st.markdown("### Inventário de Ativos"); page=st.radio("Menu",["Dashboard","Inventário"],label_visibility="collapsed"); st.divider(); st.caption(user.get("nome","Usuário")); st.caption(email); st.caption(user.get("perfil",""))
        if auth_enabled() and st.user.is_logged_in: st.button("Sair",on_click=st.logout,use_container_width=True)
    render_dashboard(email) if page=="Dashboard" else render_inventory(email)


if __name__ == "__main__":
    main()
