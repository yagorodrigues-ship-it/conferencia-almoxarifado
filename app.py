import streamlit as st
import pandas as pd
import datetime
import psycopg2
import psycopg2.extras
import io
import re
import time
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill, Font

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Auditoria & Inventário - JBA", layout="wide")

# --- LISTA OFICIAL E UNIFICADA DE ESTOQUES JBA ---
LISTA_ESTOQUES_FIXA = [
    {"id": "1077", "desc": "JBA - CLASSE D"}, {"id": "1078", "desc": "JBA - COPA E COZINHA"},
    {"id": "1080", "desc": "JBA - DADOS - CLIENTE"}, {"id": "1082", "desc": "JBA - VIVO VITA - CLIENTE"},
    {"id": "1084", "desc": "JBA - EPI-EPC"}, {"id": "1086", "desc": "JBA - EQUIPAMENTOS"},
    {"id": "1088", "desc": "JBA - FERRAMENTAL"}, {"id": "1089", "desc": "JBA - KIT FERRAMENTAL CONTRATACOES"},
    {"id": "1090", "desc": "JBA - FERRAMENTAS DE CANTEIRO"}, {"id": "1104", "desc": "JBA - MATERIAL DE ESCRITORIO - SUPRIMENTOS DE INFORMATICA"},
    {"id": "1106", "desc": "JBA - MOBILIARIO"}, {"id": "1113", "desc": "1385 - MANUTENCAO JBA - CLIENTE"},
    {"id": "1118", "desc": "JBA - PROPRIO GERAL"}, {"id": "1122", "desc": "JBA - GRAND OBRAS IMPLANTACAO"},
    {"id": "1140", "desc": "JBA - SPEEDY/FTTX - CLIENTE"}, {"id": "1144", "desc": "1385 - MANUTENCAO JBA CLIENTE RESERVADO"},
    {"id": "1149", "desc": "JBA - UNIFORME"}, {"id": "2183", "desc": "1071 - BOL IMPLANTANCAO JBA - CLIENTE"},
    {"id": "2185", "desc": "JBA - PROPRIO FATURA B PLANTA EXTERNA - BDI"}, {"id": "2188", "desc": "1071 - IMPLANTACAO JBA CLIENTE RESERVADO"},
    {"id": "2190", "desc": "JBA - DEPARTAMENTO T.I"}, {"id": "2194", "desc": "JBA - KITS FERRAMENTAL - DEVOLUCAO"},
    {"id": "2197", "desc": "JBA - EQUIPAMENTOS TI"}, {"id": "2641", "desc": "1259 - IMPLANTACAO JBA - MATERIAL REUTILIZACAO"},
    {"id": "2643", "desc": "1724 - MANUTENCAO JBA - MATERIAL REUTILIZACAO"}, {"id": "2725", "desc": "JBA - RESERVA TIM"},
    {"id": "2983", "desc": "JBA - FORNECEDORES P/ MANUTENCAO - RECARGA"}, {"id": "3193", "desc": "JBA - PROPRIO MATERIAL REAPROVEITAVEL"},
    {"id": "3395", "desc": "LPA - FTTX - CLIENTE"}, {"id": "3484", "desc": "JBA - CELULARES DEFEITO"},
    {"id": "3546", "desc": "JBA - CELULARES"}
]
MAPA_ESTOQUES_DESC = {item['id']: item['desc'] for item in LISTA_ESTOQUES_FIXA}

# --- CONEXÃO SEGURA COM O POSTGRESQL ---
def conectar_banco():
    try:
        db_url = st.secrets["postgres"]["url"]
        return psycopg2.connect(db_url, sslmode='require', connect_timeout=10)
    except Exception as e:
        st.error(f"❌ Erro de Conexão com o Banco: {e}")
        st.stop()

@st.cache_data(ttl=10)
def buscar_inventarios_cache():
    conn = conectar_banco()
    cursor = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    cursor.execute("SELECT id, nome, data, status FROM inventarios ORDER BY data DESC, CAST(REPLACE(id, '#', '') AS INTEGER) DESC;")
    rows = cursor.fetchall()
    conn.close()
    return pd.DataFrame(rows, columns=['id', 'nome', 'data', 'status']) if rows else pd.DataFrame()

@st.cache_data(ttl=10)
def buscar_historico_estoques_cache():
    conn = conectar_banco()
    df_h = pd.read_sql_query("SELECT id_estoque, ultima_data FROM ultima_contagem_estoques", conn)
    df_c = pd.read_sql_query("SELECT id_estoque, MAX(data_hora) as ultima_data FROM contagens WHERE id_estoque IS NOT NULL AND id_estoque != '' GROUP BY id_estoque", conn)
    conn.close()
    return df_h, df_c

@st.cache_data(ttl=300)
def carregar_base_em_memoria(id_pasta):
    conn = conectar_banco()
    df = pd.read_sql_query("SELECT * FROM itens_base_inventario WHERE inventario_id = %s OR inventario_id = %s", conn, params=(id_pasta, f"#{id_pasta}"))
    conn.close()
    return df if not df.empty else None

def limpar_cache_aplicacao():
    st.cache_data.clear()

def limpar_documento(doc):
    return str(doc).strip().replace(".", "").replace("-", "").replace("/", "")

def extrair_id_estoque_do_nome(nome_inventario):
    numeros = re.findall(r'\b\d{4}\b', str(nome_inventario))
    for num in numeros:
        if num in MAPA_ESTOQUES_DESC:
            return num
    return ""

def inicializar_banco():
    conn = conectar_banco()
    cursor = conn.cursor()
    tabelas = [
        "CREATE TABLE IF NOT EXISTS usuarios (id SERIAL PRIMARY KEY, nome TEXT, cpf TEXT UNIQUE, email TEXT UNIQUE, senha TEXT, perfil TEXT DEFAULT 'Almoxarife');",
        "CREATE TABLE IF NOT EXISTS inventarios (id TEXT PRIMARY KEY, nome TEXT, data TEXT, status TEXT, total_itens INTEGER DEFAULT 0, acuracidade_final TEXT DEFAULT '0%');",
        "CREATE TABLE IF NOT EXISTS itens_base_inventario (id SERIAL PRIMARY KEY, inventario_id TEXT, cod_produto TEXT, desc_produto TEXT, desc_estoque_fisico TEXT, unid_medida TEXT, qtd_estoque INTEGER, id_estoque_fisico TEXT, lote TEXT DEFAULT '', ativo TEXT DEFAULT '');",
        "CREATE TABLE IF NOT EXISTS inventarios_supervisor (id TEXT PRIMARY KEY, nome TEXT, data TEXT, status TEXT);",
        "CREATE TABLE IF NOT EXISTS contagens (id SERIAL PRIMARY KEY, inventario_id TEXT, id_estoque TEXT, desc_estoque TEXT, cod_produto TEXT, desc_produto TEXT, unid_medida TEXT, qtd_sistema INTEGER, qtd_contada INTEGER, diferenca INTEGER, ativo TEXT, observacao TEXT, operador TEXT, data_hora TEXT, lote TEXT, fase_contagem TEXT DEFAULT '1a Contagem');",
        "CREATE TABLE IF NOT EXISTS auditorias_supervisor (id SERIAL PRIMARY KEY, inventario_id TEXT, id_estoque TEXT, desc_estoque TEXT, cod_produto TEXT, desc_produto TEXT, qtd_sistema INTEGER, qtd_auditada INTEGER, diferenca INTEGER, etiqueta_correta TEXT, localizacao_correta TEXT, supervisor TEXT, data_hora TEXT, recontagem_3 TEXT DEFAULT 'Não', ativo TEXT);",
        "CREATE TABLE IF NOT EXISTS ultima_contagem_estoques (id_estoque TEXT PRIMARY KEY, ultima_data TEXT);"
    ]
    for t in tabelas:
        try:
            cursor.execute(t)
            conn.commit()
        except Exception: conn.rollback()

    alteracoes = [
        "ALTER TABLE inventarios ADD COLUMN IF NOT EXISTS total_itens INTEGER DEFAULT 0;",
        "ALTER TABLE inventarios ADD COLUMN IF NOT EXISTS acuracidade_final TEXT DEFAULT '0%';",
        "ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS perfil TEXT DEFAULT 'Almoxarife';"
    ]
    for alt in alteracoes:
        try:
            cursor.execute(alt)
            conn.commit()
        except Exception: conn.rollback()

    try:
        cursor.execute("SELECT COUNT(*) FROM usuarios WHERE email = 'admin@tel.com.br';")
        if cursor.fetchone()[0] == 0:
            cursor.execute("INSERT INTO usuarios (nome, cpf, email, senha, perfil) VALUES (%s, %s, %s, %s, %s) ON CONFLICT DO NOTHING;",
                           ("Administrador Tel", "00000000000", "admin@tel.com.br", "123", "Administrador"))
            conn.commit()
    except Exception: conn.rollback()
    conn.close()

inicializar_banco()

def converter_para_excel(df):
    output = io.BytesIO()
    colunas_ordenadas = [
        'id', 'inventario_id', 'id_estoque', 'desc_estoque', 'cod_produto',
        'desc_produto', 'unid_medida', 'qtd_sistema', 'qtd_contada', 'diferenca',
        'ativo', 'observacao', 'operador', 'data_hora', 'lote', 'fase_contagem'
    ]
    cols_finais = [c for c in colunas_ordenadas if c in df.columns]
    df_f = df[cols_finais] if cols_finais else df
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_f.to_excel(writer, index=False, sheet_name='Relatorio')
    return output.getvalue()

def gerar_relatorio_consolidado_excel(df_contagens, lista_estoques_ref):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        linhas_resumo = []
        for est in lista_estoques_ref:
            est_id = str(est["id"]).strip()
            df_est = df_contagens[df_contagens['id_estoque'].astype(str).str.replace('.0', '', regex=False).str.strip() == est_id] if not df_contagens.empty and 'id_estoque' in df_contagens.columns else pd.DataFrame()
            
            total_itens = len(df_est)
            acertos = len(df_est[pd.to_numeric(df_est['diferenca'], errors='coerce') == 0]) if total_itens > 0 else 0
            erros = len(df_est[pd.to_numeric(df_est['diferenca'], errors='coerce') != 0]) if total_itens > 0 else 0
            qtd_segunda = len(df_est[df_est['fase_contagem'].astype(str).str.contains('2a Contagem')]) if 'fase_contagem' in df_est.columns and total_itens > 0 else 0
            acuracidade_pct = (acertos / total_itens) if total_itens > 0 else 0.0
            u_data_est = str(df_est['data_hora'].max()) if not df_est.empty and 'data_hora' in df_est.columns else "—"
            
            obs_erros = []
            if not df_est.empty and 'observacao' in df_est.columns:
                obs_validas = df_est[df_est['observacao'].fillna('').astype(str).str.strip() != '']['observacao'].unique().tolist()
                if obs_validas: obs_erros.append("; ".join(map(str, obs_validas)))
            if erros > 0 and not df_est.empty:
                erros_cods = df_est[pd.to_numeric(df_est['diferenca'], errors='coerce') != 0]['cod_produto'].unique().tolist()
                obs_erros.append(f"Divergências: {', '.join(map(str, erros_cods[:5]))}" + ("..." if len(erros_cods) > 5 else ""))
            
            linhas_resumo.append({
                "Id. Estoq. Físico": est_id, "Desc. Estoque Físico": est["desc"],
                "Total de itens": total_itens, "Acertos": acertos, "Erros": erros,
                "2º Contagem": qtd_segunda, "Acuracidade": acuracidade_pct,
                "Última Contagem": u_data_est, "Obs": " | ".join(obs_erros) if obs_erros else "Nenhuma observação"
            })
            
        df_resumo = pd.DataFrame(linhas_resumo)
        df_resumo.to_excel(writer, index=False, sheet_name='Acuracidade')
        
        worksheet = writer.sheets['Acuracidade']
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        header_fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        for col_idx in range(1, len(df_resumo.columns) + 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font

        for row in range(2, len(df_resumo) + 2):
            worksheet[f'G{row}'].number_format = '0.00%'
            cell_erros = worksheet[f'E{row}']
            if cell_erros.value and int(cell_erros.value) > 0: cell_erros.fill = red_fill
            elif cell_erros.value == 0 and worksheet[f'C{row}'].value and int(worksheet[f'C{row}'].value) > 0: cell_erros.fill = green_fill

        if not df_contagens.empty and 'id_estoque' in df_contagens.columns:
            for est_id, group in df_contagens.groupby('id_estoque'):
                nome_aba = str(est_id).replace('.0', '').strip()[:31]
                if nome_aba:
                    group_limpo = group.drop(columns=['id'], errors='ignore')
                    group_limpo.to_excel(writer, index=False, sheet_name=nome_aba)

    return output.getvalue()

# --- ESTILOS CSS ---
st.markdown("""
    <style>
    div.stButton > button:first-child[kind="primary"] { background-color: #d35400; border-color: #d35400; color: white; font-weight: bold; }
    div.stButton > button:first-child[kind="primary"]:hover { background-color: #e67e22; border-color: #e67e22; color: white; }
    .card-lateral { background-color: #1a233a; padding: 10px; border-radius: 8px; margin-bottom: 8px; border-left: 4px solid #2563eb; }
    .card-lateral-titulo { color: #93c5fd; font-size: 11px; font-weight: bold; }
    .card-lateral-valor { color: white; font-size: 22px; font-weight: bold; }
    .caixa-descricao-produto { background-color: #ebf8ff; border-left: 5px solid #3182ce; padding: 12px 16px; border-radius: 6px; font-size: 16px; font-weight: bold; color: #1a365d; margin: 10px 0; }
    </style>
""", unsafe_allow_html=True)

# INICIALIZAÇÃO DE ESTADOS
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'operador' not in st.session_state: st.session_state.operador = ""
if 'perfil_usuario' not in st.session_state: st.session_state.perfil_usuario = "Almoxarife"
if 'tela_acesso' not in st.session_state: st.session_state.tela_acesso = "login"
if 'contador_reset' not in st.session_state: st.session_state.contador_reset = 0
if 'contador_reset_sup' not in st.session_state: st.session_state.contador_reset_sup = 0
if 'bases_supervisor_por_inv' not in st.session_state: st.session_state.bases_supervisor_por_inv = {}
if 'pagina_historico' not in st.session_state: st.session_state.pagina_historico = 1

# MEMÓRIA RAM LOCAL PARA CONTAGEM EM LOTE
if 'buffer_ram_contagens' not in st.session_state:
    st.session_state.buffer_ram_contagens = []
if 'ultima_sincronizacao' not in st.session_state:
    st.session_state.ultima_sincronizacao = time.time()

# FUNÇÃO DE SINCRONIZAÇÃO EM LOTE RAM -> BANCO
def sincronizar_ram_com_banco():
    if not st.session_state.buffer_ram_contagens:
        return 0
    
    conn = conectar_banco()
    cursor = conn.cursor()
    
    dados_para_inserir = []
    for item in st.session_state.buffer_ram_contagens:
        dados_para_inserir.append((
            item['inventario_id'], item['id_estoque'], item['desc_estoque'],
            item['cod_produto'], item['desc_produto'], item['unid_medida'],
            item['qtd_sistema'], item['qtd_contada'], item['diferenca'],
            item['ativo'], item['observacao'], item['operador'],
            item['data_hora'], item['lote'], item['fase_contagem']
        ))
        if item['id_estoque']:
            cursor.execute("INSERT INTO ultima_contagem_estoques (id_estoque, ultima_data) VALUES (%s, %s) ON CONFLICT (id_estoque) DO UPDATE SET ultima_data = EXCLUDED.ultima_data;", (item['id_estoque'], item['data_hora']))

    psycopg2.extras.execute_values(
        cursor,
        "INSERT INTO contagens (inventario_id, id_estoque, desc_estoque, cod_produto, desc_produto, unid_medida, qtd_sistema, qtd_contada, diferenca, ativo, observacao, operador, data_hora, lote, fase_contagem) VALUES %s",
        dados_para_inserir
    )
    
    conn.commit()
    conn.close()
    
    qtd_salva = len(st.session_state.buffer_ram_contagens)
    st.session_state.buffer_ram_contagens = [] # Reseta o buffer da RAM
    st.session_state.ultima_sincronizacao = time.time()
    limpar_cache_aplicacao()
    return qtd_salva

# --- TELA DE LOGIN CENTRALIZADA ---
if not st.session_state.logged_in:
    conn = conectar_banco()
    col_vaz1, col_central, col_vaz2 = st.columns([1, 1.2, 1])
    with col_central:
        if st.session_state.tela_acesso == "login":
            st.title("🔒 Acesso ao Sistema JBA")
            with st.form("login_form"):
                identificador = st.text_input("CPF ou E-mail")
                senha = st.text_input("Senha", type="password")
                if st.form_submit_button("Entrar no Sistema", type="primary", use_container_width=True):
                    id_limpo = identificador.strip()
                    doc_limpo = limpar_documento(id_limpo)
                    cursor = conn.cursor()
                    cursor.execute("SELECT nome, perfil FROM usuarios WHERE (email = %s OR cpf = %s) AND senha = %s", (id_limpo, doc_limpo, senha))
                    user = cursor.fetchone()
                    conn.close()
                    if user:
                        st.session_state.logged_in = True
                        st.session_state.operador = user[0]
                        st.session_state.perfil_usuario = user[1] or "Almoxarife"
                        limpar_cache_aplicacao()
                        st.rerun()
                    else: st.error("❌ Credenciais incorretas.")
            if st.button("📝 Criar nova conta de colaborador", use_container_width=True):
                conn.close()
                st.session_state.tela_acesso = "cadastro"
                st.rerun()

        elif st.session_state.tela_acesso == "cadastro":
            st.title("📝 Cadastro de Colaborador")
            with st.form("cadastro_form"):
                novo_nome = st.text_input("Nome Completo")
                novo_cpf = st.text_input("CPF (Apenas números)")
                novo_email = st.text_input("E-mail")
                nova_senha = st.text_input("Senha", type="password")
                confirma_senha = st.text_input("Confirme a Senha", type="password")
                if st.form_submit_button("Finalizar Cadastro", type="primary", use_container_width=True):
                    cpf_l = limpar_documento(novo_cpf)
                    if not novo_nome or not cpf_l or not novo_email or not nova_senha: st.error("⚠️ Preencha todos os campos!")
                    elif nova_senha != confirma_senha: st.error("❌ Senhas divergentes!")
                    else:
                        try:
                            cursor = conn.cursor()
                            cursor.execute("INSERT INTO usuarios (nome, cpf, email, senha, perfil) VALUES (%s, %s, %s, %s, 'Almoxarife')", (novo_nome.strip(), cpf_l, novo_email.strip(), nova_senha))
                            conn.commit()
                            conn.close()
                            st.success("✅ Cadastro realizado!")
                            st.session_state.tela_acesso = "login"
                            st.rerun()
                        except Exception: 
                            conn.close()
                            st.error("❌ CPF ou E-mail já cadastrado.")
            if st.button("◀ Voltar para o Login"):
                conn.close()
                st.session_state.tela_acesso = "login"
                st.rerun()

# --- APLICAÇÃO PRINCIPAL ---
else:
    conn = conectar_banco()
    df_inventarios = buscar_inventarios_cache()
    df_inventarios_sup = pd.read_sql_query("SELECT * FROM inventarios_supervisor ORDER BY data DESC, id DESC", conn)
    eh_supervisor = (st.session_state.perfil_usuario == "Administrador") or ("admin" in st.session_state.operador.lower())

    # AUTO-SINCRONIZAÇÃO A CADA 2 HORAS (7200 SEGUNDOS)
    tempo_passado = time.time() - st.session_state.ultima_sincronizacao
    if tempo_passado >= 7200 and st.session_state.buffer_ram_contagens:
        qtd_auto = sincronizar_ram_com_banco()
        st.toast(f"🔄 Auto-sincronização realizada: {qtd_auto} itens salvos no banco!")

    # --- BARRA LATERAL ---
    with st.sidebar:
        st.write(f"👤 **{st.session_state.operador}** ({st.session_state.perfil_usuario})")
        
        # STATUS DA MEMÓRIA RAM
        pendentes_ram = len(st.session_state.buffer_ram_contagens)
        if pendentes_ram > 0:
            st.warning(f"⚡ **{pendentes_ram} bips** guardados na RAM local!")
        else:
            st.success("🟢 Memória RAM sincronizada.")

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            if st.button("🔄 Atualizar", use_container_width=True):
                limpar_cache_aplicacao()
                st.rerun()
        with col_s2:
            if st.button("🚪 Sair", use_container_width=True):
                if pendentes_ram > 0:
                    sincronizar_ram_com_banco()
                st.session_state.logged_in = False
                st.session_state.operador = ""
                limpar_cache_aplicacao()
                st.rerun()
            
        st.markdown("---")
        st.write("📁 **Seleção de Inventário**")
        if df_inventarios.empty:
            id_inventario_atual = None
            inventario_selected_obj = None
            id_pasta_limpo_base = ""
            st.info("Crie um inventário abaixo.")
        else:
            lista_inv = [f"{row['id']} – {row['nome']} ({row['status']})" for idx, row in df_inventarios.iterrows()]
            inventario_selected = st.selectbox("Selecione a Pasta", lista_inv, index=0, key="sb_pasta_ativa")
            id_inventario_atual = inventario_selected.split(" – ")[0]
            inventario_selected_obj = df_inventarios[df_inventarios['id'] == id_inventario_atual].iloc[0]
            id_pasta_limpo_base = id_inventario_atual.replace("#", "").strip()

        # UPLOAD DA BASE
        st.write("📂 **Upload Base de Dados (.xlsx)**")
        if inventario_selected_obj is not None and inventario_selected_obj['status'] == "Aberto":
            uploader_key = f"func_excel_loader_{id_pasta_limpo_base if id_pasta_limpo_base else 'vazio'}_{st.session_state.contador_reset}"
            arquivo_excel = st.file_uploader("Suba o arquivo Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed", key=uploader_key)
            
            if arquivo_excel is not None and id_pasta_limpo_base:
                with st.spinner("🚀 Processando base..."):
                    df_upload_temp = pd.read_excel(arquivo_excel)
                    mapa_cols_upload = {str(col).replace('\xa0', ' ').strip().lower().replace(" ", "").replace(".", "").replace("ó", "o").replace("á", "a"): col for col in df_upload_temp.columns}

                    col_cod, col_desc = mapa_cols_upload.get("codproduto"), mapa_cols_upload.get("descproduto")
                    col_est, col_uni = mapa_cols_upload.get("descestoquefisico"), mapa_cols_upload.get("unidmedida")
                    col_qtd, col_idest = mapa_cols_upload.get("qtdestoque"), mapa_cols_upload.get("idestoquefisico")
                    col_lote, col_ativo = mapa_cols_upload.get("lote"), mapa_cols_upload.get("ativo")

                    est_id_fallback = extrair_id_estoque_do_nome(inventario_selected_obj['nome'])
                    est_desc_fallback = MAPA_ESTOQUES_DESC.get(est_id_fallback, 'ESTOQUE FÍSICO')

                    if not col_cod or not col_desc:
                        st.error("❌ Planilha fora do padrão JBA.")
                    else:
                        cursor_db = conn.cursor()
                        cursor_db.execute("DELETE FROM itens_base_inventario WHERE inventario_id = %s OR inventario_id = %s", (id_pasta_limpo_base, f"#{id_pasta_limpo_base}"))
                        dados_para_inserir = []
                        for _, r in df_upload_temp.iterrows():
                            v_cod = str(r[col_cod]).strip() if col_cod and pd.notna(r[col_cod]) else ''
                            v_desc = str(r[col_desc]).strip() if col_desc and pd.notna(r[col_desc]) else ''
                            v_est = str(r[col_est]).strip() if col_est and pd.notna(r[col_est]) and str(r[col_est]).strip() != '' else est_desc_fallback
                            v_uni = str(r[col_uni]).strip() if col_uni and pd.notna(r[col_uni]) else ''
                            v_qtd = int(pd.to_numeric(r[col_qtd], errors='coerce') or 0) if col_qtd else 0
                            v_idest = str(r[col_idest]).strip() if col_idest and pd.notna(r[col_idest]) and str(r[col_idest]).strip() != '' else est_id_fallback
                            v_lote = str(r[col_lote]).strip() if col_lote and pd.notna(r[col_lote]) and str(r[col_lote]).lower() != 'nan' else ''
                            v_ativo = str(r[col_ativo]).strip() if col_ativo and pd.notna(r[col_ativo]) and str(r[col_ativo]).lower() != 'nan' else ''

                            dados_para_inserir.append((id_pasta_limpo_base, v_cod, v_desc, v_est, v_uni, v_qtd, v_idest, v_lote, v_ativo))
                        
                        psycopg2.extras.execute_values(cursor_db, "INSERT INTO itens_base_inventario (inventario_id, cod_produto, desc_produto, desc_estoque_fisico, unid_medida, qtd_estoque, id_estoque_fisico, lote, ativo) VALUES %s", dados_para_inserir)
                        conn.commit()
                        limpar_cache_aplicacao()
                        st.success("✅ Base Carregada!")

        elif inventario_selected_obj is not None and inventario_selected_obj['status'] in ["1a Contagem", "2a Contagem"]:
            st.info(f"🔒 **Base Congelada ({inventario_selected_obj['status']})**.")
        else: st.info("🔒 Crie um inventário 'Aberto'.")

        base_sistema_atual = carregar_base_em_memoria(id_pasta_limpo_base) if id_pasta_limpo_base else None

        if inventario_selected_obj is not None and inventario_selected_obj['status'] == "Aberto" and base_sistema_atual is not None:
            st.markdown("---")
            if st.button("🚀 Salvar Base e Iniciar 1ª Contagem", type="primary", use_container_width=True):
                cursor = conn.cursor()
                cursor.execute("UPDATE inventarios SET status = '1a Contagem' WHERE id = %s OR id = %s", (id_inventario_atual, id_pasta_limpo_base))
                conn.commit()
                limpar_cache_aplicacao()
                st.success("🔒 1ª Contagem liberada.")
                st.rerun()

        with st.expander("➕ Criar Novo Inventário"):
            with st.form("form_novo", clear_on_submit=True):
                novo_nome = st.text_input("Nome do Inventário")
                if st.form_submit_button("Criar Pasta", type="primary") and novo_nome:
                    cursor = conn.cursor()
                    df_calc = pd.read_sql_query("SELECT id FROM inventarios", conn)
                    maior_id = df_calc['id'].str.replace('#', '', regex=False).astype(int).max() if not df_calc.empty else 0
                    cursor.execute("INSERT INTO inventarios (id, nome, data, status) VALUES (%s, %s, %s, 'Aberto')", (f"#{maior_id + 1}", novo_nome, datetime.date.today().strftime("%Y-%m-%d")))
                    conn.commit()
                    limpar_cache_aplicacao()
                    st.rerun()

        # FECHAMENTO DO INVENTÁRIO
        pode_fechar, itens_faltantes, itens_pendentes_2a = False, [], []
        if inventario_selected_obj is not None and inventario_selected_obj['status'] in ["1a Contagem", "2a Contagem"] and base_sistema_atual is not None:
            cursor_verif = conn.cursor()
            cursor_verif.execute("SELECT cod_produto, lote, ativo, fase_contagem FROM contagens WHERE inventario_id = %s OR inventario_id = %s", (id_pasta_limpo_base, f"#{id_pasta_limpo_base}"))
            rows_verif = cursor_verif.fetchall()
            
            set_contados_triade = {f"{str(r[0]).upper().strip()}_{str(r[1]).upper().strip() if r[1] and str(r[1]).lower() != 'nan' else ''}_{str(r[2]).upper().strip() if r[2] and str(r[2]).lower() != 'nan' else ''}" for r in rows_verif} if rows_verif else set()

            # Adiciona RAM local ao cálculo de fechamento
            for r_ram in st.session_state.buffer_ram_contagens:
                if r_ram['inventario_id'] == id_pasta_limpo_base:
                    set_contados_triade.add(f"{str(r_ram['cod_produto']).upper().strip()}_{str(r_ram['lote']).upper().strip() if r_ram['lote'] else ''}_{str(r_ram['ativo']).upper().strip() if r_ram['ativo'] else ''}")

            if inventario_selected_obj['status'] == "1a Contagem":
                for _, r_b in base_sistema_atual.iterrows():
                    c_b, l_b, a_b = str(r_b['cod_produto']).upper().strip(), str(r_b['lote']).upper().strip() if pd.notna(r_b['lote']) and str(r_b['lote']).lower() != 'nan' else "", str(r_b['ativo']).upper().strip() if pd.notna(r_b['ativo']) and str(r_b['ativo']).lower() != 'nan' else ""
                    if f"{c_b}_{l_b}_{a_b}" not in set_contados_triade: itens_faltantes.append(c_b)
                if len(itens_faltantes) == 0: pode_fechar = True
            elif inventario_selected_obj['status'] == "2a Contagem":
                itens_pendentes_2a = [str(r[0]).upper().strip() for r in rows_verif if r[3] == '2a Contagem']
                if len(itens_pendentes_2a) == 0: pode_fechar = True

            st.markdown("---")
            def fechar_e_preservar_historico(id_inv, id_limpo):
                if pendentes_ram > 0:
                    sincronizar_ram_com_banco()
                cursor = conn.cursor()
                cursor.execute("SELECT diferenca FROM contagens WHERE inventario_id = %s OR inventario_id = %s", (id_limpo, f"#{id_limpo}"))
                rows_difs = cursor.fetchall()
                tot = len(rows_difs)
                acertos = len([r for r in rows_difs if r[0] == 0]) if tot > 0 else 0
                pct_acu = f"{(acertos / tot)*100:.1f}%" if tot > 0 else "0%"
                try: cursor.execute("UPDATE inventarios SET status = 'Fechado', total_itens = %s, acuracidade_final = %s WHERE id = %s OR id = %s", (tot, pct_acu, id_inv, id_limpo))
                except Exception:
                    conn.rollback()
                    cursor.execute("UPDATE inventarios SET status = 'Fechado' WHERE id = %s OR id = %s", (id_inv, id_limpo))
                conn.commit()
                limpar_cache_aplicacao()

            if pode_fechar:
                if st.button("🔒 Fechar Inventário (100% Concluído)", use_container_width=True, type="primary"):
                    fechar_e_preservar_historico(id_inventario_atual, id_pasta_limpo_base)
                    st.success("✅ Inventário encerrado e arquivado!")
                    st.rerun()
            else:
                qtd_f = len(itens_faltantes) if inventario_selected_obj['status'] == "1a Contagem" else len(itens_pendentes_2a)
                st.warning(f"⏳ Faltam **{qtd_f}** itens para concluir.")
                if eh_supervisor:
                    if st.button("🚨 Forçar Fechamento Incompleto (ADMIN)", use_container_width=True):
                        fechar_e_preservar_historico(id_inventario_atual, id_pasta_limpo_base)
                        st.success("✅ Inventário encerrado!")
                        st.rerun()

        # KPIs SIDEBAR
        total_itens_base = len(base_sistema_atual) if base_sistema_atual is not None else 0
        cursor_side = conn.cursor()
        cursor_side.execute("SELECT COUNT(*) FROM contagens WHERE inventario_id = %s OR inventario_id = %s", (id_pasta_limpo_base, f"#{id_pasta_limpo_base}"))
        total_contados_cnt = (cursor_side.fetchone()[0] if id_pasta_limpo_base else 0) + pendentes_ram
        
        st.markdown(f'<div class="card-lateral"><div class="card-lateral-titulo">📋 ITENS NA BASE</div><div class="card-lateral-valor">{total_itens_base}</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="card-lateral"><div class="card-lateral-titulo">✅ LANÇAMENTOS</div><div class="card-lateral-valor">{total_contados_cnt}</div></div>', unsafe_allow_html=True)

        st.markdown("---")
        if st.button("🚀 Sincronizar Agora com a Nuvem", type="primary", use_container_width=True):
            if pendentes_ram > 0:
                qtd = sincronizar_ram_com_banco()
                st.success(f"✅ {qtd} bips enviados para o banco!")
            else:
                st.info("ℹ️ Nenhum bip pendente na memória.")
            st.rerun()

    # --- DECLARAÇÃO DAS ABAS PRINCIPAIS ---
    lista_abas = ["🔍 Contar Item (RAM Modo Rápido)", "📊 Lançamentos & Base", "📈 Desempenho & Acuracidade", "📁 Histórico Geral"]
    if eh_supervisor: lista_abas.append("⚙️ Gestão ADM")
    
    abas_objs = st.tabs(lista_abas)
    aba_contar, aba_lancamentos, aba_desempenho, aba_historico = abas_objs[0], abas_objs[1], abas_objs[2], abas_objs[3]
    aba_adm = abas_objs[4] if eh_supervisor else None

    # --- ABA 1: CONTAR ITEM (BIPAGEM INSTANTÂNEA EM RAM) ---
    with aba_contar:
        if pendentes_ram > 0:
            c_top1, c_top2 = st.columns([3, 1])
            c_top1.info(f"⚡ **Modo Ultra Rápido Ativo:** Você tem **{pendentes_ram} bips** guardados na RAM local.")
            with c_top2:
                df_ram_temp = pd.DataFrame(st.session_state.buffer_ram_contagens)
                st.download_button("📥 Backup RAM (.xlsx)", converter_para_excel(df_ram_temp), file_name=f"backup_ram_pasta_{id_pasta_limpo_base}.xlsx")

        if not id_inventario_atual or (base_sistema_atual is None and inventario_selected_obj['status'] == 'Aberto'):
            st.warning("⚠️ Selecione um inventário ativo e carregue a base na barra lateral.")
        elif inventario_selected_obj['status'] == "Aberto": st.warning("⚠️ Inventário em configuração. Libere a 1ª Contagem na barra lateral.")
        elif inventario_selected_obj['status'] == "Fechado": st.error("🔒 Inventário Fechado. Selecione um inventário ativo.")
        elif pode_fechar and inventario_selected_obj['status'] in ["1a Contagem", "2a Contagem"]:
            st.success("🎉 **100% dos itens desta fase já foram contados!** Você pode fechar o inventário na barra lateral.")
        else:
            codigo_input = st.text_input("💻 Bipar ou Digitar Código do Produto", value="", placeholder="Bipe a etiqueta aqui...", key=f"bip_{st.session_state.contador_reset}")
            if codigo_input:
                codigo_rastreio = str(codigo_input).upper().strip().split(" - ")[-1]
                matches_codigo = base_sistema_atual[base_sistema_atual['cod_produto'].astype(str).str.upper().str.strip() == codigo_rastreio]
                if matches_codigo.empty: st.error("❌ Código não cadastrado na planilha base!")
                else:
                    cursor = conn.cursor()
                    cursor.execute("SELECT lote, ativo, fase_contagem FROM contagens WHERE (inventario_id = %s OR inventario_id = %s) AND cod_produto = %s", (id_pasta_limpo_base, f"#{id_pasta_limpo_base}", codigo_rastreio))
                    set_ja_contados = {f"{str(r[0]).strip().upper() if r[0] else ''}_{str(r[1]).strip().upper() if r[1] else ''}" for r in cursor.fetchall()}
                    
                    for r_ram in st.session_state.buffer_ram_contagens:
                        if r_ram['cod_produto'] == codigo_rastreio:
                            set_ja_contados.add(f"{str(r_ram['lote']).strip().upper() if r_ram['lote'] else ''}_{str(r_ram['ativo']).strip().upper() if r_ram['ativo'] else ''}")

                    status_pasta_atual = inventario_selected_obj['status']

                    if status_pasta_atual == "1a Contagem":
                        matches_pendentes = [row_m for _, row_m in matches_codigo.iterrows() if f"{str(row_m['lote']).strip().upper() if pd.notna(row_m['lote']) and str(row_m['lote']).lower() != 'nan' else ''}_{str(row_m['ativo']).strip().upper() if pd.notna(row_m['ativo']) and str(row_m['ativo']).lower() != 'nan' else ''}" not in set_ja_contados]
                        if not matches_pendentes:
                            st.success(f"🎉 **Todos os ativos/lotes do produto {codigo_rastreio} já foram contabilizados!**")
                            pode_exibir_form = False
                        else:
                            matches_para_usar = pd.DataFrame(matches_pendentes)
                            pode_exibir_form = True
                    else:
                        matches_para_usar = matches_codigo
                        pode_exibir_form = True

                    if pode_exibir_form:
                        if len(matches_para_usar) > 1:
                            opcoes_item = [f"Ativo: {str(r['ativo']).strip() if pd.notna(r['ativo']) and str(r['ativo']).lower()!='nan' and str(r['ativo']).strip()!='' else 'Sem Ativo'} | Lote: {str(r['lote']).strip() if pd.notna(r['lote']) and str(r['lote']).lower()!='nan' and str(r['lote']).strip()!='' else 'Sem Lote'} (ID Linha #{r['id']})" for _, r in matches_para_usar.iterrows()]
                            item_sel_opcao = st.selectbox("Escolha o item específico PENDENTE:", opcoes_item)
                            item = matches_para_usar[matches_para_usar['id'] == int(item_sel_opcao.split("(ID Linha #")[1].replace(")", "").strip())].iloc[0]
                        else: item = matches_para_usar.iloc[0]

                        lote_auto = str(item['lote']).strip() if pd.notna(item['lote']) and str(item['lote']).lower() != 'nan' and str(item['lote']).strip() != '' else ''
                        ativo_auto = str(item['ativo']).strip() if pd.notna(item['ativo']) and str(item['ativo']).lower() != 'nan' and str(item['ativo']).strip() != '' else ''
                        id_est_limpo = str(item['id_estoque_fisico']).strip() if pd.notna(item['id_estoque_fisico']) and str(item['id_estoque_fisico']).strip() != '' else extrair_id_estoque_do_nome(inventario_selected_obj['nome'])
                        desc_est_limpo = str(item['desc_estoque_fisico']).strip() if pd.notna(item['desc_estoque_fisico']) and str(item['desc_estoque_fisico']).strip() != '' else MAPA_ESTOQUES_DESC.get(id_est_limpo, 'ESTOQUE FÍSICO')

                        c1, c2, c3 = st.columns(3)
                        c1.metric("Cód. Produto", str(item['cod_produto']))
                        c2.metric("Lote", lote_auto or "—")
                        c3.metric("Ativo", ativo_auto or "—")
                        
                        st.markdown(f'<div class="caixa-descricao-produto">📦 Descrição: {item["desc_produto"]}</div>', unsafe_allow_html=True)
                        st.info(f"📍 **Estoque Físico:** {desc_est_limpo} (ID: {id_est_limpo}) | **Unidade:** {item['unid_medida']} | **Status:** {status_pasta_atual}")
                        
                        with st.form("form_lancar_qtd", clear_on_submit=True):
                            qtd_fisica = st.number_input("📦 Quantidade Contada Fisicamente:", min_value=0, step=1, value=0)
                            confirma_zero = st.checkbox("⚠️ Marque se este item REALMENTE NÃO EXISTE no estoque (Saldo Zero)")
                            obs = st.text_input("Observação (opcional)")
                            
                            if st.form_submit_button("⚡ Salvar na RAM (Instantâneo)", type="primary", use_container_width=True):
                                if qtd_fisica == 0 and not confirma_zero: st.error("⚠️ Para salvar quantidade 0, marque a confirmação amarela!")
                                else:
                                    qtd_sys = int(item['qtd_estoque'])
                                    dif = qtd_fisica - qtd_sys
                                    fase_gravar = "2a Contagem Concluida" if status_pasta_atual == "2a Contagem" else status_pasta_atual
                                    data_hora_agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                                    st.session_state.buffer_ram_contagens.append({
                                        'inventario_id': id_pasta_limpo_base,
                                        'id_estoque': id_est_limpo,
                                        'desc_estoque': desc_est_limpo,
                                        'cod_produto': str(item['cod_produto']),
                                        'desc_produto': str(item['desc_produto']),
                                        'unid_medida': str(item['unid_medida']),
                                        'qtd_sistema': qtd_sys,
                                        'qtd_contada': qtd_fisica,
                                        'diferenca': dif,
                                        'ativo': ativo_auto,
                                        'observacao': obs,
                                        'operador': st.session_state.operador,
                                        'data_hora': data_hora_agora,
                                        'lote': lote_auto,
                                        'fase_contagem': fase_gravar
                                    })

                                    # Sincronização em lote automática a cada 10 bips
                                    if len(st.session_state.buffer_ram_contagens) >= 10:
                                        sincronizar_ram_com_banco()

                                    st.session_state.contador_reset += 1
                                    st.toast("⚡ Salvo na RAM instantaneamente!", icon="✅")
                                    st.rerun()

    # --- ABA 2: LANÇAMENTOS E ESPELHO BASE ---
    with aba_lancamentos:
        sub_aba1, sub_aba2 = st.tabs(["📋 Meus Lançamentos Nesta Pasta", "📄 Espelho Base do Saldo (Status Visual)"])
        with sub_aba1:
            df_minhas = pd.read_sql_query("SELECT * FROM contagens WHERE inventario_id = %s OR inventario_id = %s ORDER BY id DESC", conn, params=(id_pasta_limpo_base, f"#{id_pasta_limpo_base}"))
            
            if pendentes_ram > 0:
                df_ram = pd.DataFrame(st.session_state.buffer_ram_contagens)
                df_ram_pasta = df_ram[df_ram['inventario_id'] == id_pasta_limpo_base] if not df_ram.empty else pd.DataFrame()
                if not df_ram_pasta.empty:
                    df_minhas = pd.concat([df_ram_pasta, df_minhas], ignore_index=True)

            if not df_minhas.empty:
                m1, m2, m3 = st.columns(3)
                m1.metric("Lançamentos Totais", len(df_minhas))
                m2.metric("Com Divergência", len(df_minhas[df_minhas['diferenca'] != 0]))
                m3.metric("Total de Peças Contadas", int(df_minhas['qtd_contada'].sum()))
                st.download_button("📥 Exportar Lançamentos Filtrados para Excel", converter_para_excel(df_minhas), file_name=f"contagem_pasta_{id_pasta_limpo_base}.xlsx")
                st.dataframe(df_minhas[['cod_produto', 'desc_produto', 'desc_estoque', 'qtd_sistema', 'qtd_contada', 'diferenca', 'ativo', 'lote', 'observacao', 'operador', 'data_hora', 'fase_contagem']], use_container_width=True, hide_index=True)
            else: st.info("Nenhum lançamento registrado nesta pasta.")
            
        with sub_aba2:
            if base_sistema_atual is not None:
                cursor_l = conn.cursor()
                cursor_l.execute("SELECT cod_produto, ativo, lote, operador FROM contagens WHERE inventario_id = %s OR inventario_id = %s", (id_pasta_limpo_base, f"#{id_pasta_limpo_base}"))
                rows_l = cursor_l.fetchall()
                mapa_contados = {f"{str(r[0]).upper().strip()}_{str(r[2]).upper().strip() if r[2] and str(r[2]).lower()!='nan' else ''}_{str(r[1]).upper().strip() if r[1] and str(r[1]).lower()!='nan' else ''}": r[3] for r in rows_l} if rows_l else {}
                
                for r_r in st.session_state.buffer_ram_contagens:
                    if r_r['inventario_id'] == id_pasta_limpo_base:
                        mapa_contados[f"{str(r_r['cod_produto']).upper().strip()}_{str(r_r['lote']).upper().strip() if r_r['lote'] else ''}_{str(r_r['ativo']).upper().strip() if r_r['ativo'] else ''}"] = r_r['operador']

                def obter_status(row):
                    c, l, a = str(row['cod_produto']).upper().strip(), str(row['lote']).upper().strip() if pd.notna(row['lote']) and str(row['lote']).lower()!='nan' else "", str(row['ativo']).upper().strip() if pd.notna(row['ativo']) and str(row['ativo']).lower()!='nan' else ""
                    key = f"{c}_{l}_{a}"
                    if key in mapa_contados: return f"🟩 Contabilizado por ({mapa_contados[key]})"
                    elif f"{c}__" in mapa_contados: return f"🟩 Contabilizado por ({mapa_contados[f'{c}__']})"
                    return "🟥 Não Contado"
                
                df_espelho = base_sistema_atual.copy()
                df_espelho['Status de Contagem'] = df_espelho.apply(obter_status, axis=1)
                st.dataframe(df_espelho[['Status de Contagem', 'cod_produto', 'desc_produto', 'desc_estoque_fisico', 'unid_medida', 'qtd_estoque', 'id_estoque_fisico', 'lote', 'ativo']], use_container_width=True, hide_index=True)
            else: st.info("Nenhuma base carregada.")

    # --- ABA 3: DESEMPENHO E ACURACIDADE ---
    with aba_desempenho:
        sub_d1, sub_d2 = st.tabs(["🔴 Desempenho & Prazos por Estoque", "📈 Acuracidade Auditada pelo Supervisor"])
        with sub_d1:
            st.subheader("🏆 Status de Atualização dos Estoques Físicos")
            mapa_datas = {}
            df_ult_historico, df_ult_contagens = buscar_historico_estoques_cache()
            
            if not df_ult_historico.empty:
                for _, r_h in df_ult_historico.iterrows(): mapa_datas[str(r_h['id_estoque']).strip()] = str(r_h['ultima_data'])
            if not df_ult_contagens.empty:
                for _, r_u in df_ult_contagens.iterrows(): mapa_datas[str(r_u['id_estoque']).strip()] = str(r_u['ultima_data'])
            if not df_inventarios.empty:
                for _, r_p in df_inventarios.iterrows():
                    id_e = extrair_id_estoque_do_nome(r_p['nome'])
                    if id_e and id_e not in mapa_datas: mapa_datas[id_e] = str(r_p['data']) + " 12:00:00"

            linhas_desempenho, hoje = [], datetime.datetime.now()
            b_count, a_count, c_count = 0, 0, 0
            
            for est in LISTA_ESTOQUES_FIXA:
                est_id = str(est["id"]).strip()
                u_data = mapa_datas.get(est_id, None)
                if u_data:
                    try:
                        str_dt = str(u_data).split(".")[0]
                        if len(str_dt) == 10: str_dt += " 12:00:00"
                        dt = datetime.datetime.strptime(str_dt, "%Y-%m-%d %H:%M:%S")
                        dias, dt_fmt = (hoje - dt).days, dt.strftime("%d/%m/%Y %H:%M")
                    except Exception: dias, dt_fmt = 999, "Sem histórico"
                else: dias, dt_fmt = 999, "Nunca Contado"
                
                if dias <= 7: status, b_count = "🟢 Em Dia", b_count + 1
                elif dias <= 14: status, a_count = "🟡 Necessário Auditar", a_count + 1
                else: status, c_count = "🔴 Crítico (+2 semanas)", c_count + 1
                    
                linhas_desempenho.append({"Id. Estoque": est_id, "Descrição do Estoque Físico": est["desc"], "Última Contagem": dt_fmt, "Dias Sem Contar": dias if dias != 999 else 999, "Criticidade": status})

            df_desempenho_full = pd.DataFrame(linhas_desempenho)
            k1, k2, k3 = st.columns(3)
            k1.metric("🟢 Em Dia (até 7 dias)", b_count)
            k2.metric("🟡 Necessário Auditar (8 a 14 dias)", a_count)
            k3.metric("🔴 Crítico (+14 dias)", c_count)
            
            col_f_stat, col_f_busca = st.columns([1, 1])
            filtro_criticidade = col_f_stat.selectbox("🎯 Filtrar por Criticidade:", ["Todos os Estoques", "🟢 Em Dia", "🟡 Necessário Auditar", "🔴 Crítico (+2 semanas)"])
            busca_estoque = col_f_busca.text_input("🔍 Pesquisar Estoque por Código ou Nome:")

            df_exibir = df_desempenho_full.copy()
            if filtro_criticidade != "Todos os Estoques": df_exibir = df_exibir[df_exibir['Criticidade'] == filtro_criticidade]
            if busca_estoque.strip():
                termo = busca_estoque.strip().lower()
                df_exibir = df_exibir[df_exibir['Id. Estoque'].str.lower().str.contains(termo) | df_exibir['Descrição do Estoque Físico'].str.lower().str.contains(termo)]

            df_exibir['Dias Sem Contar'] = df_exibir['Dias Sem Contar'].apply(lambda x: "—" if x == 999 else x)
            st.dataframe(df_exibir.sort_values(by=['Criticidade', 'Id. Estoque'], ascending=[False, True]), use_container_width=True, hide_index=True)

        with sub_d2:
            st.subheader("📈 Tabela de Acuracidade Geral por Depósito")
            df_auds = pd.read_sql_query("SELECT * FROM auditorias_supervisor ORDER BY id DESC", conn)
            if df_auds.empty: st.info("💡 Nenhuma amostragem coletada pelo supervisor.")
            else:
                linhas_acu = []
                for dep_id, grp in df_auds.groupby('id_estoque'):
                    tot = len(grp)
                    p_s, p_e, p_l = (len(grp[grp['diferenca'] == 0])/tot)*100, (len(grp[grp['etiqueta_correta'] == "Sim"])/tot)*100, (len(grp[grp['localizacao_correta'] == "Sim"])/tot)*100
                    linhas_acu.append({"CÓDIGO ESTOQUE": dep_id, "DESCRIÇÃO DO ESTOQUE": grp.iloc[0]['desc_estoque'] if 'desc_estoque' in grp.columns else "Não Informado", "ACURACIDADE SALDO": f"{'🟢' if p_s==100 else '🔴'} {p_s:.1f}%", "ACURACIDADE ETIQUETAS": f"{'🟢' if p_e==100 else '🔴'} {p_e:.1f}%", "ACURACIDADE LOCALIZAÇÃO": f"{'🟢' if p_l==100 else '🔴'} {p_l:.1f}%", "ITENS AUDITADOS": tot, "ÚLTIMA AUDITORIA": grp.iloc[0]['data_hora'].split(" ")[0] if 'data_hora' in grp.columns else ""})
                st.dataframe(pd.DataFrame(linhas_acu), use_container_width=True, hide_index=True)

            st.markdown("---")
            if not df_inventarios_sup.empty:
                for _, inv_s in df_inventarios_sup.iterrows():
                    df_hist_sup = pd.read_sql_query("SELECT * FROM auditorias_supervisor WHERE inventario_id = %s ORDER BY id DESC", conn, params=(inv_s['id'],))
                    c_exp, c_del = st.columns([8, 2])
                    with c_exp:
                        with st.expander(f"📁 Pasta {inv_s['id']} – {inv_s['nome']} | Data: {inv_s['data']} | Status: {inv_s['status']} ({len(df_hist_sup)} itens auditados)"):
                            if not df_hist_sup.empty:
                                st.download_button("📥 Exportar Esta Auditoria para Excel", converter_para_excel(df_hist_sup), file_name=f"auditoria_{inv_s['id']}.xlsx", key=f"dl_sup_acu_{inv_s['id']}")
                                st.dataframe(df_hist_sup, use_container_width=True, hide_index=True)
                            else: st.info("Nenhum item auditado nesta pasta.")
                    with c_del:
                        if eh_supervisor and st.button("🗑️ Excluir Pasta", key=f"del_sup_f_{inv_s['id']}", use_container_width=True):
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM inventarios_supervisor WHERE id = %s", (inv_s['id'],))
                            cursor.execute("DELETE FROM auditorias_supervisor WHERE inventario_id = %s", (inv_s['id'],))
                            conn.commit()
                            limpar_cache_aplicacao()
                            st.rerun()

    # --- ABA 4: HISTÓRICO GERAL (COM PAGINAÇÃO DE 10 EM 10 E ORDENAÇÃO DECRESCENTE) ---
    with aba_historico:
        st.title("📁 Arquivo Geral de Movimentações")
        if df_inventarios.empty: 
            st.info("Nenhum inventário registrado.")
        else:
            df_inv_ordenados = df_inventarios.copy()
            df_inv_ordenados['id_num'] = df_inv_ordenados['id'].astype(str).str.replace('#', '', regex=False).astype(int)
            df_inv_ordenados = df_inv_ordenados.sort_values(by=['data', 'id_num'], ascending=[False, False])

            itens_por_pagina = 10
            total_itens = len(df_inv_ordenados)
            total_paginas = max(1, (total_itens + itens_por_pagina - 1) // itens_por_pagina)

            if st.session_state.pagina_historico > total_paginas:
                st.session_state.pagina_historico = total_paginas

            inicio_idx = (st.session_state.pagina_historico - 1) * itens_por_pagina
            fim_idx = inicio_idx + itens_por_pagina
            fatia_pastas = df_inv_ordenados.iloc[inicio_idx:fim_idx]

            for idx, inv in fatia_pastas.iterrows():
                id_proc = str(inv['id']).replace('#','').strip()
                df_h = pd.read_sql_query("SELECT * FROM contagens WHERE inventario_id = %s OR inventario_id = %s ORDER BY id DESC", conn, params=(id_proc, f"#{id_proc}"))
                tot_reg, acu_reg = len(df_h) if not df_h.empty else inv.get('total_itens', 0), inv.get('acuracidade_final', '—')

                c_exp, c_del = st.columns([8, 2])
                with c_exp:
                    with st.expander(f"📁 Pasta {inv['id']} - {inv['nome']} | Data: {inv['data']} | Status: {inv['status']} | Acuracidade: {acu_reg} ({tot_reg} itens contados)"):
                        if not df_h.empty:
                            st.download_button("📥 Baixar Planilha de Lançamentos (.xlsx)", converter_para_excel(df_h), file_name=f"inventario_{id_proc}.xlsx", key=f"dl_hist_{id_proc}")
                            st.dataframe(df_h[['id', 'inventario_id', 'id_estoque', 'desc_estoque', 'cod_produto', 'desc_produto', 'unid_medida', 'qtd_sistema', 'qtd_contada', 'diferenca', 'ativo', 'lote', 'observacao', 'operador', 'data_hora', 'fase_contagem']], use_container_width=True, hide_index=True)
                        else: st.info("Nenhum lançamento registrado nesta pasta.")
                with c_del:
                    if eh_supervisor and st.button("🗑️ Excluir Pasta", key=f"del_hist_inv_{inv['id']}", use_container_width=True):
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM inventarios WHERE id = %s OR id = %s", (inv['id'], id_proc))
                        cursor.execute("DELETE FROM contagens WHERE inventario_id = %s OR inventario_id = %s", (id_proc, id_proc))
                        cursor.execute("DELETE FROM itens_base_inventario WHERE inventario_id = %s OR inventario_id = %s", (id_proc, id_proc))
                        conn.commit()
                        limpar_cache_aplicacao()
                        st.rerun()

            st.markdown("---")
            c_pag1, c_pag2, c_pag3 = st.columns([2, 3, 2])
            with c_pag1:
                if st.button("◀ Anterior", disabled=(st.session_state.pagina_historico <= 1), use_container_width=True):
                    st.session_state.pagina_historico -= 1
                    st.rerun()
            with c_pag2:
                st.markdown(f"<h5 style='text-align: center; color: #666;'>Página {st.session_state.pagina_historico} de {total_paginas} ({total_itens} pastas no total)</h5>", unsafe_allow_html=True)
            with c_pag3:
                if st.button("Próxima ▶", disabled=(st.session_state.pagina_historico >= total_paginas), use_container_width=True):
                    st.session_state.pagina_historico += 1
                    st.rerun()

    # --- ABA 5: GESTÃO ADM ---
    if eh_supervisor and aba_adm is not None:
        with aba_adm:
            st.title("⚙️ Módulo de Gestão do Administrador")
            opcao_adm = st.selectbox("Escolha o Módulo de Ação:", ["🚨 Liberar / Encerrar Divergências", "🔬 Auditoria Amostral (Supervisor)", "📊 Relatório Consolidado (Excel Gerencial)", "👥 Gestão de Usuários & Senhas"])
            st.markdown("---")

            if opcao_adm == "🚨 Liberar / Encerrar Divergências":
                st.subheader("Tratamento de Erros de Contagem da Equipe")
                
                df_invs_div = pd.read_sql_query("""
                    SELECT DISTINCT c.inventario_id, i.nome, i.data 
                    FROM contagens c 
                    LEFT JOIN inventarios i ON (REPLACE(i.id, '#', '') = REPLACE(c.inventario_id, '#', '')) 
                    WHERE c.diferenca != 0 
                      AND (c.fase_contagem IS NULL OR c.fase_contagem NOT IN ('2a Contagem', 'Encerrado com Divergencia'))
                """, conn)
                
                if df_invs_div.empty:
                    st.success("🎉 Nenhuma divergência pendente de tratamento no momento!")
                else:
                    mapa_invs_div = {}
                    opcoes_invs_div = []
                    for _, r_inv in df_invs_div.iterrows():
                        id_p_limpo = str(r_inv['inventario_id']).replace('#', '').strip()
                        nome_p = r_inv['nome'] if pd.notna(r_inv['nome']) else f"Pasta #{id_p_limpo}"
                        label_inv = f"Pasta #{id_p_limpo} – {nome_p}"
                        opcoes_invs_div.append(label_inv)
                        mapa_invs_div[label_inv] = id_p_limpo
                    
                    inv_alvo_sel = st.selectbox("📁 Selecione o Inventário/Pasta Com Divergência:", opcoes_invs_div)
                    id_pasta_div_target = mapa_invs_div[inv_alvo_sel]
                    
                    df_items_div = pd.read_sql_query("""
                        SELECT id, cod_produto, desc_produto, lote, ativo, qtd_sistema, qtd_contada, diferenca, operador 
                        FROM contagens 
                        WHERE diferenca != 0 
                          AND (fase_contagem IS NULL OR fase_contagem NOT IN ('2a Contagem', 'Encerrado com Divergencia'))
                          AND REPLACE(inventario_id, '#', '') = %s
                    """, conn, params=(id_pasta_div_target,))
                    
                    if df_items_div.empty:
                        st.info("Nenhuma divergência pendente nesta pasta.")
                    else:
                        mapa_items = {}
                        opcoes_items = []
                        for _, r_it in df_items_div.iterrows():
                            l_str = f" | Lote: {r_it['lote']}" if pd.notna(r_it['lote']) and str(r_it['lote']).strip() != '' else ""
                            a_str = f" | Ativo: {r_it['ativo']}" if pd.notna(r_it['ativo']) and str(r_it['ativo']).strip() != '' else ""
                            label_item = f"{r_it['cod_produto']} - {r_it['desc_produto']}{l_str}{a_str} (Dif: {r_it['diferenca']})"
                            opcoes_items.append(label_item)
                            mapa_items[label_item] = r_it

                        item_alvo_sel = st.selectbox("📦 Selecione o Item Com Divergência Nesta Pasta:", opcoes_items)
                        row_item_alvo = mapa_items[item_alvo_sel]
                        cod_target = str(row_item_alvo['cod_produto']).strip()
                        
                        st.info(f"📊 **Resumo do Item:** Sistema: **{row_item_alvo['qtd_sistema']}** | Contado: **{row_item_alvo['qtd_contada']}** | Diferença: **{row_item_alvo['diferenca']}** (Lançado por: {row_item_alvo['operador']})")
                        justificativa_adm = st.text_input("📝 Informe a justificativa/observação:", value="Divergente", placeholder="Ex: Liberando para 2a contagem...")
                        
                        col_act1, col_act2 = st.columns(2)
                        with col_act1:
                            if st.button("🚨 Abrir 2ª Contagem para Almoxarife", type="primary", use_container_width=True):
                                if not justificativa_adm.strip():
                                    st.error("⚠️ Digite uma justificativa antes de liberar!")
                                else:
                                    obs_final = f"ADM ({st.session_state.operador}): [LIBERADO 2ª CONTAGEM] - {justificativa_adm.strip()}"
                                    cursor = conn.cursor()
                                    cursor.execute("UPDATE contagens SET fase_contagem = '2a Contagem', observacao = %s WHERE id = %s", (obs_final, row_item_alvo['id']))
                                    cursor.execute("UPDATE inventarios SET status = '2a Contagem' WHERE REPLACE(id, '#', '') = %s", (id_pasta_div_target,))
                                    conn.commit()
                                    limpar_cache_aplicacao()
                                    st.success(f"✅ Material {cod_target} enviado para 2ª Contagem!")
                                    time.sleep(1)
                                    st.rerun()

                        with col_act2:
                            if st.button("🔒 Finalizar e Manter Divergência Atual", use_container_width=True):
                                if not justificativa_adm.strip():
                                    st.error("⚠️ Digite uma justificativa antes de encerrar!")
                                else:
                                    obs_final = f"ADM ({st.session_state.operador}): [ENCERRADO COM DIVERGÊNCIA] - {justificativa_adm.strip()}"
                                    cursor = conn.cursor()
                                    cursor.execute("UPDATE contagens SET fase_contagem = 'Encerrado com Divergencia', observacao = %s WHERE id = %s", (obs_final, row_item_alvo['id']))
                                    cursor.execute("UPDATE inventarios SET status = 'Fechado' WHERE REPLACE(id, '#', '') = %s", (id_pasta_div_target,))
                                    conn.commit()
                                    limpar_cache_aplicacao()
                                    st.success(f"✅ Item {cod_target} resolvido e Pasta #{id_pasta_div_target} encerrada!")
                                    time.sleep(1)
                                    st.rerun()

            elif opcao_adm == "🔬 Auditoria Amostral (Supervisor)":
                st.subheader("Módulo de Auditoria Amostral Própria")
                if df_inventarios_sup.empty: id_sup_act, inv_sup_obj = None, None
                else:
                    sel_s = st.selectbox("Selecione a Pasta de Auditoria Ativa:", [f"{r['id']} – {r['nome']} ({r['status']})" for _, r in df_inventarios_sup.iterrows()], key="sb_sup_active")
                    id_sup_act = sel_s.split(" – ")[0]
                    inv_sup_obj = df_inventarios_sup[df_inventarios_sup['id'] == id_sup_act].iloc[0]

                if inv_sup_obj is not None and inv_sup_obj['status'] == "Aberto":
                    if st.button("🔒 Fechar Esta Pasta de Auditoria", type="primary"):
                        cursor = conn.cursor()
                        cursor.execute("UPDATE inventarios_supervisor SET status = 'Fechado' WHERE id = %s", (id_sup_act,))
                        conn.commit()
                        limpar_cache_aplicacao()
                        st.rerun()

                with st.expander("➕ Nova Pasta de Auditoria do Supervisor"):
                    with st.form("form_sup_new"):
                        nom_s = st.text_input("Nome da Pasta Amostral")
                        if st.form_submit_button("Criar Pasta Supervisor", type="primary") and nom_s:
                            cursor = conn.cursor()
                            df_c_s = pd.read_sql_query("SELECT id FROM inventarios_supervisor", conn)
                            m_id_s = df_c_s['id'].str.replace('SUP-#', '', regex=False).astype(int).max() if not df_c_s.empty else 0
                            cursor.execute("INSERT INTO inventarios_supervisor (id, nome, data, status) VALUES (%s, %s, %s, 'Aberto')", (f"SUP-#{m_id_s + 1}", nom_s, datetime.date.today().strftime("%Y-%m-%d")))
                            conn.commit()
                            limpar_cache_aplicacao()
                            st.rerun()

                arq_sup = st.file_uploader("Suba a planilha Excel de amostras (.xlsx)", type=["xlsx"], key="up_excel_sup")
                if arq_sup is not None and id_sup_act:
                    try:
                        st.session_state.bases_supervisor_por_inv[id_sup_act] = pd.read_excel(arq_sup)
                        st.success("✅ Planilha anexa com sucesso!")
                    except Exception as e: st.error(f"Erro: {e}")

                base_sup_curr = st.session_state.bases_supervisor_por_inv.get(id_sup_act, None)

                if id_sup_act and base_sup_curr is not None and inv_sup_obj['status'] == "Aberto":
                    cols_sup = list(base_sup_curr.columns)
                    def mapear_col_s(opcoes, idx_padrao):
                        for op in opcoes:
                            for c in cols_sup:
                                if op.lower().replace(" ", "").replace(".", "") in str(c).lower().replace(" ", "").replace(".", ""): return c
                        return cols_sup[idx_padrao] if idx_padrao < len(cols_sup) else cols_sup[0]

                    col_cod_s, col_desc_s = mapear_col_s(['códproduto', 'codproduto', 'codigo'], 0), mapear_col_s(['descproduto', 'descricao'], 1)
                    col_local_s, col_qtd_s, col_id_est_s = mapear_col_s(['descestoquefisico', 'localizacao'], 2), mapear_col_s(['qtdestoque', 'quantidade'], -1), mapear_col_s(['idestoquefísico', 'idestoque'], 0)

                    df_ja_auditados = pd.read_sql_query("SELECT cod_produto FROM auditorias_supervisor WHERE inventario_id = %s", conn, params=(id_sup_act,))
                    cods_ja_auditados = set(df_ja_auditados['cod_produto'].astype(str).str.upper().str.strip().tolist()) if not df_ja_auditados.empty else set()

                    tot_amostra = len(base_sup_curr)
                    st.progress(min(1.0, len(cods_ja_auditados) / tot_amostra) if tot_amostra > 0 else 0.0)
                    st.dataframe(base_sup_curr, use_container_width=True, hide_index=True)

                    base_pendente = base_sup_curr[~base_sup_curr[col_cod_s].astype(str).str.upper().str.strip().isin(cods_ja_auditados)]

                    col_bip, col_select = st.columns([1, 1])
                    bip_sup = col_bip.text_input("Bipe o código com leitor (opcional):", key=f"bip_sup_{st.session_state.contador_reset_sup}")
                    item_combo_sup = col_select.selectbox("Ou selecione o material da lista:", [f"{r[col_cod_s]} - {r[col_desc_s]}" for _, r in base_pendente.iterrows()] if not base_pendente.empty else [], key="combo_sup_amostra") if not base_pendente.empty else None

                    cod_sup_clean = str(bip_sup).upper().strip().split(" - ")[-1].strip() if bip_sup else (str(item_combo_sup).split(" - ")[0].strip() if item_combo_sup else "")

                    if cod_sup_clean:
                        match_s = base_sup_curr[base_sup_curr[col_cod_s].astype(str).str.upper().str.strip() == cod_sup_clean]
                        if not match_s.empty:
                            row_s = match_s.iloc[0]
                            st.info(f"📦 **Auditando Item:** {cod_sup_clean} - {row_s[col_desc_s]}")
                            with st.form("form_auditar_item_completo", clear_on_submit=True):
                                c_f1, c_f2, c_f3 = st.columns(3)
                                q_aud = c_f1.number_input("Quantidade Real Encontrada Fisicamente:", min_value=0, step=1, value=0)
                                e_ok = c_f2.selectbox("A Etiqueta Física está Correta?", ["Sim", "Não"])
                                l_ok = c_f3.selectbox("O Endereçamento/Localização está Correto?", ["Sim", "Não"])
                                at_sup = st.text_input("Número do Ativo (Opcional)")
                                if st.form_submit_button("💾 Salvar Auditoria do Item", type="primary", use_container_width=True):
                                    cursor = conn.cursor()
                                    cursor.execute("INSERT INTO auditorias_supervisor (inventario_id, id_estoque, desc_estoque, cod_produto, desc_produto, qtd_sistema, qtd_auditada, diferenca, etiqueta_correta, localizacao_correta, supervisor, data_hora, ativo) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                                                   (id_sup_act, str(row_s[col_id_est_s]).strip() if col_id_est_s in base_sup_curr.columns else "", str(row_s[col_local_s]) if col_local_s in base_sup_curr.columns else "Não Informado", cod_sup_clean, str(row_s[col_desc_s]), int(pd.to_numeric(row_s[col_qtd_s], errors='coerce') or 0), q_aud, q_aud - int(pd.to_numeric(row_s[col_qtd_s], errors='coerce') or 0), e_ok, l_ok, st.session_state.operador, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), at_sup.strip().upper()))
                                    conn.commit()
                                    limpar_cache_aplicacao()
                                    st.success("✅ Auditoria registrada!")
                                    st.session_state.contador_reset_sup += 1
                                    st.rerun()

            elif opcao_adm == "📊 Relatório Consolidado (Excel Gerencial)":
                st.subheader("Gerar Planilha Gerencial por Período (Múltiplas Abas)")
                
                hoje = datetime.date.today()
                quinze_dias_atras = hoje - datetime.timedelta(days=15)
                
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    data_inicio = st.date_input("📅 Data Inicial:", value=quinze_dias_atras, format="DD/MM/YYYY")
                with col_d2:
                    data_fim = st.date_input("📅 Data Final:", value=hoje, format="DD/MM/YYYY")
                
                str_ini = data_inicio.strftime("%Y-%m-%d")
                str_fim = data_fim.strftime("%Y-%m-%d")
                
                df_pastas_periodo = pd.read_sql_query("""
                    SELECT id, nome, data 
                    FROM inventarios 
                    WHERE data >= %s AND data <= %s 
                    ORDER BY data DESC, id DESC
                """, conn, params=(str_ini, str_fim))
                
                if df_pastas_periodo.empty:
                    st.info(f"ℹ️ Nenhum inventário registrado entre **{data_inicio.strftime('%d/%m/%Y')}** e **{data_fim.strftime('%d/%m/%Y')}**.")
                else:
                    st.success(f"📌 Foram encontrados **{len(df_pastas_periodo)}** inventários no período de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}.")
                    
                    ids_pastas = [str(r['id']).replace('#', '').strip() for _, r in df_pastas_periodo.iterrows()]
                    
                    if st.button("🚀 Gerar Excel Consolidado do Período", type="primary", use_container_width=True):
                        if ids_pastas:
                            ph = ', '.join(['%s'] * len(ids_pastas))
                            df_c_fil = pd.read_sql_query(f"""
                                SELECT * FROM contagens 
                                WHERE inventario_id IN ({ph}) OR inventario_id IN ({', '.join(['%s'] * len(ids_pastas))})
                            """, conn, params=ids_pastas + [f"#{x}" for x in ids_pastas])
                            
                            if df_c_fil.empty:
                                st.warning("⚠️ Os inventários do período foram criados, mas ainda não possuem nenhum lançamento de contagem registrado.")
                            else:
                                bytes_ex = gerar_relatorio_consolidado_excel(df_c_fil, LISTA_ESTOQUES_FIXA)
                                st.download_button(
                                    label=f"📥 Clique para Baixar o Relatório ({data_inicio.strftime('%d-%m-%Y')} a {data_fim.strftime('%d-%m-%Y')}).xlsx", 
                                    data=bytes_ex, 
                                    file_name=f"Relatorio_Gerencial_{str_ini}_a_{str_fim}.xlsx", 
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                    use_container_width=True
                                )

            elif opcao_adm == "👥 Gestão de Usuários & Senhas":
                st.subheader("Gerenciamento de Colaboradores e Perfis")
                df_usrs = pd.read_sql_query("SELECT id, nome, cpf, email, perfil FROM usuarios ORDER BY nome", conn)
                st.dataframe(df_usrs, use_container_width=True, hide_index=True)
                c_u1, c_u2 = st.columns(2)
                with c_u1:
                    u_sel = st.selectbox("Escolha o Colaborador:", [f"{r['id']} - {r['nome']}" for _, r in df_usrs.iterrows()])
                    n_senha, n_perfil = st.text_input("Nova Senha", type="password"), st.selectbox("Nível de Acesso:", ["Almoxarife", "Administrador"])
                    if st.button("🔄 Atualizar Dados do Usuário", type="primary", use_container_width=True):
                        cursor = conn.cursor()
                        if n_senha.strip(): cursor.execute("UPDATE usuarios SET senha = %s, perfil = %s WHERE id = %s", (n_senha.strip(), n_perfil, u_sel.split(" - ")[0]))
                        else: cursor.execute("UPDATE usuarios SET perfil = %s WHERE id = %s", (n_perfil, u_sel.split(" - ")[0]))
                        conn.commit()
                        limpar_cache_aplicacao()
                        st.success("✅ Atualizado!")
                        st.rerun()
                with c_u2:
                    u_del = st.selectbox("Remover Colaborador:", [f"{r['id']} - {r['nome']}" for _, r in df_usrs.iterrows()], key="sb_del")
                    if st.button("❌ Confirmar Exclusão", type="primary", use_container_width=True):
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM usuarios WHERE id = %s", (u_del.split(" - ")[0],))
                        conn.commit()
                        limpar_cache_aplicacao()
                        st.success("✅ Usuário removido!")
                        st.rerun()

    conn.close()
