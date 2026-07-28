import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from PIL import Image

# ====================== CONFIGURAÇÕES GERAIS ======================
ARQUIVO = "dados_funcionarios.xlsx"
PASTA_DOCS = "Documentos_Lojas"
PASTA_DOCS_FUNC = "Documentos_Funcionarios"
PASTA_FOTOS = "Fotos_Funcionarios"
PASTA_COMPROVANTES = "Comprovantes_Viagens"
os.makedirs(PASTA_DOCS, exist_ok=True)
os.makedirs(PASTA_DOCS_FUNC, exist_ok=True)
os.makedirs(PASTA_FOTOS, exist_ok=True)
os.makedirs(PASTA_COMPROVANTES, exist_ok=True)

MESES = ["Todos", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MAP_MES = {
    "Jan":1, "Fev":2, "Mar":3, "Abr":4, "Mai":5, "Jun":6,
    "Jul":7, "Ago":8, "Set":9, "Out":10, "Nov":11, "Dez":12
}

SITUACOES = [
    "Ativo", "Pré-cadastro", "Abandono", "Término de Contrato",
    "Demitido S/JC", "Demitido C/JC", "Pedido de Conta",
    "Rescisão Indireta", "Férias", "Doença", "Acidente", "Maternidade"
]

STATUS_VIAGEM = ["Pendente", "Em Andamento", "Concluída", "Cancelada"]
TIPOS_DESPESA = ["Alimentação", "Transporte", "Hospedagem", "Combustível", "Pedágio", "Estacionamento", "Outros"]

# ====================== BANCO DE DADOS ======================
@st.cache_data(ttl=0, show_spinner=False)
def carregar_dados():
    try:
        dados = pd.read_excel(ARQUIVO, sheet_name=None, dtype=str, keep_default_na=False)
    except:
        dados = {}
    
    padrao = {
        "Base_Dados": [
            "Matricula","Nome","CPF","RG","PIS","Nascimento","Admissao",
            "Telefone","Endereco","Loja","Cargo","Salario","Situacao",
            "DataAvisoPrevio","DiasAvisoPrevio","DataTerminoAviso",
            "DataFeriasInicio","DiasFerias","DataRetornoFerias",
            "DataPedidoConta","DataRescisao","DataAbandono",
            "DataTerminoContrato",
            "DataLicenca","DiasLicenca","DataTerminoLicenca",
            "DataAfastamento","DiasAfastamento","DataRetornoAfastamento",
            "CaminhoFoto"
        ],
        "Historico": [
            "DataEvento","TipoEvento","Matricula","Nome","CPF","RG","PIS",
            "Nascimento","Admissao","Telefone","Endereco","Loja","Cargo",
            "Salario","Situacao","DataAvisoPrevio","DiasAvisoPrevio","DataTerminoAviso",
            "DataFeriasInicio","DiasFerias","DataRetornoFerias",
            "DataPedidoConta","DataRescisao","DataAbandono",
            "DataTerminoContrato",
            "DataLicenca","DiasLicenca","DataTerminoLicenca",
            "DataAfastamento","DiasAfastamento","DataRetornoAfastamento","Detalhes"
        ],
        "Auxiliares": ["Loja", "Cargo"],
        "Docs_Lojas": ["Loja","Mes","Ano","NomeArquivo","Caminho","DataAnexado","Responsavel"],
        "Docs_Funcionarios": ["Matricula","Nome","TipoDoc","NomeArquivo","Caminho","DataAnexado"],
        "Viagens": [
            "ID","Matricula","Nome","Destino","DataInicio","DataFim",
            "ValorRecebido","ValorGasto","ValorRestante","Status","Observacoes","DataCadastro"
        ],
        "PrestacaoContas": [
            "ID","IDViagem","DataDespesa","TipoDespesa","Descricao",
            "Valor","CaminhoComprovante","DataRegistro"
        ]
    }
    
    for aba, cols in padrao.items():
        if aba not in dados:
            dados[aba] = pd.DataFrame(columns=cols)
        else:
            for c in cols:
                if c not in dados[aba].columns:
                    dados[aba][c] = ""
            if "Matricula" in dados[aba].columns:
                dados[aba]["Matricula"] = dados[aba]["Matricula"].astype(str).str.strip()
            if "Situacao" in dados[aba].columns:
                dados[aba]["Situacao"] = dados[aba]["Situacao"].astype(str).str.strip()
            if "ID" in dados[aba].columns:
                dados[aba]["ID"] = dados[aba]["ID"].astype(str).str.strip()
            if "IDViagem" in dados[aba].columns:
                dados[aba]["IDViagem"] = dados[aba]["IDViagem"].astype(str).str.strip()
    return dados

def salvar_dados(dados):
    try:
        with pd.ExcelWriter(ARQUIVO, engine="openpyxl", mode="w") as f:
            for aba, df in dados.items():
                df.to_excel(f, sheet_name=aba, index=False)
        st.cache_data.clear()
    except PermissionError:
        st.error("❌ ERRO: O arquivo dados_funcionarios.xlsx está ABERTO! Feche o Excel e tente novamente.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {str(e)}")
        st.stop()

def lista_lojas():
    d = carregar_dados()
    todas = sorted(set(
        [str(l).strip() for l in d["Base_Dados"]["Loja"] if str(l).strip() != ""] +
        [str(l).strip() for l in d["Auxiliares"]["Loja"] if str(l).strip() != ""]
    ))
    return todas if todas else ["Sem Loja"]

def lista_cargos():
    d = carregar_dados()
    todas = sorted(set(
        [str(c).strip() for c in d["Base_Dados"]["Cargo"] if str(c).strip() != ""] +
        [str(c).strip() for c in d["Auxiliares"]["Cargo"] if str(c).strip() != ""]
    ))
    return todas if todas else ["Sem Cargo"]

def calcular_e_atualizar(form):
    if form.get("dt_aviso") and form.get("dias_aviso") and str(form["dias_aviso"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_aviso"], "%d/%m/%Y")
            form["termino_aviso"] = (dt + timedelta(days=int(form["dias_aviso"]))).strftime("%d/%m/%Y")
        except: form["termino_aviso"] = ""
    else: form["termino_aviso"] = ""

    if form.get("dt_lic") and form.get("dias_lic") and str(form["dias_lic"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_lic"], "%d/%m/%Y")
            form["termino_lic"] = (dt + timedelta(days=int(form["dias_lic"]))).strftime("%d/%m/%Y")
        except: form["termino_lic"] = ""
    else: form["termino_lic"] = ""

    if form.get("dt_fer") and form.get("dias_fer") and str(form["dias_fer"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_fer"], "%d/%m/%Y")
            form["retorno_fer"] = (dt + timedelta(days=int(form["dias_fer"]))).strftime("%d/%m/%Y")
            if not any([
                form.get("dt_pedido","").strip(), form.get("dt_rescisao","").strip(),
                form.get("dt_abandono","").strip(), form.get("dt_termino_cont","").strip()
            ]):
                form["situacao"] = "Férias"
        except:
            form["retorno_fer"] = ""
    else:
        form["retorno_fer"] = ""

    if form.get("dt_af") and form.get("dias_af") and str(form["dias_af"]).isdigit():
        try:
            dt = datetime.strptime(form["dt_af"], "%d/%m/%Y")
            form["retorno_af"] = (dt + timedelta(days=int(form["dias_af"]))).strftime("%d/%m/%Y")
        except: form["retorno_af"] = ""
    else: form["retorno_af"] = ""

    if form.get("dt_termino_cont") and form.get("dt_termino_cont").strip():
        form["situacao"] = "Término de Contrato"
    elif form.get("dt_pedido") and form.get("dt_pedido").strip():
        form["situacao"] = "Pedido de Conta"
    elif form.get("dt_rescisao") and form.get("dt_rescisao").strip():
        form["situacao"] = "Rescisão Indireta"
    elif form.get("dt_abandono") and form.get("dt_abandono").strip():
        form["situacao"] = "Abandono"

    return form

def add_historico_auto(mat, nome, acao, dados_completos):
    dados = carregar_dados()
    registro = {"DataEvento": datetime.now().strftime("%d/%m/%Y"), "TipoEvento": acao, "Detalhes": ""}
    registro.update(dados_completos)
    idx = dados["Historico"].index[dados["Historico"]["Matricula"] == mat].tolist()
    if idx: dados["Historico"].iloc[idx[0]] = registro
    else: dados["Historico"] = pd.concat([dados["Historico"], pd.DataFrame([registro])], ignore_index=True)
    salvar_dados(dados)

def gerar_id_viagem():
    dados = carregar_dados()
    if dados["Viagens"].empty:
        return "V001"
    try:
        ultimos = [int(str(x).replace("V","")) for x in dados["Viagens"]["ID"] if str(x).startswith("V") and str(x).replace("V","").isdigit()]
        return f"V{max(ultimos)+1:03d}"
    except:
        return f"V{len(dados['Viagens'])+1:03d}"

def atualizar_totais_viagem(id_viagem):
    dados = carregar_dados()
    despesas = dados["PrestacaoContas"][dados["PrestacaoContas"]["IDViagem"] == id_viagem]
    total_gasto = 0.0
    for _, row in despesas.iterrows():
        try:
            v = float(str(row["Valor"]).replace(",",".").replace("R$","").strip())
            total_gasto += v
        except:
            pass
    idx = dados["Viagens"].index[dados["Viagens"]["ID"] == id_viagem].tolist()
    if idx:
        try:
            recebido = float(str(dados["Viagens"].iloc[idx[0]]["ValorRecebido"]).replace(",",".").replace("R$","").strip())
        except:
            recebido = 0.0
        restante = recebido - total_gasto
        dados["Viagens"].at[idx[0], "ValorGasto"] = f"{total_gasto:.2f}"
        dados["Viagens"].at[idx[0], "ValorRestante"] = f"{restante:.2f}"
        salvar_dados(dados)

# ====================== INTERFACE PRINCIPAL ======================
st.set_page_config(page_title="SISTEMA RH COMPLETO", layout="wide", initial_sidebar_state="collapsed")
st.title("📋 SISTEMA RH COMPLETO")

aba1, aba2, aba3, aba4, aba5, aba6, aba7, aba8 = st.tabs([
    "Cadastro", "Painel", "Prazos e Férias", "Histórico", "Relatórios", "📎 Documentos", "🚗 Guia de Viagem", "⚙️ Lojas e Cargos"
])

# ================ ABA 1 - CADASTRO ================
with aba1:
    dados = carregar_dados()
    
    busca = st.text_input("🔍 Buscar por Matrícula ou Nome", placeholder="Digite exatamente como está na planilha")
    
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        filtro_loja = st.selectbox("Filtrar por Loja", ["Todas"] + lista_lojas())
    with col_f2:
        filtro_sit = st.selectbox("Filtrar por Situação", ["Todas"] + SITUACOES)
    with col_f3:
        filtro_cargo = st.selectbox("Filtrar por Cargo", ["Todos"] + lista_cargos())

    lista = dados["Base_Dados"].copy()
    lista["Matricula"] = lista["Matricula"].fillna("").astype(str).str.strip()
    lista["Nome"] = lista["Nome"].fillna("").astype(str).str.strip()

    if busca.strip():
        lista = lista[
            (lista["Matricula"].str.contains(busca, case=False, na=False)) |
            (lista["Nome"].str.contains(busca, case=False, na=False))
        ]
    if filtro_loja != "Todas":
        lista = lista[lista["Loja"] == filtro_loja]
    if filtro_sit != "Todas":
        lista = lista[lista["Situacao"] == filtro_sit]
    if filtro_cargo != "Todos":
        lista = lista[lista["Cargo"] == filtro_cargo]

    st.dataframe(
        lista[["Matricula","Nome","Loja","Situacao","Cargo"]],
        use_container_width=True, hide_index=True
    )

    mat_sel = st.text_input("✏️ Digite a Matrícula EXATA para editar / excluir", placeholder="Igual coluna A da planilha")
    reg = pd.DataFrame()
    if mat_sel.strip():
        mat_busca = str(mat_sel).strip()
        reg = dados["Base_Dados"][dados["Base_Dados"]["Matricula"] == mat_busca]

    val_campo = lambda nome: reg.iloc[0][nome] if not reg.empty else ""

    prazos_exp = []
    if not reg.empty and val_campo("Admissao").strip():
        try:
            dt_adm = datetime.strptime(val_campo("Admissao"), "%d/%m/%Y")
            hoje = datetime.now()
            dias_corridos = (hoje - dt_adm).days
            for prazo in [29, 44, 59, 89]:
                rest = prazo - dias_corridos
                if rest > 0:
                    status = f"Faltam {rest} dias"
                elif rest == 0:
                    status = "HOJE"
                else:
                    status = f"Vencido há {abs(rest)} dias"
                prazos_exp.append([f"{prazo} dias", (dt_adm + timedelta(days=prazo)).strftime("%d/%m/%Y"), status])
        except:
            pass

    if not reg.empty:
        temp = {
            "dt_aviso": val_campo("DataAvisoPrevio"), "dias_aviso": val_campo("DiasAvisoPrevio"),
            "dt_lic": val_campo("DataLicenca"), "dias_lic": val_campo("DiasLicenca"),
            "dt_fer": val_campo("DataFeriasInicio"), "dias_fer": val_campo("DiasFerias"),
            "dt_af": val_campo("DataAfastamento"), "dias_af": val_campo("DiasAfastamento"),
            "dt_pedido": val_campo("DataPedidoConta"), "dt_rescisao": val_campo("DataRescisao"),
            "dt_abandono": val_campo("DataAbandono"), "dt_termino_cont": val_campo("DataTerminoContrato"),
            "situacao": val_campo("Situacao"), "caminho_foto": val_campo("CaminhoFoto")
        }
        temp = calcular_e_atualizar(temp)
        term_aviso_val, term_lic_val, ret_fer_val, ret_af_val, situacao_val, caminho_foto_atual = temp["termino_aviso"], temp["termino_lic"], temp["retorno_fer"], temp["retorno_af"], temp["situacao"], temp["caminho_foto"]
    else:
        term_aviso_val = term_lic_val = ret_fer_val = ret_af_val = caminho_foto_atual = ""
        situacao_val = "Ativo"

    # Botão de limpar FORA do formulário
    if st.button("🗑️ LIMPAR TODOS OS CAMPOS", use_container_width=True, type="secondary"):
        st.rerun()

    with st.form("form_cadastro", clear_on_submit=True):
        st.subheader("Dados Básicos")
        col_foto, col_dados = st.columns([1,3])
        
        with col_foto:
            st.markdown("**Foto do Funcionário**")
            if caminho_foto_atual and os.path.exists(caminho_foto_atual):
                st.image(caminho_foto_atual, width=180, caption="Foto atual")
            else:
                st.info("Sem foto")
            
            nova_foto = st.file_uploader("Enviar/Trocar foto", type=["jpg","jpeg","png"], key=f"foto_{mat_sel}")
            excluir_foto = st.checkbox("🗑️ Excluir foto atual", value=False)

        with col_dados:
            c1,c2,c3 = st.columns(3)
            with c1:
                matricula = st.text_input("Matrícula * (igual planilha)", value=val_campo("Matricula"))
                nome = st.text_input("Nome Completo", value=val_campo("Nome"))
                cpf = st.text_input("CPF", value=val_campo("CPF"))
                rg = st.text_input("RG", value=val_campo("RG"))
                pis = st.text_input("PIS", value=val_campo("PIS"))
            with c2:
                nascimento = st.text_input("Data Nascimento (dd/mm/aaaa)", value=val_campo("Nascimento"))
                admissao = st.text_input("Data Admissão (dd/mm/aaaa)", value=val_campo("Admissao"))
                telefone = st.text_input("Telefone", value=val_campo("Telefone"))
                endereco = st.text_input("Endereço Completo", value=val_campo("Endereco"))
            with c3:
                lojas = lista_lojas()
                idx_loja = lojas.index(val_campo("Loja")) if val_campo("Loja") in lojas else 0
                loja = st.selectbox("🏬 Loja", lojas, index=idx_loja)

                cargos = lista_cargos()
                idx_cargo = cargos.index(val_campo("Cargo")) if val_campo("Cargo") in cargos else 0
                cargo = st.selectbox("💼 Cargo", cargos, index=idx_cargo)

                salario = st.text_input("Salário", value=val_campo("Salario"))

                idx_sit = SITUACOES.index(situacao_val) if situacao_val in SITUACOES else 0
                situacao = st.selectbox("📊 Situação", SITUACOES, index=idx_sit)

        if prazos_exp:
            st.markdown("---")
            st.subheader("⏳ PRAZOS DE EXPERIÊNCIA")
            st.dataframe(
                pd.DataFrame(prazos_exp, columns=["Prazo", "Data Final", "Situação"]),
                use_container_width=True, hide_index=True
            )
        elif not reg.empty:
            st.info("ℹ️ Informe a Data de Admissão para visualizar os prazos.")

        st.markdown("---")
        st.subheader("Eventos Trabalhistas")
        av1,av2,av3 = st.columns(3)
        with av1:
            st.markdown("**Aviso Prévio**")
            dt_aviso = st.text_input("Data Aviso", value=val_campo("DataAvisoPrevio"))
            dias_aviso = st.text_input("Dias Aviso", value=val_campo("DiasAvisoPrevio"))
            term_aviso = st.text_input("Término Aviso", value=term_aviso_val, disabled=True)
        with av2:
            st.markdown("**Licença**")
            dt_lic = st.text_input("Data Licença", value=val_campo("DataLicenca"))
            dias_lic = st.text_input("Dias Licença", value=val_campo("DiasLicenca"))
            term_lic = st.text_input("Término Licença", value=term_lic_val, disabled=True)
        with av3:
            st.markdown("**Férias**")
            dt_fer = st.text_input("Início Férias", value=val_campo("DataFeriasInicio"))
            dias_fer = st.text_input("Dias Férias", value=val_campo("DiasFerias"))
            ret_fer = st.text_input("Retorno Férias", value=ret_fer_val, disabled=True)

        af1,af2 = st.columns(2)
        with af1:
            st.markdown("**Afastamento**")
            dt_af = st.text_input("Data Afastamento", value=val_campo("DataAfastamento"))
            dias_af = st.text_input("Dias Afastamento", value=val_campo("DiasAfastamento"))
            ret_af = st.text_input("Retorno Afastamento", value=ret_af_val, disabled=True)
            tipo_af = st.selectbox("Tipo Afastamento", ["Nenhum", "Doença", "Acidente", "Maternidade"])
        with af2:
            st.markdown("**Desligamento**")
            dt_ped = st.text_input("Data Pedido Conta", value=val_campo("DataPedidoConta"))
            dt_res = st.text_input("Data Rescisão", value=val_campo("DataRescisao"))
            dt_aband = st.text_input("Data Abandono", value=val_campo("DataAbandono"))
            dt_termino_cont = st.text_input("📅 Data Término de Contrato", value=val_campo("DataTerminoContrato"))

        btn_salvar = st.form_submit_button("💾 SALVAR CADASTRO", type="primary", use_container_width=True)
        if btn_salvar:
            matricula_tratada = str(matricula).strip()
            if not matricula_tratada:
                st.error("❌ INFORME A MATRÍCULA!")
                st.stop()
            if tipo_af != "Nenhum" and dt_af.strip():
                situacao = tipo_af

            caminho_final_foto = caminho_foto_atual
            if excluir_foto and caminho_final_foto and os.path.exists(caminho_final_foto):
                os.remove(caminho_final_foto)
                caminho_final_foto = ""
            if nova_foto:
                if caminho_final_foto and os.path.exists(caminho_final_foto):
                    os.remove(caminho_final_foto)
                extensao = os.path.splitext(nova_foto.name)[1].lower()
                nome_foto = f"{matricula_tratada}_foto_{datetime.now().strftime('%Y%m%d%H%M%S')}{extensao}"
                caminho_final_foto = os.path.join(PASTA_FOTOS, nome_foto)
                img = Image.open(nova_foto)
                img.save(caminho_final_foto)

            dados_form = calcular_e_atualizar({
                "mat": matricula_tratada, "nome": nome, "cpf": cpf, "rg": rg, "pis": pis,
                "nasc": nascimento, "adm": admissao, "tel": telefone, "end": endereco,
                "loja": loja, "cargo": cargo, "sal": salario, "situacao": situacao,
                "dt_aviso": dt_aviso, "dias_aviso": dias_aviso, "termino_aviso": term_aviso,
                "dt_lic": dt_lic, "dias_lic": dias_lic, "termino_lic": term_lic,
                "dt_fer": dt_fer, "dias_fer": dias_fer, "retorno_fer": ret_fer,
                "dt_af": dt_af, "dias_af": dias_af, "retorno_af": ret_af,
                "dt_pedido": dt_ped, "dt_rescisao": dt_res, "dt_abandono": dt_aband,
                "dt_termino_cont": dt_termino_cont
            })
            registro_final = {
                "Matricula": dados_form["mat"], "Nome": dados_form["nome"], "CPF": dados_form["cpf"],
                "RG": dados_form["rg"], "PIS": dados_form["pis"], "Nascimento": dados_form["nasc"],
                "Admissao": dados_form["adm"], "Telefone": dados_form["tel"], "Endereco": dados_form["end"],
                "Loja": dados_form["loja"], "Cargo": dados_form["cargo"], "Salario": dados_form["sal"],
                "Situacao": dados_form["situacao"], "DataAvisoPrevio": dados_form["dt_aviso"],
                "DiasAvisoPrevio": dados_form["dias_aviso"], "DataTerminoAviso": dados_form["termino_aviso"],
                "DataFeriasInicio": dados_form["dt_fer"], "DiasFerias": dados_form["dias_fer"],
                "DataRetornoFerias": dados_form["retorno_fer"], "DataPedidoConta": dados_form["dt_pedido"],
                "DataRescisao": dados_form["dt_rescisao"], "DataAbandono": dados_form["dt_abandono"],
                "DataTerminoContrato": dados_form["dt_termino_cont"],
                "DataLicenca": dados_form["dt_lic"], "DiasLicenca": dados_form["dias_lic"],
                "DataTerminoLicenca": dados_form["termino_lic"],
                "DataAfastamento": dados_form["dt_af"], "DiasAfastamento": dados_form["dias_af"],
                "DataRetornoAfastamento": dados_form["retorno_af"],
                "CaminhoFoto": caminho_final_foto
            }
            indice = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"] == dados_form["mat"]].tolist()
            acao_hist = "Atualização Cadastral" if indice else "Novo Cadastro"
            if indice: dados["Base_Dados"].iloc[indice[0]] = registro_final
            else: dados["Base_Dados"] = pd.concat([dados["Base_Dados"], pd.DataFrame([registro_final])], ignore_index=True)
            salvar_dados(dados)
            add_historico_auto(dados_form["mat"], dados_form["nome"], acao_hist, registro_final)
            st.success(f"✅ Salvo! Matrícula: **{dados_form['mat']}**")
            st.rerun()

    if mat_sel.strip() and st.button("🗑️ EXCLUIR REGISTRO", use_container_width=True, type="secondary"):
        if st.checkbox("⚠️ CONFIRMA EXCLUSÃO PERMANENTE?"):
            indice = dados["Base_Dados"].index[dados["Base_Dados"]["Matricula"] == mat_sel.strip()].tolist()
            if indice:
                dados_excluir = dados["Base_Dados"].iloc[indice[0]].to_dict()
                if dados_excluir.get("CaminhoFoto") and os.path.exists(dados_excluir["CaminhoFoto"]):
                    os.remove(dados_excluir["CaminhoFoto"])
                docs_excluir = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"] == mat_sel.strip()]
                for _, d in docs_excluir.iterrows():
                    if os.path.exists(d["Caminho"]): os.remove(d["Caminho"])
                dados["Docs_Funcionarios"] = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"] != mat_sel.strip()]
                dados["Base_Dados"].drop(indice[0], inplace=True)
                salvar_dados(dados)
                add_historico_auto(mat_sel.strip(), dados_excluir["Nome"], "Exclusão de Cadastro", dados_excluir)
                st.success("✅ Registro, foto e documentos excluídos!")
                st.rerun()

    st.markdown("---")
    st.subheader("📎 DOCUMENTOS DO FUNCIONÁRIO")
    if mat_sel.strip() and not reg.empty:
        mat_atual = mat_sel.strip()
        nome_atual = val_campo("Nome")
        tipo_doc = st.selectbox("Tipo de Documento", [
            "RG", "CPF", "PIS", "Carteira de Trabalho", "Comprovante Residência",
            "Exame Admissional", "Exame Demissional", "Contrato", "Atestados",
            "Férias", "Rescisão", "Outros"
        ])
        arquivos_func = st.file_uploader("Anexar documentos", type=["pdf","doc","docx","xls","xlsx","jpg","png"], accept_multiple_files=True, key=f"up_{mat_atual}")
        if arquivos_func and st.button("SALVAR DOCUMENTOS", type="primary"):
            qtd = 0
            for arq in arquivos_func:
                nome_arq = f"{mat_atual}_{tipo_doc}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{arq.name}"
                caminho = os.path.join(PASTA_DOCS_FUNC, nome_arq)
                with open(caminho, "wb") as f: f.write(arq.read())
                dados["Docs_Funcionarios"] = pd.concat([dados["Docs_Funcionarios"], pd.DataFrame([{
                    "Matricula": mat_atual, "Nome": nome_atual, "TipoDoc": tipo_doc,
                    "NomeArquivo": arq.name, "Caminho": caminho,
                    "DataAnexado": datetime.now().strftime("%d/%m/%Y %H:%M")
                }])], ignore_index=True)
                qtd += 1
            salvar_dados(dados)
            st.success(f"✅ {qtd} documento(s) salvo(s)!")
            st.rerun()
        st.markdown("---")
        docs_func = dados["Docs_Funcionarios"][dados["Docs_Funcionarios"]["Matricula"] == mat_atual]
        if docs_func.empty: st.info("📂 Nenhum documento anexado.")
        else:
            st.markdown(f"**Total: {len(docs_func)} documento(s)**")
            for idx, doc in docs_func.iterrows():
                with st.expander(f"📄 {doc['TipoDoc']} - {doc['NomeArquivo']} | {doc['DataAnexado']}"):
                    col_v, col_b, col_e = st.columns([3,1,1])
                    with col_b:
                        with open(doc["Caminho"], "rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=doc["NomeArquivo"], key=f"dw_{idx}")
                    with col_e:
                        if st.button("🗑️ EXCLUIR", key=f"del_{idx}"):
                            if os.path.exists(doc["Caminho"]): os.remove(doc["Caminho"])
                            dados["Docs_Funcionarios"].drop(idx, inplace=True)
                            salvar_dados(dados)
                            st.rerun()
    else:
        st.info("ℹ️ Digite a Matrícula exata para ver/anexar documentos.")

# ================ DEMAIS ABAS ================
with aba2:
    st.subheader("📊 RESUMO GERAL")
    dados_painel = carregar_dados()
    base = dados_painel["Base_Dados"].copy()
    base["Situacao"] = base["Situacao"].fillna("").astype(str).str.strip()

    contagem = {
        "👷 Ativo": len(base[base["Situacao"] == "Ativo"]),
        "📝 Pré-cadastro": len(base[base["Situacao"] == "Pré-cadastro"]),
        "🏖️ Férias": len(base[base["Situacao"] == "Férias"]),
        "🚪 Abandono": len(base[base["Situacao"] == "Abandono"]),
        "⏹️ Término de Contrato": len(base[base["Situacao"] == "Término de Contrato"]),
        "📉 Demitido S/JC": len(base[base["Situacao"] == "Demitido S/JC"]),
        "📉 Demitido C/JC": len(base[base["Situacao"] == "Demitido C/JC"]),
        "🙋 Pedido de Conta": len(base[base["Situacao"] == "Pedido de Conta"]),
        "⚖️ Rescisão Indireta": len(base[base["Situacao"] == "Rescisão Indireta"]),
        "🏥 Doença": len(base[base["Situacao"] == "Doença"]),
        "🚑 Acidente": len(base[base["Situacao"] == "Acidente"]),
        "🤰 Maternidade": len(base[base["Situacao"] == "Maternidade"])
    }

    cols = st.columns(3)
    for i, (rotulo, qtd) in enumerate(contagem.items()):
        cols[i % 3].metric(rotulo, qtd)

    if st.button("🔄 Atualizar Resumo"):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.subheader("🔍 Conferência Rápida - Apenas Férias")
    tab_fer = base[base["Situacao"] == "Férias"][["Matricula","Nome","Loja","DataFeriasInicio","DataRetornoFerias"]]
    if tab_fer.empty:
        st.warning("⚠️ Nenhum funcionário com situação marcada como 'Férias' no momento.")
        st.info("💡 Dica: Se a data de férias estiver preenchida mas a situação não for 'Férias', edite o cadastro e confirme se a situação está selecionada corretamente — os dados não são apagados!")
    else:
        st.dataframe(tab_fer, use_container_width=True, hide_index=True)

with aba3:
    hoje = datetime.now()
    st.subheader("⚠️ PRAZOS DE EXPERIÊNCIA PRÓXIMOS")
    tabela_exp = []
    for _, func in dados["Base_Dados"].iterrows():
        if func["Situacao"] not in ["Ativo","Pré-cadastro"]: continue
        try:
            dt_adm = datetime.strptime(str(func["Admissao"]).strip(), "%d/%m/%Y")
            dias = (hoje - dt_adm).days
            for p in [30,45,60,90]:
                if 0 <= p - dias <=10:
                    tabela_exp.append([func["Matricula"], func["Nome"], func["Loja"], f"{p} dias", f"Faltam {p-dias} dias"])
                    break
        except: pass
    st.dataframe(pd.DataFrame(tabela_exp, columns=["Matrícula","Nome","Loja","Prazo","Dias Restantes"]), use_container_width=True, hide_index=True)

    st.subheader("🗓️ FÉRIAS - POR MÊS DE ADMISSÃO")
    filtro_loja = st.selectbox("Loja", ["Todas"] + lista_lojas(), key="fl")
    filtro_mes = st.selectbox("Mês", MESES, key="fm")
    tabela_fer = []
    for _, f in dados["Base_Dados"].iterrows():
        if f["Situacao"] not in ["Ativo","Pré-cadastro"]: continue
        if filtro_loja != "Todas" and f["Loja"] != filtro_loja: continue
        try:
            dt = datetime.strptime(str(f["Admissao"]).strip(), "%d/%m/%Y")
            if filtro_mes != "Todos" and dt.month != MAP_MES[filtro_mes]: continue
            meses = (hoje.year - dt.year)*12 + (hoje.month - dt.month) - (1 if hoje.day < dt.day else 0)
            if 23 <= meses < 24:
                tabela_fer.append([f["Matricula"], f["Nome"], f["Loja"], f["Cargo"], f["Admissao"], f"{meses}m"])
        except: pass
    st.dataframe(pd.DataFrame(tabela_fer, columns=["Matrícula","Nome","Loja","Cargo","Admissão","Tempo"]), use_container_width=True, hide_index=True)

with aba4:
    st.subheader("📝 HISTÓRICO")
    st.dataframe(dados["Historico"][["DataEvento","TipoEvento","Matricula","Nome","Situacao","Detalhes"]], use_container_width=True, hide_index=True)
    with st.form("add_ev"):
        t,d,det = st.columns([1,1,3])
        te = t.selectbox("Tipo", ["Reunião","Atestado","Advertência","Elogio","Outros"])
        de = d.text_input("Data", value=datetime.now().strftime("%d/%m/%Y"))
        dee = det.text_input("Detalhes")
        if st.form_submit_button("✅ ADICIONAR") and mat_sel.strip():
            rf = dados["Base_Dados"][dados["Base_Dados"]["Matricula"] == mat_sel.strip()]
            if not rf.empty:
                nr = {"DataEvento":de,"TipoEvento":te,"Detalhes":dee}
                nr.update(rf.iloc[0].to_dict())
                ih = dados["Historico"].index[dados["Historico"]["Matricula"] == mat_sel.strip()].tolist()
                if ih: dados["Historico"].iloc[ih[0]] = nr
                else: dados["Historico"] = pd.concat([dados["Historico"], pd.DataFrame([nr])], ignore_index=True)
                salvar_dados(dados)
                st.success("Adicionado!")
                st.rerun()

with aba5:
    st.subheader("📄 RELATÓRIOS")
    rel = st.selectbox("Escolha", ["Prazos Experiência","Ativos","Pré-cadastro","Férias","Afastados","Avisos","Histórico","Individual"])
    if rel == "Individual":
        mr = st.text_input("Matrícula")
        if mr.strip():
            fd = dados["Base_Dados"][dados["Base_Dados"]["Matricula"] == mr.strip()]
            fh = dados["Historico"][dados["Historico"]["Matricula"] == mr.strip()]
            if fd.empty: st.error("Não encontrado")
            elif st.button("GERAR"):
                with pd.ExcelWriter(f"Rel_{mr}_{fd.iloc[0]['Nome'].replace(' ','_')}.xlsx") as arq:
                    fd.to_excel(arq, index=False, sheet_name="Dados")
                    fh.to_excel(arq, index=False, sheet_name="Histórico") if not fh.empty else pd.DataFrame([{"Aviso":"Sem histórico"}]).to_excel(arq, index=False, sheet_name="Histórico")
                with open(f"Rel_{mr}_{fd.iloc[0]['Nome'].replace(' ','_')}.xlsx","rb") as f:
                    st.download_button("⬇️ BAIXAR", f, file_name=f"Rel_{mr}.xlsx")
    elif st.button("GERAR E BAIXAR"):
        if rel == "Prazos Experiência": df = pd.DataFrame(tabela_exp, columns=["Matrícula","Nome","Loja","Prazo","Dias Restantes"])
        elif rel == "Ativos": df = dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Ativo"]
        elif rel == "Pré-cadastro": df = dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Pré-cadastro"]
        elif rel == "Férias": df = dados["Base_Dados"][dados["Base_Dados"]["Situacao"] == "Férias"]
        elif rel == "Afastados": df = dados["Base_Dados"][dados["Base_Dados"]["Situacao"].isin(["Doença","Acidente","Maternidade"])]
        elif rel == "Avisos": df = dados["Base_Dados"][dados["Base_Dados"]["DataAvisoPrevio"].str.strip()!=""]
        else: df = dados["Historico"]
        with pd.ExcelWriter("rel_temp.xlsx") as arq: df.to_excel(arq, index=False, sheet_name=rel)
        with open("rel_temp.xlsx","rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=f"Rel_{rel.replace(' ','_')}.xlsx")
        os.remove("rel_temp.xlsx")

with aba6:
    st.subheader("📎 DOCUMENTOS DAS LOJAS")
    ls = lista_lojas()
    l,m,a = st.columns(3)
    sl = l.selectbox("Loja", ls)
    sm = m.selectbox("Mês", MESES)
    sa = a.selectbox("Ano", [str(x) for x in range(2020, datetime.now().year+2)], index=datetime.now().year-2020)
    st.markdown("---")
    arquivos = st.file_uploader("Anexar arquivos", type=["pdf","doc","docx","xls","xlsx","jpg","png"], accept_multiple_files=True)
    resp = st.text_input("Responsável")
    if arquivos and st.button("SALVAR TODOS", type="primary"):
        salvos = 0
        for arq in arquivos:
            nome = f"{sl}_{sm}_{sa}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{arq.name}"
            cam = os.path.join(PASTA_DOCS, nome)
            with open(cam,"wb") as f: f.write(arq.read())
            dados["Docs_Lojas"] = pd.concat([dados["Docs_Lojas"], pd.DataFrame([{
                "Loja":sl,"Mes":sm,"Ano":sa,"NomeArquivo":arq.name,"Caminho":cam,
                "DataAnexado":datetime.now().strftime("%d/%m/%Y %H:%M"),"Responsavel":resp
            }])], ignore_index=True)
            salvos += 1
        salvar_dados(dados)
        st.success(f"✅ {salvos} arquivo(s) salvo(s)!")
        st.rerun()
    st.markdown("---")
    filt = dados["Docs_Lojas"].copy()
    if sl != "Todas": filt = filt[filt["Loja"]==sl]
    if sm != "Todos": filt = filt[filt["Mes"]==sm]
    filt = filt[filt["Ano"]==sa]
    if filt.empty: st.info("Nenhum documento.")
    else:
        for i,d in filt.iterrows():
            with st.expander(f"📄 {d['NomeArquivo']} | {d['Mes']}/{d['Ano']}"):
                with open(d["Caminho"],"rb") as f: st.download_button("⬇️ BAIXAR", f, file_name=d["NomeArquivo"], key=f"d{i}")
                if st.button("🗑️ EXCLUIR", key=f"x{i}"):
                    os.remove(d["Caminho"])
                    dados["Docs_Lojas"].drop(i,inplace=True)
                    salvar_dados(dados)
                    st.rerun()

# ================ ABA 7 - GUIA DE VIAGEM ================
with aba7:
    st.subheader("🚗 GUIA DE VIAGEM")
    dados_v = carregar_dados()

    sub1, sub2 = st.tabs(["➕ Cadastrar Viagem", "🔍 Consultar Viagens"])

    # ---------- Cadastrar Viagem ----------
    with sub1:
        st.markdown("---")
        mat_viagem = st.text_input("Matrícula do Funcionário", placeholder="Digite a matrícula exata", key="mat_viagem")
        nome_viagem = ""
        if mat_viagem.strip():
            df_func = dados_v["Base_Dados"][dados_v["Base_Dados"]["Matricula"] == mat_viagem.strip()]
            if not df_func.empty:
                nome_viagem = df_func.iloc[0]["Nome"]
                st.success(f"Funcionário encontrado: **{nome_viagem}**")
            else:
                st.warning("⚠️ Funcionário não encontrado. Você pode cadastrar mesmo assim.")

        with st.form("form_viagem", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                destino = st.text_input("Destino da Viagem *")
                dt_inicio = st.text_input("Data Início (dd/mm/aaaa) *")
                dt_fim = st.text_input("Data Fim (dd/mm/aaaa)")
            with c2:
                valor_recebido = st.text_input("Valor Recebido para Despesas (R$) *", placeholder="0,00")
                status = st.selectbox("Status", STATUS_VIAGEM)
                observacoes = st.text_area("Observações", height=100)
            with c3:
                st.markdown("**Resumo Financeiro**")
                st.info("O valor gasto e o restante serão calculados automaticamente conforme as prestações de contas forem adicionadas.")

            if st.form_submit_button("💾 SALVAR VIAGEM", type="primary", use_container_width=True):
                if not mat_viagem.strip() or not destino.strip() or not dt_inicio.strip() or not valor_recebido.strip():
                    st.error("❌ Preencha os campos obrigatórios!")
                    st.stop()
                try:
                    vrec = float(valor_recebido.replace(",",".").replace("R$","").strip())
                except:
                    st.error("❌ Valor recebido inválido!")
                    st.stop()
                novo_id = gerar_id_viagem()
                reg_viagem = {
                    "ID": novo_id,
                    "Matricula": mat_viagem.strip(),
                    "Nome": nome_viagem if nome_viagem else "",
                    "Destino": destino.strip(),
                    "DataInicio": dt_inicio.strip(),
                    "DataFim": dt_fim.strip(),
                    "ValorRecebido": f"{vrec:.2f}",
                    "ValorGasto": "0.00",
                    "ValorRestante": f"{vrec:.2f}",
                    "Status": status,
                    "Observacoes": observacoes,
                    "DataCadastro": datetime.now().strftime("%d/%m/%Y %H:%M")
                }
                dados_v["Viagens"] = pd.concat([dados_v["Viagens"], pd.DataFrame([reg_viagem])], ignore_index=True)
                salvar_dados(dados_v)
                st.success(f"✅ Viagem **{novo_id}** cadastrada com sucesso!")
                st.rerun()

    # ---------- Consultar Viagens ----------
    with sub2:
        st.markdown("---")
        st.markdown("**Filtros de Pesquisa**")
        fcol1, fcol2, fcol3, fcol4 = st.columns(4)
        with fcol1:
            f_id = st.text_input("ID da Viagem", key="f_id")
        with fcol2:
            f_mat = st.text_input("Matrícula", key="f_mat")
        with fcol3:
            f_dest = st.text_input("Destino", key="f_dest")
        with fcol4:
            f_status = st.selectbox("Status", ["Todos"] + STATUS_VIAGEM, key="f_status")

        df_viagens = dados_v["Viagens"].copy()
        if f_id.strip():
            df_viagens = df_viagens[df_viagens["ID"].str.contains(f_id, case=False, na=False)]
        if f_mat.strip():
            df_viagens = df_viagens[df_viagens["Matricula"].str.contains(f_mat, case=False, na=False)]
        if f_dest.strip():
            df_viagens = df_viagens[df_viagens["Destino"].str.contains(f_dest, case=False, na=False)]
        if f_status != "Todos":
            df_viagens = df_viagens[df_viagens["Status"] == f_status]

        if df_viagens.empty:
            st.info("📂 Nenhuma viagem encontrada com os filtros aplicados.")
        else:
            st.markdown(f"**{len(df_viagens)} viagem(s) encontrada(s)**")
            for idx_v, row_v in df_viagens.iterrows():
                id_v = row_v["ID"]
                with st.expander(f"🚗 {id_v} | {row_v['Destino']} | {row_v['Matricula']} - {row_v['Nome']} | Status: {row_v['Status']}"):
                    col_info, col_fin = st.columns([2,1])
                    with col_info:
                        st.markdown(f"**Destino:** {row_v['Destino']}")
                        st.markdown(f"**Período:** {row_v['DataInicio']} até {row_v['DataFim']}")
                        st.markdown(f"**Funcionário:** {row_v['Nome']} (Matrícula: {row_v['Matricula']})")
                        st.markdown(f"**Observações:** {row_v['Observacoes']}")
                        st.markdown(f"**Cadastrado em:** {row_v['DataCadastro']}")
                    with col_fin:
                        try:
                            vrec = float(str(row_v["ValorRecebido"]).replace(",","."))
                            vgas = float(str(row_v["ValorGasto"]).replace(",","."))
                            vrest = float(str(row_v["ValorRestante"]).replace(",","."))
                        except:
                            vrec = vgas = vrest = 0.0
                        st.metric("💰 Recebido", f"R$ {vrec:,.2f}")
                        st.metric("💸 Gasto", f"R$ {vgas:,.2f}")
                        st.metric("🟢 Restante", f"R$ {vrest:,.2f}")
                        if vrest < 0:
                            st.error("⚠️ Atenção: ultrapassou o valor recebido!")

                    st.markdown("---")
                    st.markdown("**📝 PRESTAÇÃO DE CONTAS**")

                    # Adicionar nova despesa
                    with st.form(f"form_prestacao_{id_v}", clear_on_submit=True):
                        st.markdown("**Adicionar Nova Despesa**")
                        p1, p2, p3 = st.columns(3)
                        with p1:
                            dt_desp = st.text_input("Data da Despesa (dd/mm/aaaa)", value=datetime.now().strftime("%d/%m/%Y"), key=f"dt_desp_{id_v}")
                            tipo_desp = st.selectbox("Tipo de Despesa", TIPOS_DESPESA, key=f"tipo_desp_{id_v}")
                        with p2:
                            desc_desp = st.text_input("Descrição / Detalhes", key=f"desc_desp_{id_v}")
                            valor_desp = st.text_input("Valor (R$) *", placeholder="0,00", key=f"val_desp_{id_v}")
                        with p3:
                            comprovante = st.file_uploader("Comprovante (opcional)", type=["pdf","jpg","jpeg","png"], key=f"comp_{id_v}")
                        if st.form_submit_button("➕ ADICIONAR DESPESA", type="primary", use_container_width=True):
                            if not valor_desp.strip():
                                st.error("❌ Informe o valor da despesa!")
                                st.stop()
                            try:
                                v_desp = float(valor_desp.replace(",",".").replace("R$","").strip())
                            except:
                                st.error("❌ Valor inválido!")
                                st.stop()
                            cam_comp = ""
                            if comprovante:
                                ext = os.path.splitext(comprovante.name)[1].lower()
                                nome_comp = f"{id_v}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}{ext}"
                                cam_comp = os.path.join(PASTA_COMPROVANTES, nome_comp)
                                with open(cam_comp, "wb") as f: f.write(comprovante.read())
                            novo_id_prest = f"P{len(dados_v['PrestacaoContas'])+1:04d}"
                            reg_prest = {
                                "ID": novo_id_prest,
                                "IDViagem": id_v,
                                "DataDespesa": dt_desp.strip(),
                                "TipoDespesa": tipo_desp,
                                "Descricao": desc_desp.strip(),
                                "Valor": f"{v_desp:.2f}",
                                "CaminhoComprovante": cam_comp,
                                "DataRegistro": datetime.now().strftime("%d/%m/%Y %H:%M")
                            }
                            dados_v["PrestacaoContas"] = pd.concat([dados_v["PrestacaoContas"], pd.DataFrame([reg_prest])], ignore_index=True)
                            salvar_dados(dados_v)
                            atualizar_totais_viagem(id_v)
                            st.success("✅ Despesa adicionada!")
                            st.rerun()

                    # Listar despesas da viagem
                    despesas = dados_v["PrestacaoContas"][dados_v["PrestacaoContas"]["IDViagem"] == id_v]
                    if despesas.empty:
                        st.info("📂 Nenhuma despesa registrada para esta viagem.")
                    else:
                        st.markdown(f"**Total de despesas: {len(despesas)}**")
                        for idx_d, row_d in despesas.iterrows():
                            with st.container(border=True):
                                dcol1, dcol2, dcol3, dcol4 = st.columns([2,2,1,1])
                                with dcol1:
                                    st.markdown(f"**{row_d['TipoDespesa']}**")
                                    st.markdown(f"📅 {row_d['DataDespesa']} | 💰 R$ {row_d['Valor']}")
                                with dcol2:
                                    st.markdown(f"_{row_d['Descricao']}_")
                                    st.caption(f"Registrado em: {row_d['DataRegistro']}")
                                with dcol3:
                                    if row_d["CaminhoComprovante"] and os.path.exists(row_d["CaminhoComprovante"]):
                                        with open(row_d["CaminhoComprovante"], "rb") as f:
                                            st.download_button("⬇️ Comprovante", f, file_name=os.path.basename(row_d["CaminhoComprovante"]), key=f"dwc_{idx_d}")
                                    else:
                                        st.caption("Sem comprovante")
                                with dcol4:
                                    if st.button("🗑️ Excluir", key=f"del_desp_{idx_d}"):
                                        if row_d["CaminhoComprovante"] and os.path.exists(row_d["CaminhoComprovante"]):
                                            os.remove(row_d["CaminhoComprovante"])
                                        dados_v["PrestacaoContas"].drop(idx_d, inplace=True)
                                        salvar_dados(dados_v)
                                        atualizar_totais_viagem(id_v)
                                        st.rerun()

                    # Alterar status da viagem
                    st.markdown("---")
                    novo_status = st.selectbox("Alterar Status da Viagem", STATUS_VIAGEM, index=STATUS_VIAGEM.index(row_v["Status"]) if row_v["Status"] in STATUS_VIAGEM else 0, key=f"st_{id_v}")
                    if st.button("💾 ATUALIZAR STATUS", key=f"btn_st_{id_v}"):
                        ix = dados_v["Viagens"].index[dados_v["Viagens"]["ID"] == id_v].tolist()
                        if ix:
                            dados_v["Viagens"].at[ix[0], "Status"] = novo_status
                            salvar_dados(dados_v)
                            st.success("✅ Status atualizado!")
                            st.rerun()

                    # Excluir viagem
                    if st.checkbox(f"⚠️ Confirmar exclusão da viagem {id_v}", key=f"conf_del_v_{id_v}"):
                        if st.button("🗑️ EXCLUIR VIAGEM", type="secondary", key=f"btn_del_v_{id_v}"):
                            docs_v = dados_v["PrestacaoContas"][dados_v["PrestacaoContas"]["IDViagem"] == id_v]
                            for _, dv in docs_v.iterrows():
                                if dv["CaminhoComprovante"] and os.path.exists(dv["CaminhoComprovante"]):
                                    os.remove(dv["CaminhoComprovante"])
                            dados_v["PrestacaoContas"] = dados_v["PrestacaoContas"][dados_v["PrestacaoContas"]["IDViagem"] != id_v]
                            dados_v["Viagens"].drop(idx_v, inplace=True)
                            salvar_dados(dados_v)
                            st.success("✅ Viagem e despesas excluídas!")
                            st.rerun()

        st.markdown("---")
        st.markdown("**📊 Resumo Geral das Viagens**")
        if not dados_v["Viagens"].empty:
            total_rec = 0.0
            total_gas = 0.0
            total_res = 0.0
            for _, rv in dados_v["Viagens"].iterrows():
                try:
                    total_rec += float(str(rv["ValorRecebido"]).replace(",","."))
                    total_gas += float(str(rv["ValorGasto"]).replace(",","."))
                    total_res += float(str(rv["ValorRestante"]).replace(",","."))
                except:
                    pass
            cres1, cres2, cres3 = st.columns(3)
            cres1.metric("💰 Total Recebido", f"R$ {total_rec:,.2f}")
            cres2.metric("💸 Total Gasto", f"R$ {total_gas:,.2f}")
            cres3.metric("🟢 Total Restante", f"R$ {total_res:,.2f}")
        else:
            st.info("Nenhuma viagem cadastrada.")

# ================ ABA 8 - LOJAS E CARGOS ================
with aba8:
    st.subheader("⚙️ CADASTRO DE LOJAS E CARGOS")
    dados = carregar_dados()
    col1, col2 = st.columns(2)
    with col1:
        nova_loja = st.text_input("Nova Loja")
        if st.button("➕ ADICIONAR LOJA", type="primary") and nova_loja.strip():
            if not dados["Auxiliares"]["Loja"].str.strip().eq(nova_loja.strip()).any():
                dados["Auxiliares"] = pd.concat([dados["Auxiliares"], pd.DataFrame([{"Loja": nova_loja.strip(), "Cargo": ""}])], ignore_index=True)
                salvar_dados(dados)
                st.success("✅ Loja cadastrada!")
                st.rerun()
            else: st.warning("⚠️ Já existe!")
    with col2:
        novo_cargo = st.text_input("Novo Cargo")
        if st.button("➕ ADICIONAR CARGO", type="primary") and novo_cargo.strip():
            if not dados["Auxiliares"]["Cargo"].str.strip().eq(novo_cargo.strip()).any():
                dados["Auxiliares"] = pd.concat([dados["Auxiliares"], pd.DataFrame([{"Loja": "", "Cargo": novo_cargo.strip()}])], ignore_index=True)
                salvar_dados(dados)
                st.success("✅ Cargo cadastrado!")
                st.rerun()
            else: st.warning("⚠️ Já existe!")
