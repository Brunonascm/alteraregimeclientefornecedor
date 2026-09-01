import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Atualizador de Regime - Domínio", layout="wide")

# Código IBGE padrão de contingência (São Paulo - SP)
CODIGO_IBGE_PADRAO = "3550308"

# Opções oficiais para Fornecedores (Reg. 0020)
REGIME_OPTIONS_FORNECEDORES = {
    "N": "N - Normal",
    "M": "M - ME",
    "E": "E - EPP",
    "O": "O - Outros",
    "S": "S - ME - Simples Nacional",
    "P": "P - EPP - Simples Nacional",
    "U": "U - Imune",
    "I": "I - Isenta"
}

# Opções oficiais para Clientes (Reg. 0010) - Não aceita 'S' nem 'P'
REGIME_OPTIONS_CLIENTES = {
    "N": "N - Normal",
    "M": "M - ME",
    "E": "E - EPP",
    "O": "O - Outros",
    "U": "U - Imune",
    "I": "I - Isenta"
}

# --- FUNÇÃO DE NORMALIZAÇÃO DE CÓDIGOS ---
def normalizar_codigo(codigo):
    if not codigo:
        return ""
    val = str(codigo).strip()
    if val.endswith('.0'):
        val = val[:-2]
    if val.isdigit():
        return str(int(val))
    return val.lower()


# --- FUNÇÃO SILENCIOSA DE LEITURA DE MUNICÍPIOS ---
def ler_municipios_silencioso():
    municipio_map = {}
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    caminho_arquivo = os.path.join(diretorio_atual, "Municipios.txt")
    
    if os.path.exists(caminho_arquivo):
        try:
            with open(caminho_arquivo, 'r', encoding='cp1252') as f:
                for linha in f:
                    linha_limpa = linha.strip()
                    if not linha_limpa:
                        continue
                    colunas = linha_limpa.split("\t")
                    if len(colunas) >= 5:
                        id_interno = normalizar_codigo(colunas[0])
                        codigo_ibge = colunas[4].strip()
                        if id_interno and codigo_ibge:
                            municipio_map[id_interno] = codigo_ibge
        except:
            pass
    return municipio_map


# Inicializa o mapa de municípios (em segundo plano)
mapa_municipios = ler_municipios_silencioso()


# --- TÍTULO PRINCIPAL ---
st.title("🔄 Atualizador em Lote: Regime de Apuração (Domínio)")
st.write(
    "Carregue o arquivo TXT extraído do SQL, altere os regimes em lote, via Excel ou individualmente, "
    "e gere o arquivo final formatado no **Leiaute Padrão de Importação (Separado por Pipe |)**."
)

# --- SELETOR DE MÓDULO ---
tipo_registro = st.radio(
    "Escolha o tipo de cadastro que deseja processar:",
    options=["Clientes (Reg. 0010)", "Fornecedores (Reg. 0020)"],
    horizontal=True
)

# Define o conjunto correto de opções de regime para a tela
regime_options_atual = REGIME_OPTIONS_CLIENTES if tipo_registro == "Clientes (Reg. 0010)" else REGIME_OPTIONS_FORNECEDORES

# Upload do arquivo original do SQL
uploaded_file = st.file_uploader(f"1️⃣ Selecione o arquivo TXT do SQL de {tipo_registro}", type=["txt"])

# --- PROCESSAMENTO DO ARQUIVO SQL ---
if uploaded_file is not None:
    file_key = f"{tipo_registro}_{uploaded_file.name}_{uploaded_file.size}"
    
    if "current_file_key" not in st.session_state or st.session_state.current_file_key != file_key:
        st.session_state.current_file_key = file_key
        
        raw_bytes = uploaded_file.getvalue()
        try:
            content = raw_bytes.decode("cp1252")
        except UnicodeDecodeError:
            content = raw_bytes.decode("latin-1")
            
        lines = content.splitlines()
        parsed_data = []
        
        for idx, line in enumerate(lines):
            if not line.strip():
                continue
                
            cols = line.split("\t")
            if idx == 0 and ("codi_emp" in cols[0].lower() or "efclientes" in cols[0].lower()):
                continue
                
            def get_val(index, default=""):
                if index < len(cols):
                    val = cols[index].strip()
                    if val.startswith('"') and val.endswith('"'):
                        val = val[1:-1]
                    if val.upper() == "NULL":
                        val = ""
                    return val
                return default

            # --- MAPEAR BASEADO NO TIPO DE REGISTRO ---
            if tipo_registro == "Clientes (Reg. 0010)":
                cnpj_cpf = get_val(11)       # efclientes.cgce_cli
                razao_social = get_val(5)    # efclientes.nomr_cli
                if not cnpj_cpf or not razao_social:
                    continue
                
                raw_date = get_val(30)       # efclientes.CADASTRO_CLI
                
                # Tratamento de Regime para Clientes (Converte S -> M e P -> E automaticamente)
                regime_atual = get_val(24).upper()
                if regime_atual == "S":
                    regime_atual = "M"
                elif regime_atual == "P":
                    regime_atual = "E"
                
                codm_interno = get_val(27) if get_val(27) else get_val(4)
                
                # Órgão público federal / Natureza jurídica - padrão '7' (Empresa Privada)
                orgao_pub = get_val(25)
                if not orgao_pub or orgao_pub not in ["1", "2", "3", "4", "5", "6", "7", "8", "9"]:
                    orgao_pub = "7"
                
                possui_inter = get_val(33).upper() # efclientes.POSSUI_INTERDEPENDENCIA_CLI
                if possui_inter not in ["S", "N"]:
                    possui_inter = "N"

                insc_paa = get_val(43).upper() # efclientes.INSCRITO_PROGRAMA_AQUISICAO_ALIMENTOS_CLI
                if insc_paa not in ["S", "N"]:
                    insc_paa = "N"

                agro_ent = get_val(15).upper()
                if agro_ent not in ["S", "N"]:
                    agro_ent = "N"

                icms_ent = get_val(16).upper()
                if icms_ent not in ["S", "N"]:
                    icms_ent = "N"

                parsed_data.append({
                    "codigo_entidade": get_val(1),
                    "cgce_entidade": cnpj_cpf,
                    "nomr_entidade": razao_social,
                    "nome_entidade": get_val(6),
                    "ende_entidade": get_val(7),
                    "nume_entidade": get_val(8),
                    "complemento_entidade": get_val(29),
                    "bair_entidade": get_val(21),
                    "codm_interno": codm_interno,
                    "sigl_est": get_val(2),
                    "codigo_pais": get_val(26),
                    "cepe_entidade": get_val(10),
                    "insc_entidade": get_val(12),
                    "imun_entidade": get_val(20),
                    "insc_suframa": get_val(28),
                    "dddf_entidade": get_val(18),
                    "fone_entidade": get_val(13),
                    "faxe_entidade": get_val(14),
                    "cadastro_data": raw_date,
                    "codi_cta": get_val(3),
                    "agro_entidade": agro_ent,
                    "regime_entidade": regime_atual,
                    "icms_entidade": icms_ent,
                    "aliq_cli": get_val(19),
                    "categoria_estabel": get_val(23),
                    "orgao_pub_federal": orgao_pub,
                    "possui_interdependencia": possui_inter,
                    "percentual_carga_media": get_val(39),
                    "inscrito_paa": insc_paa,
                    "tins_entidade": get_val(17),
                    "sequencial_paj": get_val(48),
                })
                
            else:
                # Fornecedores (Reg. 0020)
                cnpj_cpf = get_val(11)       # effornecedores.cgce_for
                razao_social = get_val(5)    # effornecedores.nomr_for
                if not cnpj_cpf or not razao_social:
                    continue
                
                raw_date = get_val(30)       # effornecedores.CADASTRO_FOR
                regime_atual = get_val(25).upper()
                codm_interno = get_val(27) if get_val(27) else get_val(3)

                possui_inter = get_val(33).upper()
                if possui_inter not in ["S", "N"]:
                    possui_inter = "N"

                cprb_for = get_val(44).upper()
                if cprb_for not in ["S", "N"]:
                    cprb_for = "N"

                agro_ent = get_val(15).upper()
                if agro_ent not in ["S", "N"]:
                    agro_ent = "N"

                icms_ent = get_val(16).upper()
                if icms_ent not in ["S", "N"]:
                    icms_ent = "N"

                parsed_data.append({
                    "codigo_entidade": get_val(1),
                    "cgce_entidade": cnpj_cpf,
                    "nomr_entidade": razao_social,
                    "nome_entidade": get_val(6),
                    "ende_entidade": get_val(7),
                    "nume_entidade": get_val(8),
                    "complemento_entidade": get_val(29),
                    "bair_entidade": get_val(20),
                    "codm_interno": codm_interno,
                    "sigl_est": get_val(2),
                    "codigo_pais": get_val(26),
                    "cepe_entidade": get_val(10),
                    "insc_entidade": get_val(12),
                    "imun_entidade": get_val(19),
                    "insc_suframa": get_val(28),
                    "dddf_entidade": get_val(18),
                    "fone_entidade": get_val(13),
                    "faxe_entidade": get_val(14),
                    "cadastro_data": raw_date,
                    "codi_cta": get_val(4),
                    "agro_entidade": agro_ent,
                    "regime_entidade": regime_atual,
                    "icms_entidade": icms_ent,
                    "categoria_estabel": get_val(22),
                    "iestst_entidade": get_val(23),
                    "email_entidade": get_val(24),
                    "possui_interdependencia": possui_inter,
                    "contribuinte_cprb": cprb_for,
                    "sequencial_paj": get_val(47),
                    "tins_entidade": get_val(17)
                })

        st.session_state.df = pd.DataFrame(parsed_data)

    # Verifica se os dados estão na sessão
    if "df" in st.session_state and len(st.session_state.df) > 0:
        st.success(f"✔️ {len(st.session_state.df)} registros carregados do SQL ({tipo_registro}).")

        # --- SEÇÃO EXCEL ROUNDTRIP (EXPORTAR / REIMPORTAR) ---
        st.subheader("📦 Integração com Excel")
        col_ex1, col_ex2 = st.columns(2)
        
        with col_ex1:
            st.write("Baixe a planilha para ajustar os regimes de forma massiva fora do sistema:")
            excel_buffer = io.BytesIO()
            excel_df = st.session_state.df[["codigo_entidade", "cgce_entidade", "nomr_entidade", "regime_entidade"]].copy()
            excel_df.columns = ["Código", "CNPJ_CPF", "Nome_Razao_Social", "Regime de Apuração"]
            
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                excel_df.to_excel(writer, index=False, sheet_name="Regimes")
                
            st.download_button(
                label=f"📥 Baixar Relação em Excel (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=f"relacao_{tipo_registro.split(' ')[0].lower()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col_ex2:
            st.write("Suba o Excel ajustado de volta para atualizar o sistema instantaneamente:")
            uploaded_excel = st.file_uploader("Subir Excel editado", type=["xlsx"], label_visibility="collapsed")
            
            if uploaded_excel is not None:
                try:
                    excel_imported = pd.read_excel(uploaded_excel)
                    excel_imported.columns = [str(c).strip().lower() for c in excel_imported.columns]
                    
                    col_id = None
                    col_regime = None
                    for c in excel_imported.columns:
                        if 'codigo' in c or 'codi_' in c:
                            col_id = c
                        if 'regime' in c or 'regime_entidade' in c:
                            col_regime = c
                            
                    if col_id and col_regime:
                        mapa_excel = {}
                        for _, row_ex in excel_imported.iterrows():
                            id_ent = normalizar_codigo(row_ex[col_id])
                            reg_val = str(row_ex[col_regime]).strip().upper()
                            if reg_val:
                                reg_val = reg_val[0]
                            
                            # Conversão para Clientes
                            if "Clientes" in tipo_registro:
                                if reg_val == "S": reg_val = "M"
                                elif reg_val == "P": reg_val = "E"
                            
                            if reg_val in regime_options_atual:
                                mapa_excel[id_ent] = reg_val
                                
                        atualizados = 0
                        for idx, row_main in st.session_state.df.iterrows():
                            id_main = normalizar_codigo(row_main["codigo_entidade"])
                            if id_main in mapa_excel:
                                st.session_state.df.at[idx, "regime_entidade"] = mapa_excel[id_main]
                                atualizados += 1
                        
                        st.success(f"✔️ Sucesso! {atualizados} regimes atualizados do Excel.")
                        st.rerun()
                    else:
                        st.error("Erro: Colunas 'Código' e 'Regime de Apuração' não foram localizadas na planilha.")
                except Exception as e:
                    st.error(f"Erro ao processar a planilha: {e}")

        # --- SEÇÃO ALTERAÇÃO EM LOTE ---
        st.subheader("🛠️ Alteração Rápida em Lote")
        col_b1, col_b2 = st.columns([2, 1])
        
        with col_b1:
            regime_bulk = st.selectbox(
                "Selecione o Regime de Apuração para aplicar a TODOS:",
                options=list(regime_options_atual.keys()),
                format_func=lambda x: regime_options_atual[x]
            )
        with col_b2:
            st.write(" ")
            st.write(" ")
            if st.button("Aplicar a Todos"):
                st.session_state.df["regime_entidade"] = regime_bulk
                st.success("Regime alterado em todos os registros!")
                st.rerun()

        # --- SEÇÃO TABELA EDITÁVEL ---
        st.subheader(f"📝 Tabela de {tipo_registro}")
        
        edited_df = st.data_editor(
            st.session_state.df,
            column_order=["codigo_entidade", "cgce_entidade", "nomr_entidade", "regime_entidade"],
            column_config={
                "codigo_entidade": st.column_config.TextColumn("Código", disabled=True),
                "cgce_entidade": st.column_config.TextColumn("CNPJ / CPF", disabled=True),
                "nomr_entidade": st.column_config.TextColumn("Nome / Razão Social", disabled=True),
                "regime_entidade": st.column_config.SelectboxColumn(
                    "Regime de Apuração",
                    options=list(regime_options_atual.keys()),
                    format_func=lambda x: regime_options_atual[x],
                    required=True
                )
            },
            width="stretch",
            num_rows="fixed",
            key="entidades_editor"
        )

        # --- SEÇÃO EXPORTAÇÃO REGISTRO ---
        st.subheader("💾 Gerar Arquivo Domínio (Leiaute Padrão)")
        
        nome_botao_txt = "Gerar Arquivo de Importação (TXT Clientes - Reg 0010)" if "Clientes" in tipo_registro else "Gerar Arquivo de Importação (TXT Fornecedores - Reg 0020)"
        
        if st.button(nome_botao_txt):
            output = io.StringIO()
            
            for _, row in edited_df.iterrows():
                # Tratamento de IBGE
                chave_busca = normalizar_codigo(row["codm_interno"])
                if mapa_municipios and chave_busca in mapa_municipios:
                    codigo_municipio_final = mapa_municipios[chave_busca]
                else:
                    codigo_municipio_final = CODIGO_IBGE_PADRAO

                # Tratamento de Data
                raw_date = row["cadastro_data"]
                formatted_date = ""
                if raw_date:
                    raw_date_clean = str(raw_date).split(" ")[0].strip()
                    if "-" in raw_date_clean:
                        parts = raw_date_clean.split("-")
                        if len(parts) == 3:
                            formatted_date = f"{parts[2]}/{parts[1]}/{parts[0]}"
                    else:
                        formatted_date = raw_date_clean

                if "Clientes" in tipo_registro:
                    # REGISTRO 0010 - EXATAMENTE 32 COLUNAS CONFORME MANUAL (SEM DESLOCAMENTOS)
                    line_fields = [
                        "0010",                                                     # 1. Identificação do Registro
                        row["cgce_entidade"],                                       # 2. Inscrição (CNPJ/CPF)
                        row["nomr_entidade"][:150],                                 # 3. Razão Social
                        row["nome_entidade"][:40] if row["nome_entidade"] else row["nomr_entidade"][:40], # 4. Apelido
                        row["ende_entidade"],                                       # 5. Endereço
                        row["nume_entidade"],                                       # 6. Número do endereço
                        row["complemento_entidade"],                                # 7. Complemento
                        row["bair_entidade"],                                       # 8. Bairro
                        codigo_municipio_final,                                     # 9. Código do município (IBGE)
                        row["sigl_est"],                                            # 10. UF
                        row["codigo_pais"] if row["codigo_pais"] else "1058",       # 11. Código do País
                        row["cepe_entidade"],                                       # 12. CEP
                        row["insc_entidade"],                                       # 13. Inscrição Estadual
                        row["imun_entidade"],                                       # 14. Inscrição Municipal
                        row["insc_suframa"],                                        # 15. Inscrição Suframa
                        row["dddf_entidade"],                                       # 16. DDD
                        row["fone_entidade"],                                       # 17. Telefone
                        row["faxe_entidade"],                                       # 18. FAX
                        formatted_date,                                             # 19. Data do cadastro
                        row["codi_cta"],                                            # 20. Conta contábil cliente
                        "",                                                         # 21. Conta contábil fornecedor (Vazio)
                        row["agro_entidade"] if row["agro_entidade"] else "N",       # 22. Agropecuário (S/N)
                        "7",                                                        # 23. Natureza jurídica / Órgão Público Federal (Default '7' = Empresa Privada)
                        row["regime_entidade"],                                     # 24. Regime de apuração (Garante N, M, E, O, U, I)
                        row["icms_entidade"] if row["icms_entidade"] else "N",       # 25. Contribuinte ICMS (S/N)
                        row["aliq_cli"],                                            # 26. Alíquota ICMS
                        row["categoria_estabel"],                                   # 27. Categoria do estabelecimento
                        row["possui_interdependencia"] if row["possui_interdependencia"] else "N", # 28. Interdependência com a empresa (S/N)
                        row["percentual_carga_media"],                              # 29. MT - Percentual Carga Média
                        row["inscrito_paa"] if row["inscrito_paa"] else "N",        # 30. Inscrito no PAA (S/N)
                        row["tins_entidade"],                                       # 31. Tipo Inscrição (1=CAEPF)
                        row["sequencial_paj"]                                       # 32. Processo administrativo/judicial
                    ]
                    file_name = "IMPORTACAO_REGISTRO_0010.txt"
                else:
                    # REGISTRO 0020 - EXATAMENTE 33 COLUNAS CONFORME MANUAL (SEM DESLOCAMENTOS)
                    line_fields = [
                        "0020",                                                     # 1. Identificação do Registro
                        row["cgce_entidade"],                                       # 2. Inscrição (CNPJ/CPF)
                        row["nomr_entidade"][:150],                                 # 3. Razão Social
                        row["nome_entidade"][:40] if row["nome_entidade"] else row["nomr_entidade"][:40], # 4. Apelido
                        row["ende_entidade"],                                       # 5. Endereço
                        row["nume_entidade"],                                       # 6. Número do endereço
                        row["complemento_entidade"],                                # 7. Complemento
                        row["bair_entidade"],                                       # 8. Bairro
                        codigo_municipio_final,                                     # 9. Código do município (IBGE)
                        row["sigl_est"],                                            # 10. UF
                        row["codigo_pais"] if row["codigo_pais"] else "1058",       # 11. Código do País
                        row["cepe_entidade"],                                       # 12. CEP
                        row["insc_entidade"],                                       # 13. Inscrição Estadual
                        row["imun_entidade"],                                       # 14. Inscrição Municipal
                        row["insc_suframa"],                                        # 15. Inscrição Suframa
                        row["dddf_entidade"],                                       # 16. DDD
                        row["fone_entidade"],                                       # 17. Telefone
                        row["faxe_entidade"],                                       # 18. FAX
                        formatted_date,                                             # 19. Data do cadastro
                        row["codi_cta"],                                            # 20. Conta contábil fornecedor
                        "",                                                         # 21. Conta contábil cliente (Vazio)
                        row["agro_entidade"] if row["agro_entidade"] else "N",       # 22. Agropecuário (S/N)
                        "7",                                                        # 23. Natureza jurídica / Órgão Público Federal (Default '7' = Empresa Privada)
                        row["regime_entidade"],                                     # 24. Regime de apuração (Garante N, M, E, O, S, P, U, I)
                        row["icms_entidade"] if row["icms_entidade"] else "N",       # 25. Contribuinte ICMS (S/N)
                        "",                                                         # 26. Alíquota ICMS (Vazio)
                        row["categoria_estabel"],                                   # 27. Categoria do estabelecimento
                        row["iestst_entidade"],                                     # 28. Inscrição Estadual ST
                        row["email_entidade"],                                      # 29. Email
                        row["possui_interdependencia"] if row["possui_interdependencia"] else "N", # 30. Interdependência com a empresa (S/N)
                        row["contribuinte_cprb"] if row["contribuinte_cprb"] else "N", # 31. Contribuinte da CPRB (S/N)
                        row["sequencial_paj"],                                      # 32. Processo administrativo/judicial
                        row["tins_entidade"]                                        # 33. Tipo Inscrição (1=CAEPF)
                    ]
                    file_name = "IMPORTACAO_REGISTRO_0020.txt"
                
                # Une com Pipe "|" tratando valores nulos e limpando espaços nas bordas
                cleaned_line = [str(f).strip() if f is not None else "" for f in line_fields]
                output.write("|".join(cleaned_line) + "\n")
            
            # Codificação ANSI (CP1252) exigida pelo importador da Domínio
            file_data = output.getvalue().encode("cp1252", errors="replace")
            
            st.download_button(
                label=f"📥 Baixar {file_name}",
                data=file_data,
                file_name=file_name,
                mime="text/plain"
            )
            st.success(f"🎉 Arquivo {file_name} gerado com sucesso!")