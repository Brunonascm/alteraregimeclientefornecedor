import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Atualizador de Regime Apuração Fornecedores - Domínio", layout="wide")

# Código IBGE padrão de contingência (São Paulo - SP)
CODIGO_IBGE_PADRAO = "3550308"

# Opções oficiais exigidas pelo leiaute da Domínio
REGIME_OPTIONS = {
    "N": "N - Normal",
    "M": "M - ME",
    "E": "E - EPP",
    "O": "O - Outros",
    "S": "S - ME - Simples Nacional",
    "P": "P - EPP - Simples Nacional",
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
    """
    Busca o Municipios.txt na pasta do script de forma silenciosa.
    Mapeia: Coluna 1 (Índice 0) -> Código Interno | Coluna 5 (Índice 4) -> IBGE
    """
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
                        codigo_ibge = colunas[4].strip()  # 5ª coluna (Índice 4)
                        if id_interno and codigo_ibge:
                            municipio_map[id_interno] = codigo_ibge
        except:
            pass  # Falha silenciosa
    return municipio_map


# Inicializa o mapa de municípios (em segundo plano)
mapa_municipios = ler_municipios_silencioso()


# --- TÍTULO PRINCIPAL ---
st.title("🔄 Atualizador em Lote: Regime de Apuração (Reg. 0020 - Domínio)")
st.write(
    "Carregue o arquivo TXT gerado na Domínio, altere os regimes em lote, via Excel ou individualmente, "
    "e gere o arquivo formatado em **Registro 0020 (Pipe |)** para importar na Domínio."
)

# Upload do arquivo original do SQL
uploaded_file = st.file_uploader("1️⃣ Selecione o arquivo TXT do SQL", type=["txt"])

# --- PROCESSAMENTO DO ARQUIVO SQL ---
if uploaded_file is not None:
    # Identificador único do arquivo para controle de estado
    file_key = f"{uploaded_file.name}_{uploaded_file.size}"
    
    # Se for um arquivo novo, inicializa o Session State
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
            if idx == 0 and "codi_emp" in cols[0].lower():
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

            cnpj_cpf = get_val(11)
            razao_social = get_val(5)
            
            if not cnpj_cpf or not razao_social:
                continue

            # Tratamento de Data
            raw_date = get_val(30)
            formatted_date = ""
            if raw_date:
                raw_date_clean = raw_date.split(" ")[0]
                if "-" in raw_date_clean:
                    parts = raw_date_clean.split("-")
                    if len(parts) == 3:
                        formatted_date = f"{parts[2]}/{parts[1]}/{parts[0]}"
                else:
                    formatted_date = raw_date_clean

            # Regime Inicial
            regime_atual = get_val(25).upper()
            if regime_atual not in REGIME_OPTIONS:
                regime_atual = "N"

            # Tratativa Silenciosa do Município (Código Interno -> IBGE)
            codm_for_3 = get_val(3)
            codm_for_27 = get_val(27)
            codm_for_interno = codm_for_27 if codm_for_27 else codm_for_3
            
            chave_busca = normalizar_codigo(codm_for_interno)
            if mapa_municipios and chave_busca in mapa_municipios:
                codigo_municipio_final = mapa_municipios[chave_busca]
            else:
                codigo_municipio_final = CODIGO_IBGE_PADRAO  # Default contingência SP

            parsed_data.append({
                "codi_emp": get_val(0),
                "codi_for": get_val(1),
                "sigl_est": get_val(2),
                "codm_for_interno": codm_for_interno,
                "codi_cta": get_val(4),
                "nomr_for": razao_social,
                "nome_for": get_val(6),
                "ende_for": get_val(7),
                "nume_for": get_val(8),
                "cida_for": get_val(9),
                "cepe_for": get_val(10),
                "cgce_for": cnpj_cpf,
                "insc_for": get_val(12),
                "fone_for": get_val(13),
                "faxe_for": get_val(14),
                "agro_for": get_val(15),
                "icms_for": get_val(16),
                "tins_for": get_val(17),
                "dddf_for": get_val(18),
                "imun_for": get_val(19),
                "bair_for": get_val(20),
                "categoria_estabel_for": get_val(22),
                "iestst_for": get_val(23),
                "email_for": get_val(24),
                "regime_for": regime_atual,
                "codigo_pais": get_val(26),
                "codigo_municipio": codigo_municipio_final, # IBGE Correto salvo aqui
                "insc_suframa": get_val(28),
                "complemento_for": get_val(29),
                "CADASTRO_FOR": formatted_date,
                "CONTA_CLIENTE_FOR": get_val(31),
                "POSSUI_INTERDEPENDENCIA_FOR": get_val(33),
                "CONTRIBUINTE_CPRB_FOR": get_val(44),
                "SEQUENCIAL_PAJ_FOR": get_val(47),
            })
            
        st.session_state.df = pd.DataFrame(parsed_data)

    # Verifica se os dados estão na sessão
    if "df" in st.session_state and len(st.session_state.df) > 0:
        st.success(f"✔️ {len(st.session_state.df)} fornecedores carregados com sucesso do arquivo SQL.")

        # --- SEÇÃO EXCEL ROUNDTRIP (EXPORTAR / REIMPORTAR) ---
        st.subheader("📦 Integração com Excel")
        col_ex1, col_ex2 = st.columns(2)
        
        with col_ex1:
            st.write("Baixe a planilha para ajustar os regimes de forma massiva fora do sistema:")
            # Gera Excel em memória
            excel_buffer = io.BytesIO()
            excel_df = st.session_state.df[["codi_for", "cgce_for", "nomr_for", "regime_for"]].copy()
            excel_df.columns = ["Código Fornecedor", "CNPJ_CPF", "Razão Social", "Regime de Apuração"]
            
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                excel_df.to_excel(writer, index=False, sheet_name="Regimes")
                
            st.download_button(
                label="📥 Baixar Relação em Excel (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name="relacao_fornecedores.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col_ex2:
            st.write("Suba o Excel ajustado de volta para atualizar o sistema instantaneamente:")
            uploaded_excel = st.file_uploader("Subir Excel editado", type=["xlsx"], label_visibility="collapsed")
            
            if uploaded_excel is not None:
                try:
                    excel_imported = pd.read_excel(uploaded_excel)
                    excel_imported.columns = [str(c).strip().lower() for c in excel_imported.columns]
                    
                    # Localiza as colunas chave
                    col_id = None
                    col_regime = None
                    for c in excel_imported.columns:
                        if 'codigo' in c or 'codi_for' in c:
                            col_id = c
                        if 'regime' in c or 'regime_for' in c:
                            col_regime = c
                            
                    if col_id and col_regime:
                        # Cria o dicionário de mapeamento {id_fornecedor: regime}
                        mapa_excel = {}
                        for _, row_ex in excel_imported.iterrows():
                            id_for = normalizar_codigo(row_ex[col_id])
                            reg_val = str(row_ex[col_regime]).strip().upper()
                            if reg_val:
                                reg_val = reg_val[0] # Pega apenas a primeira letra (Ex: "S" de "S - ME")
                            if reg_val in REGIME_OPTIONS:
                                mapa_excel[id_for] = reg_val
                                
                        # Atualiza o DataFrame principal na sessão
                        atualizados = 0
                        for idx, row_main in st.session_state.df.iterrows():
                            id_main = normalizar_codigo(row_main["codi_for"])
                            if id_main in mapa_excel:
                                st.session_state.df.at[idx, "regime_for"] = mapa_excel[id_main]
                                atualizados += 1
                        
                        st.success(f"✔️ Sucesso! {atualizados} regimes atualizados a partir do Excel.")
                        st.rerun()
                    else:
                        st.error("Erro: Colunas 'Código Fornecedor' e 'Regime de Apuração' não foram localizadas na planilha.")
                except Exception as e:
                    st.error(f"Erro ao processar a planilha: {e}")

        # --- SEÇÃO ALTERAÇÃO EM LOTE (STREAMLIT) ---
        st.subheader("🛠️ Alteração Rápida em Lote")
        col_b1, col_b2 = st.columns([2, 1])
        
        with col_b1:
            regime_bulk = st.selectbox(
                "Selecione o Regime de Apuração para aplicar a TODOS:",
                options=list(REGIME_OPTIONS.keys()),
                format_func=lambda x: REGIME_OPTIONS[x]
            )
        with col_b2:
            st.write(" ")
            st.write(" ")
            if st.button("Aplicar a Todos"):
                st.session_state.df["regime_for"] = regime_bulk
                st.success("Regime alterado em todos os registros!")
                st.rerun()

        # --- SEÇÃO TABELA EDITÁVEL ---
        st.subheader("📝 Tabela de Fornecedores")
        st.write("Você pode fazer ajustes manuais na tabela abaixo caso precise:")

        # Tabela editável espelha o dataframe do session_state
        edited_df = st.data_editor(
            st.session_state.df,
            column_order=["codi_for", "cgce_for", "nomr_for", "regime_for"],
            column_config={
                "codi_for": st.column_config.TextColumn("Código Fornecedor", disabled=True),
                "cgce_for": st.column_config.TextColumn("CNPJ / CPF", disabled=True),
                "nomr_for": st.column_config.TextColumn("Nome / Razão Social", disabled=True),
                "regime_for": st.column_config.SelectboxColumn(
                    "Regime de Apuração",
                    options=list(REGIME_OPTIONS.keys()),
                    format_func=lambda x: REGIME_OPTIONS[x],
                    required=True
                )
            },
            width="stretch",
            num_rows="fixed",
            key="fornecedores_editor"
        )

        # --- SEÇÃO EXPORTAÇÃO REGISTRO 0020 ---
        st.subheader("💾 Gerar Arquivo Domínio")
        
        if st.button("Gerar Arquivo de Importação (TXT)"):
            output = io.StringIO()
            
            # Varre o dataframe editado finalizado pelo usuário na tela
            for _, row in edited_df.iterrows():
                line_fields = [
                    "0020",                                                     # 1. Identificação do Registro
                    row["cgce_for"],                                            # 2. CNPJ/CPF
                    row["nomr_for"][:150],                                      # 3. Nome/Razão Social (Máx 150)
                    row["nome_for"][:40] if row["nome_for"] else row["nomr_for"][:40], # 4. Nome Fantasia (Máx 40)
                    row["ende_for"],                                            # 5. Endereço
                    row["nume_for"],                                            # 6. Número do endereço
                    row["complemento_for"],                                     # 7. Complemento
                    row["bair_for"],                                            # 8. Bairro
                    row["codigo_municipio"],                                    # 9. Código do município (IBGE Oficial, ex: 3550308!)
                    row["sigl_est"],                                            # 10. UF
                    row["codigo_pais"],                                         # 11. Código do País
                    row["cepe_for"],                                            # 12. CEP
                    row["insc_for"],                                            # 13. Inscrição Estadual
                    row["imun_for"],                                            # 14. Inscrição Municipal
                    row["insc_suframa"],                                        # 15. Inscrição Suframa
                    row["dddf_for"],                                            # 16. DDD
                    row["fone_for"],                                            # 17. Telefone
                    row["faxe_for"],                                            # 18. FAX
                    row["CADASTRO_FOR"],                                        # 19. Data do cadastro (DD/MM/YYYY)
                    row["codi_cta"],                                            # 20. Conta contábil
                    row["CONTA_CLIENTE_FOR"],                                   # 21. Conta contábil cliente
                    row["agro_for"],                                            # 22. Agropecuário (S/N)
                    "",                                                         # 23. Natureza jurídica
                    row["regime_for"],                                          # 24. REGIME DE APURAÇÃO (Alterado!)
                    row["icms_for"],                                            # 25. Contribuinte ICMS (S/N)
                    "",                                                         # 26. Alíquota ICMS
                    row["categoria_estabel_for"],                               # 27. Categoria do estabelecimento
                    row["iestst_for"],                                          # 28. Inscrição Estadual ST
                    row["email_for"],                                           # 29. Email
                    row["POSSUI_INTERDEPENDENCIA_FOR"] if row["POSSUI_INTERDEPENDENCIA_FOR"] else "N", # 30. Interdependência
                    row["CONTRIBUINTE_CPRB_FOR"] if row["CONTRIBUINTE_CPRB_FOR"] else "N",          # 31. CPRB
                    row["SEQUENCIAL_PAJ_FOR"],                                  # 32. Processo
                    row["tins_for"]                                             # 33. Tipo de inscrição
                ]
                
                # Une com Pipe "|"
                output.write("|".join(line_fields) + "\n")
            
            # Converte em ANSI
            file_data = output.getvalue().encode("cp1252", errors="replace")
            
            st.download_button(
                label="📥 Baixar IMPORTACAO_REGISTRO_0020.txt",
                data=file_data,
                file_name="IMPORTACAO_REGISTRO_0020.txt",
                mime="text/plain"
            )
            st.success("🎉 Arquivo Registro 0020 (com códigos IBGE corretos) gerado com sucesso!")
