import streamlit as st
import pandas as pd
import datetime
import io
import plotly.express as px
import xlsxwriter

st.set_page_config(page_title="Dashboard de Monitoramento SIGA", layout="wide")

st.title("📊 Dashboard de Monitoramento SIGA")

# --- Barra Lateral para Upload de Arquivos e Filtros ---
st.sidebar.header("Configurações e Filtros")

arquivo = st.sidebar.file_uploader("📎 Envie a planilha do SIGA (.xlsx)", type=["xlsx"])

# --- Função de carregamento e pré-processamento de dados (com caching) ---
@st.cache_data # <--- Caching para otimizar o carregamento
def load_and_preprocess_data(uploaded_file):
    """
    Carrega e pré-processa os dados da planilha SIGA.
    Esta função é cacheada para melhorar a performance.
    """
    df = pd.read_excel(uploaded_file)

    # Converter para datetime do Pandas (datetime64[ns])
    df["Data de Abertura"] = pd.to_datetime(df["Data de Abertura"], errors="coerce")
    df["Última Fiscalização"] = pd.to_datetime(df["Última Fiscalização"], errors="coerce")

    colunas_excluir = ["Prioritária?", "Status", "Percentual", "Empresa Executora", "Link da OS", "Localização Google Maps"]
    df = df.drop(columns=[col for col in colunas_excluir if col in df.columns])

    fiscais = {
        "norconsultdr045@gmail.com": "Fiscal Drenagem RPA 4.5",
        "norconsultdr001@gmail.com": "Fiscal Drenagem RPA 1",
        "norconsult003@gmail.com": "Fiscal SIGA RPA 3",
        "rpa2norconsult@gmail.com": "Fiscal SIGA RPA 2",
        "norconsult004@gmail.com": "Fiscal SIGA RPA 4",
        "norconsult005@gmail.com": "Fiscal SIGA RPA 5",
        "norconsult001@gmail.com": "Fiscal SIGA RPA 1",
        "norconsult006@gmail.com": "Fiscal SIGA RPA 6",
        "norconsultdr023@gmail.com": "Fiscal Drenagem RPA 2.3",
        "norconsultdr006@gmail.com": "Fiscal Drenagem RPA 6",
    }
    df["Fiscal"] = df["Fiscal"].replace(fiscais)
    
    return df

# --- Função auxiliar para formatar e baixar o DataFrame ---
def download_excel_with_formatting(df_to_export, filename, sheet_name):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        # Se houver colunas de data/hora, formate-as para evitar problemas de exibição no Excel
        for col in df_to_export.select_dtypes(include=['datetime64[ns]']).columns:
            df_to_export[col] = df_to_export[col].dt.strftime('%Y-%m-%d') # Formato 'AAAA-MM-DD'

        df_to_export.to_excel(writer, index=False, sheet_name=sheet_name)

        # Obter o objeto workbook e worksheet do xlsxwriter
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Definir formato para centralizar o texto
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

        # Aplicar formato e ajustar largura das colunas
        for i, col in enumerate(df_to_export.columns):
            # Aplicar centralização a todas as células da coluna
            worksheet.set_column(i, i, None, center_format)

            # Ajustar largura da coluna automaticamente com base no conteúdo
            # Calcula o comprimento máximo do cabeçalho ou do dado mais longo na coluna
            max_len = max(
                len(str(col)), # Comprimento do nome da coluna
                df_to_export[col].astype(str).map(len).max() # Comprimento máximo dos dados na coluna
            ) + 2 # Adiciona um pequeno padding para melhor visualização
            worksheet.set_column(i, i, max_len)

    buffer.seek(0)
    return buffer

# --- Funções de callback para o botão de limpar filtros ---
def _clear_temporal_filters():
    """Função para resetar os filtros temporais no session_state."""
    if "df_original_available" in st.session_state and not st.session_state["df_original_available"].empty:
        df_temp = st.session_state["df_original_available"].copy()
        
        # Correção: Garante que 'Última Fiscalização' é datetime64[ns] para esta operação
        df_temp["Última Fiscalização"] = pd.to_datetime(df_temp["Última Fiscalização"], errors="coerce")
        
        valid_fiscalizacao_dates = df_temp["Última Fiscalização"].dropna()

        if not valid_fiscalizacao_dates.empty:
            reset_year = valid_fiscalizacao_dates.max().year # Agora deve ser seguro
            
            month_names_map = {
                1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 
                5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 
                9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
            }
            # .dt.month agora é seguro para usar
            reset_months_nums = sorted(list(set(d.month for d in valid_fiscalizacao_dates if pd.notna(d) and d.year == reset_year)))
            reset_month_names = [month_names_map[m] for m in reset_months_nums]
            
            st.session_state["fiscal_year_slider"] = reset_year
            st.session_state["fiscal_month_multiselect"] = reset_month_names
        else: # No valid dates in 'Última Fiscalização' even after coercion
            st.session_state["fiscal_year_slider"] = datetime.date.today().year
            st.session_state["fiscal_month_multiselect"] = []
    else: # df_original_available not in session_state or is empty
        st.session_state["fiscal_year_slider"] = datetime.date.today().year
        st.session_state["fiscal_month_multiselect"] = []
    

if arquivo:
    # Chama a função de carregamento e pré-processamento cacheada
    df_original = load_and_preprocess_data(arquivo)

    # Armazenar df_original no session_state para acesso no callback
    st.session_state["df_original_available"] = df_original

    # Cria uma cópia para aplicar os filtros
    df = df_original.copy()

    # --- Filtros ---
    st.sidebar.subheader("Filtros de Dados")

    # Filtro por tipo de serviço
    tipos_disponiveis = sorted(df["Tipo de Serviço"].dropna().unique())
    tipo_padrao = ["Buraco SIGA"] if "Buraco SIGA" in tipos_disponiveis else ([] if not tipos_disponiveis else [tipos_disponiveis[0]])
    tipos_selecionados = st.sidebar.multiselect("🛠️ Tipos de Serviço:", tipos_disponiveis, default=tipo_padrao)

    if not tipos_selecionados:
        st.warning("⚠️ Selecione ao menos um tipo de serviço para continuar.")
        st.stop()

    df = df[df["Tipo de Serviço"].isin(tipos_selecionados)]

    # Filtro por Fiscal
    fiscais_disponiveis = sorted(df["Fiscal"].dropna().unique())
    fiscais_selecionados = st.sidebar.multiselect("🧑‍💼 Fiscais:", fiscais_disponiveis, default=fiscais_disponiveis)
    if fiscais_selecionados:
        df = df[df["Fiscal"].isin(fiscais_selecionados)]
    else:
        st.warning("⚠️ Selecione ao menos um fiscal para continuar.")
        st.stop()

    # Filtro por RPA
    rpas_disponiveis = sorted(df["RPA"].dropna().unique())
    rpas_selecionadas = st.sidebar.multiselect("📍 RPAs:", rpas_disponiveis, default=rpas_disponiveis)
    if rpas_selecionadas:
        df = df[df["RPA"].isin(rpas_selecionadas)]
    else:
        st.warning("⚠️ Selecione ao menos uma RPA para continuar.")
        st.stop()

    # --- Análise Temporal da Fiscalização (Unificada) ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("🗓️ Análise Temporal da Fiscalização")

    # Obter apenas datas válidas para o filtro de ano/mês
    valid_fiscalizacao_dates_for_filter = df["Última Fiscalização"].dropna()

    selected_year = None
    
    # Definir estado inicial do slider e multiselect de meses para o st.session_state
    if "fiscal_year_slider" not in st.session_state:
        if not df_original["Última Fiscalização"].dropna().empty:
            st.session_state["fiscal_year_slider"] = df_original["Última Fiscalização"].dropna().max().year
        else:
            st.session_state["fiscal_year_slider"] = datetime.date.today().year

    if "fiscal_month_multiselect" not in st.session_state:
        st.session_state["fiscal_month_multiselect"] = []

    if not valid_fiscalizacao_dates_for_filter.empty:
        min_year = valid_fiscalizacao_dates_for_filter.min().year
        max_year = valid_fiscalizacao_dates_for_filter.max().year

        # Garante que o valor no session_state esteja dentro do intervalo válido
        if st.session_state["fiscal_year_slider"] < min_year:
            st.session_state["fiscal_year_slider"] = min_year
        elif st.session_state["fiscal_year_slider"] > max_year:
            st.session_state["fiscal_year_slider"] = max_year
            
        if min_year == max_year:
            selected_year = min_year
            st.sidebar.write(f"**Ano Selecionado:** {selected_year}")
            st.session_state["fiscal_year_slider"] = selected_year
        else:
            selected_year = st.sidebar.slider(
                "Selecione o Ano:",
                min_value=min_year,
                max_value=max_year,
                value=st.session_state["fiscal_year_slider"],
                key="fiscal_year_slider"
            )
        
        month_names = {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 
            5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 
            9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
        }
        
        months_in_selected_year = sorted(list(set(d.month for d in valid_fiscalizacao_dates_for_filter if d.year == selected_year)))
        available_month_names = [month_names[m] for m in months_in_selected_year]

        default_months_for_multiselect = [
            m for m in st.session_state["fiscal_month_multiselect"] if m in available_month_names
        ]
        if not default_months_for_multiselect and available_month_names:
            default_months_for_multiselect = available_month_names

        if available_month_names:
            selected_months_names = st.sidebar.multiselect(
                "Selecione os Meses:",
                available_month_names,
                default=default_months_for_multiselect,
                key="fiscal_month_multiselect"
            )
            selected_months_nums = [list(month_names.keys())[list(month_names.values()).index(m)] for m in selected_months_names]

            df = df[
                (df["Última Fiscalização"].dt.year == selected_year) &
                (df["Última Fiscalização"].dt.month.isin(selected_months_nums))
            ]
        else:
            st.sidebar.info(f"Nenhum mês disponível para o ano de {selected_year}.")
    else:
        st.sidebar.info("Nenhuma data de última fiscalização válida para filtrar.")

    # Calcula dias desde a última fiscalização para o df principal (para KPIs e gráficos)
    hoje = datetime.date.today()
    df["Dias desde última fiscalização"] = df["Última Fiscalização"].apply(
        lambda x: (hoje - x.date()).days if pd.notna(x) else None
    )

    # Botão para limpar filtros de data, usando o callback
    st.sidebar.button("Limpar Filtros Temporais", on_click=_clear_temporal_filters)

    # Slider para definir o limite de alerta de dias
    st.sidebar.markdown("---")
    alerta_dias_config = st.sidebar.slider(
        "Alerta: Dias sem fiscalização acima de:",
        min_value=0,
        max_value=180, 
        value=30, 
        step=5
    )

    # --- Métricas / KPIs ---
    st.subheader("📈 Métricas Principais")
    total_servicos = len(df)
    
    datas_validas_fiscalizacao = df["Última Fiscalização"].dropna()
    
    rpa_contagem = df["RPA"].value_counts()
    rpa_max = rpa_contagem.idxmax() if not rpa_contagem.empty else "-"
    rpa_min = rpa_contagem.idxmin() if not rpa_contagem.empty else "-"
    
    media_fiscalizacao = df["Dias desde última fiscalização"].mean()
    media_fiscalizacao_str = f"{int(media_fiscalizacao)} dias" if not pd.isna(media_fiscalizacao) else "N/A"

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📌 Total de Serviços Filtrados", total_servicos)
    col2.metric("📅 Última Fiscalização Mais Antiga", datas_validas_fiscalizacao.min().strftime("%d/%m/%Y") if not datas_validas_fiscalizacao.empty else "N/A")
    col3.metric("📊 Média de dias sem fiscalização", media_fiscalizacao_str)
    col4.metric("📍 RPA com mais serviços", rpa_max)

    col5, col6 = st.columns(2)
    col5.metric("📍 RPA com menos serviços", rpa_min)
    
    # --- Gráficos ---
    st.subheader("🧑‍💼 Atuação dos Fiscais")
    grafico_fiscal = df["Fiscal"].value_counts().reset_index()
    grafico_fiscal.columns = ["Fiscal", "Quantidade"]
    
    if not grafico_fiscal.empty:
        fig_fiscal = px.bar(
            grafico_fiscal, 
            x="Fiscal", 
            y="Quantidade", 
            title="Quantidade de Serviços por Fiscal",
            labels={"Fiscal": "Nome do Fiscal", "Quantidade": "Número de Serviços"},
            template="plotly_white"
        )
        st.plotly_chart(fig_fiscal, use_container_width=True)
    else:
        st.info("Não há dados de fiscal para exibir no gráfico com os filtros selecionados.")


    st.subheader("🗺️ Serviços por RPA")
    df_rpa = rpa_contagem.reset_index()
    df_rpa.columns = ["RPA", "Quantidade"]
    
    # Usar raw string r'(\d+)' para regex para evitar SyntaxWarning
    df_rpa["RPA_Num"] = df_rpa["RPA"].astype(str).str.extract(r'(\d+)').astype(int, errors='ignore')
    if pd.api.types.is_numeric_dtype(df_rpa["RPA_Num"]):
        df_rpa = df_rpa.sort_values(by="RPA_Num").drop(columns="RPA_Num")
    else:
        df_rpa = df_rpa.sort_values(by="RPA")
        df_rpa = df_rpa.drop(columns="RPA_Num")

    if not df_rpa.empty:
        fig_rpa = px.bar(
            df_rpa,
            x="RPA",
            y="Quantidade",
            title="Quantidade de Serviços por RPA",
            labels={"RPA": "Região Político Administrativa", "Quantidade": "Número de Serviços"},
            template="plotly_white"
        )
        st.plotly_chart(fig_rpa, use_container_width=True)
    else:
        st.info("Não há dados de RPA para exibir no gráfico com os filtros selecionados.")


    # Gráfico de Tendência Temporal da Última Fiscalização
    st.subheader("📅 Tendência de Última Fiscalização")
    # Crie uma cópia para evitar SettingWithCopyWarning e garanta o tipo datetime64[ns]
    df_fiscalizacao_mensal = df.dropna(subset=["Última Fiscalização"]).copy()
    
    if not df_fiscalizacao_mensal.empty:
        # AGORA É SEGURO: .dt.to_period("M") porque a coluna é datetime64[ns]
        df_fiscalizacao_mensal["Ano-Mês"] = df_fiscalizacao_mensal["Última Fiscalização"].dt.to_period("M").astype(str)
        fiscalizacao_mensal_contagem = df_fiscalizacao_mensal["Ano-Mês"].value_counts().sort_index().reset_index()
        fiscalizacao_mensal_contagem.columns = ["Ano-Mês", "Número de Fiscalizações"]

        if not fiscalizacao_mensal_contagem.empty:
            fig_tendencia = px.line(
                fiscalizacao_mensal_contagem,
                x="Ano-Mês",
                y="Número de Fiscalizações",
                title="Número de Últimas Fiscalizações por Mês",
                labels={"Ano-Mês": "Período", "Número de Fiscalizações": "Contagem"},
                markers=True,
                template="plotly_white"
            )
            st.plotly_chart(fig_tendencia, use_container_width=True)
        else:
            st.info("Não há dados de última fiscalização para exibir a tendência temporal com os filtros selecionados.")
    else:
        st.info("Não há dados de última fiscalização para exibir a tendência temporal com os filtros selecionados.")


    # --- Tabela de Alerta (com limite configurável) ---
    df_alerta_visual = df[df["Dias desde última fiscalização"].apply(
        lambda x: x > alerta_dias_config if x is not None else False
    )]

    st.subheader(f"🔴 Serviços com mais de {alerta_dias_config} dias sem fiscalização (Alerta)") 
    if not df_alerta_visual.empty:
        df_alerta_ordenado = df_alerta_visual.sort_values(by="Dias desde última fiscalização", ascending=False).reset_index(drop=True)
        df_alerta_ordenado.index += 1
        st.dataframe(df_alerta_ordenado[[
            "Id", "RPA", "Bairro", "Logradouro", "Trecho", "Tipo de Serviço",
            "Última Fiscalização", "Dias desde última fiscalização"
        ]])
    else:
        st.success(f"✅ Nenhum serviço ultrapassou o limite de {alerta_dias_config} dias para este alerta.")


    # --- Download dos dados tratados ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("💾 Exportar Dados")

    # Opção 1: Baixar Dados Completos (Tratados) - PRESENTE AQUI!
    buffer_completo = download_excel_with_formatting(
        df_original, "dados_siga_completos_tratados.xlsx", "Dados Completos Tratados"
    )
    st.sidebar.download_button(
        label="📥 Baixar Dados Completos (Tratados)",
        data=buffer_completo,
        file_name="dados_siga_completos_tratados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Baixa todos os dados da planilha original após o pré-processamento, sem considerar os filtros."
    )

    # Opção 2: Baixar Dados Filtrados (Atuais) - PRESENTE AQUI!
    buffer_filtrado = download_excel_with_formatting(
        df, "dados_siga_filtrados.xlsx", "Dados Filtrados"
    )
    st.sidebar.download_button(
        label="📄 Baixar Dados Filtrados (Atuais)",
        data=buffer_filtrado,
        file_name="dados_siga_filtrados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Baixa os dados visíveis no dashboard, considerando todos os filtros aplicados."
    )

else:
    st.info("⬆️ Por favor, envie um arquivo Excel para começar a analisar os dados do SIGA.")
