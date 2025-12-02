import streamlit as st
import pandas as pd
import ita_calc
import io

st.set_page_config(page_title="Calculadora ITA", layout="wide")

st.title("Calculadora de Índice de Trajetória Acadêmica (ITA)")

st.markdown("""
Esta aplicação calcula o ITA com base nas planilhas fornecidas.
Por favor, insira os links públicos (ou acessíveis) das planilhas do Google Sheets.
Certifique-se de que os links terminam com `export?format=xlsx` ou são links diretos para download.
""")

# Default URLs from the notebook for convenience
default_main = "https://docs.google.com/spreadsheets/d/1cpXhYwbhTlIexWjTprVITGAJA8jXMU83/export?format=xlsx"
default_criteria = "https://docs.google.com/spreadsheets/d/12Qjx_6-2Bed0cSXH57H3d2xrL-UP-JArwsEu5ys5WiQ/export?format=xlsx"
default_form = "/export?format=xlsx"

url_main = st.text_input("URL da Planilha rendimento/vulnerabilidade (Coag)", value=default_main)
url_criteria = st.text_input("URL da Planilha de atendimento (Social/Psicologia/Pedagogia)", value=default_criteria)
url_form = st.text_input("URL da Planilha de Formulário dos estudantes", value=default_form)

if st.button("Calcular ITA"):
    if not url_main or not url_criteria or not url_form:
        st.error("Por favor, preencha todos os campos de URL.")
    else:
        with st.spinner("Processando planilhas e calculando ITA..."):
            try:
                df_ITA = ita_calc.calculate_ita(url_main, url_criteria, url_form)
                
                st.success("Cálculo concluído com sucesso!")
                  # Excel Download
                buffer = io.BytesIO()
                buffer_ita = io.BytesIO()
                #with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                #    df_result.to_excel(writer, index=False, sheet_name='ITA Calculado')
                with pd.ExcelWriter(buffer_ita, engine='xlsxwriter') as writer:
                    df_ITA.to_excel(writer, index=False, sheet_name='ITA')                
               
                st.download_button(
                    label="Baixar Planilha ITA (Excel)",
                    data=buffer_ita.getvalue(),
                    file_name="ita_sem_form.xlsx",
                    mime="application/vnd.ms-excel"
                )
                # --- Dashboard Section ---
                st.markdown("---")
                st.header("Dashboard de Análise")

                # Sidebar Filters
                st.sidebar.header("Filtros")
                
                # Filter by Course
                cursos = df_ITA["curso"].unique().tolist() if "curso" in df_ITA.columns else []
                selected_cursos = st.sidebar.multiselect("Filtrar por Curso", options=cursos, default=cursos)
                
                # Filter by Risk Classification
                classificacoes = df_ITA["classificacao_ita"].unique().tolist() if "classificacao_ita" in df_ITA.columns else []
                selected_classificacoes = st.sidebar.multiselect("Filtrar por Classificação de Risco", options=classificacoes, default=classificacoes)

                # Filter by Income Class (if available)
                if "classe-da-renda" in df_ITA.columns:
                    rendas = df_ITA["classe-da-renda"].astype(str).unique().tolist()
                    selected_rendas = st.sidebar.multiselect("Filtrar por Classe de Renda", options=rendas, default=rendas)
                else:
                    selected_rendas = []

                # Apply Filters
                df_filtered = df_ITA.copy()
                if selected_cursos:
                    df_filtered = df_filtered[df_filtered["curso"].isin(selected_cursos)]
                if selected_classificacoes:
                    df_filtered = df_filtered[df_filtered["classificacao_ita"].isin(selected_classificacoes)]
                if "classe-da-renda" in df_ITA.columns and selected_rendas:
                    df_filtered = df_filtered[df_filtered["classe-da-renda"].astype(str).isin(selected_rendas)]

                # Metrics Row
                col1, col2, col3, col4, col5 = st.columns(5)
                
                total_alunos = len(df_filtered)
                media_ita = df_filtered["ITA"].mean() if "ITA" in df_filtered.columns else 0
                alunos_alto_risco = len(df_filtered[df_filtered["classificacao_ita"] == "61 a 100 - risco alto"]) if "classificacao_ita" in df_filtered.columns else 0
                alunos_moderado_risco = len(df_filtered[df_filtered["classificacao_ita"] == "31 a 60 - risco moderado"]) if "classificacao_ita" in df_filtered.columns else 0
                alunos_baixo_risco = len(df_filtered[df_filtered["classificacao_ita"] == "0 a 30 - baixo risco"]) if "classificacao_ita" in df_filtered.columns else 0
                
                col1.metric("Total de Alunos", total_alunos)
                col2.metric("Média do ITA", f"{media_ita:.2f}")
                col3.metric("Alunos em Alto Risco", alunos_alto_risco)
                col4.metric("Alunos em Moderado Risco", alunos_moderado_risco)
                col5.metric("Alunos em Baixo Risco", alunos_baixo_risco)
                
                #if "IRA SEM" in df_filtered.columns:
                #      media_ira = pd.to_numeric(df_filtered["IRA SEM"], errors='coerce').mean()
                #     col4.metric("Média IRA Semestral", f"{media_ira:.2f}")

                # Charts Row 1
                st.subheader("Distribuição e Classificação")
                c1, c2 = st.columns(2)
                
                import plotly.express as px
                
                with c1:
                    if "classificacao_ita" in df_filtered.columns:
                        fig_pie = px.pie(df_filtered, names="classificacao_ita", title="Distribuição por Classificação de Risco", hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                        st.plotly_chart(fig_pie, use_container_width=True)
                
                with c2:
                    if "curso" in df_filtered.columns:
                        # Count per course
                        df_course_count = df_filtered["curso"].value_counts().reset_index()
                        df_course_count.columns = ["curso", "count"]
                        fig_bar = px.bar(df_course_count, x="curso", y="count", title="Alunos por Curso", color="curso", text="count")
                        st.plotly_chart(fig_bar, use_container_width=True)

                # Charts Row 2
                st.subheader("Análise Detalhada")
                c3, c4 = st.columns(2)
                
                with c3:
                    if "ITA" in df_filtered.columns:
                        fig_hist = px.histogram(df_filtered, x="ITA", nbins=20, title="Distribuição das Notas do ITA", color_discrete_sequence=['#636EFA'])
                        st.plotly_chart(fig_hist, use_container_width=True)

                with c4:
                    if "ITA" in df_filtered.columns and "nota_final" in df_filtered.columns:
                        fig_scatter = px.scatter(df_filtered, x="nota_final", y="ITA", color="classificacao_ita", title="Correlação: Nota Final vs ITA", hover_data=["NOME"])
                        st.plotly_chart(fig_scatter, use_container_width=True)

                st.markdown("---")
                st.subheader("Tabela de Resultados (Filtrada)")
                st.dataframe(df_filtered.head(50))
                
                # Excel Download (Filtered)
                #buffer = io.BytesIO()
                #with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                #    df_filtered.to_excel(writer, index=False, sheet_name='ITA Filtrado')
                
                #st.download_button(
                #    label="Baixar Planilha Filtrada (Excel)",
                #    data=buffer.getvalue(),
                #    file_name="ita_filtrado.xlsx",
                #    mime="application/vnd.ms-excel"
                #)
                
            except Exception as e:
                st.error(f"Ocorreu um erro durante o processamento: {e}")
                st.exception(e)
