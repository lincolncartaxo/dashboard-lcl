import streamlit as st
import pandas as pd
import plotly.express as px
from PIL import Image
import os

# Carregar os dados a partir do arquivo Excel
file_path = "dados.xlsx"
if os.path.exists(file_path):
    try:
        # Carregar os dados
        df = pd.read_excel(file_path, engine='openpyxl', decimal=",")
        df["Data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y")

        # Criar coluna 'Mes'
        df["Mes"] = df["Data"].dt.strftime("%Y-%m")

        # Configuração geral da página
        st.set_page_config(layout="wide")

        # Barra lateral
        st.sidebar.header('Análise das Finanças - LCL Projetos')

        meses = st.sidebar.multiselect("Mês", df["Mes"].unique(), default=df["Mes"].unique())
        projetos = st.sidebar.multiselect("Projeto", df["Projeto"].unique(), default=df["Projeto"].unique())
        tipos = st.sidebar.multiselect("Tipo", df["Tipo_Despesa"].unique(), default=df["Tipo_Despesa"].unique())

        # Dashboard
        st.markdown('<style>div.block-container{padding-top:3rem;}</style>', unsafe_allow_html=True)
        logo = Image.open('logo_bw.png')

        col1, col2 = st.columns([0.1, 0.9])

        with col1:
            st.image(logo, width=100)

        html_title = """
            <style>
            .title-test {
            font-weight:bold;
            padding:5px;
            border-radius:6px;
            }
            </style>
            <center><h1 class="title-test">LCL Projetos</h1></center>"""

        with col2:
            st.markdown(html_title, unsafe_allow_html=True)

        df_filtrado = df[df["Mes"].isin(meses)]

        if projetos:
            df_filtrado = df_filtrado[df_filtrado["Projeto"].isin(projetos)]

        if tipos:
            df_filtrado = df_filtrado[df_filtrado["Tipo_Despesa"].isin(tipos)]

        # Aplicar a transformação condicional no DataFrame filtrado
        df_filtrado['Valor Total'] = df_filtrado.apply(lambda row: row['Valor Total'] * -1 if row['Tipo_Despesa'] == 'CUSTO' else row['Valor Total'], axis=1)

        st.subheader("Soma Mensal do Valor Total")
        col1, col2 = st.columns(2)
        with col1:
            # Criando o gráfico de barras mensal com Plotly Express para o DataFrame filtrado
            fig_mes2 = px.bar(df_filtrado, x='Mes', y='Valor Total', color="Projeto", title='Soma Mensal do Valor Total')

            # Exibindo o gráfico
            st.plotly_chart(fig_mes2)

        with col2:
            st.dataframe(df_filtrado)

        st.subheader("Soma Mensal do Valor Total por Projeto")
        # Agrupar por data e projeto no DataFrame filtrado
        df_agrupado = df_filtrado.groupby(['Mes', 'Projeto'])['Valor Total'].sum().reset_index()
        col1, col2 = st.columns(2)
        with col1:
            # Criando o gráfico de área mensal com Plotly Express para o DataFrame agrupado
            fig_agrupado = px.bar(df_agrupado, x='Mes', y='Valor Total', color='Projeto', title='Soma Mensal do Valor Total por Projeto')

            # Exibindo o gráfico
            st.plotly_chart(fig_agrupado)

        with col2:
            # Exibir o DataFrame agrupado
            st.dataframe(df_agrupado)

        st.subheader("Dados Agrupados por Mês")
        df_agrupado_mes = df_filtrado.groupby('Mes')['Valor Total'].sum().reset_index()
        col1, col2 = st.columns(2)
        with col1:
            # Criando o gráfico de área mensal com Plotly Express para o DataFrame agrupado
            fig_agrupado_mes = px.bar(df_agrupado_mes, x='Mes', y='Valor Total', title='Soma Mensal do Valor Total')

            # Exibindo o gráfico
            st.plotly_chart(fig_agrupado_mes)

        with col2:
            st.dataframe(df_agrupado_mes)

        st.subheader("Faturamento por Projeto")
        # Filtrar o DataFrame para incluir apenas as linhas onde 'tipo_despesa' é "RECEITA"
        df_receita = df.loc[df['Tipo_Despesa'] == "RECEITA"]

        # Agrupar por 'Mes' e 'Projeto' e somar os valores de 'Valor Total'
        df_faturamento = df_receita.groupby(['Mes', 'Projeto'])['Valor Total'].sum().reset_index()
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            # Criando o gráfico de área mensal com Plotly Express para o DataFrame agrupado
            fig_faturamento = px.bar(df_faturamento, x='Mes', y='Valor Total', color='Projeto')

            # Exibindo o gráfico
            st.plotly_chart(fig_faturamento)

        with col2:
            df_faturamento_total = df_receita.groupby(['Projeto'])['Valor Total'].sum().reset_index()
            fig_faturamento_total = px.pie(df_receita, values='Valor Total', names='Projeto')
            st.plotly_chart(fig_faturamento_total)

        with col3:
            st.dataframe(df_faturamento)

        with col4:
            df_faturamento_valor_total = df_receita['Valor Total'].sum()
            st.metric(label="Total Faturado", value=f"R$ {df_faturamento_valor_total:.2f}".replace('.', ','))

        st.subheader("Fluxo de Caixa")
        # Agrupar por data e despesa no DataFrame filtrado
        df_agrupado_fluxo = df.groupby(['Mes', 'Tipo_Despesa'])['Valor Total'].sum().reset_index()
        col1, col2 = st.columns(2)
        with col1:
            color_map = {'CUSTO': 'rgb(255, 0, 0)',  # Vermelho para despesas
                         'RECEITA': 'rgb(0, 128, 0)'}  # Verde para receitas
            fig_fluxo = px.bar(df_agrupado_fluxo, x='Mes', y='Valor Total', color='Tipo_Despesa', title='Fluxo de Caixa Mensal', color_discrete_map=color_map, barmode='group')

            st.plotly_chart(fig_fluxo)

        with col2:
            st.dataframe(df_agrupado_fluxo)

        st.subheader("Total Geral")

        col1, col2 = st.columns(2)

        with col1:
            df_agrupado_projeto = df_filtrado.groupby('Projeto')['Valor Total'].sum().reset_index()
            st.dataframe(df_agrupado_projeto)
        with col2:
            df_total_geral = df_filtrado['Valor Total'].sum()
            st.subheader("Resultado")
            st.metric(label="Total de Despesas", value=f"R$ {df_total_geral:.2f}".replace('.', ','))

    except Exception as e:
        st.error(f"Erro ao ler ou processar o arquivo Excel: {e}")

else:
    st.error(f"Arquivo '{file_path}' não encontrado.")
