import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy.orm import declarative_base
from sqlalchemy.orm import sessionmaker
from streamlit_option_menu import option_menu
from sqlalchemy import create_engine, text
from datetime import datetime, date
from urllib.parse import quote_plus
from streamlit_option_menu import option_menu

# ==========================================
# CONFIGURAÇÕES DO BANCO
# ==========================================
BD_SERVER = st.secrets["BD_SERVER"]
BD_NAME = st.secrets["BD_NAME"]
BD_USER = st.secrets["BD_USER"]
BD_PASSWORD = st.secrets["BD_PASSWORD"]
BD_DRIVER = st.secrets["BD_DRIVER"]

driver_encoded = quote_plus(BD_DRIVER)

DATABASE_URL = (
    f"mssql+pyodbc://{BD_USER}:{BD_PASSWORD}@{BD_SERVER}/{BD_NAME}"
    f"?driver={driver_encoded}"
)

# Criar engine
engine = create_engine(
    DATABASE_URL,
    fast_executemany=True,
)

SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

#tabela de mês de numero para texto
dados6 = [['1', 'Janeiro'],
          ['2', 'Fevereiro'],
          ['3', 'Março'],
          ['4', 'Abril'],
          ['5', 'Maio'],
          ['6', 'Junho'],
          ['7', 'Julho'],
          ['8', 'Agosto'],
          ['9', 'Setembro'],
          ['10', 'Outubro'],
          ['11', 'Novembro'],
          ['12', 'Dezembro'],

]
refmes = pd.DataFrame(dados6, columns=['Me', 'Mês'])
refmes.set_index('Me', inplace=True)


# ==========================================
# STREAMLIT
# ==========================================
st.set_page_config(page_title="Valorizada", layout="wide")
st.title("Importação de Dados Valorizada")

# Carregar tabela da BD ao iniciar
#valorizada = pd.read_sql("SELECT * FROM Valorizada", engine)


# Menu lateral
with st.sidebar:
    selected = option_menu(
        menu_title="Menu",
        options=["Início", "Importação", "Extração Valorizada"],
        icons=["house", "table", "gear"],
        menu_icon="cast",
        default_index=1,
    )

if selected == "Início":
    st.write("Bem-vindo ao sistema de importação de dados.")

elif selected == "Importação":
    
    # Criar colunas
    left_column, middle_column, right_column = st.columns(3)

    # Somente a seleção de data fica na esquerda
    with left_column:
        # definir estrutura de data
        min_date = date(2020, 1, 1)
        max_date = date(2070, 12, 31)
        default_start = date(2020, 1, 1)
        default_end = date(2070, 12, 31)

        # configurar data
        dados = st.date_input(
            "Definir Data:",
            value=default_start,
            min_value=min_date,
            max_value=max_date
        )
        data = pd.DataFrame({"Data": [dados]})

    # TUDO daqui para baixo deve ficar FORA do bloco "with left_column"
    upload_file = st.file_uploader("Importar Fac. Valorizada", type="xlsx")

    if upload_file:
        st.markdown("---")
        valor = pd.read_excel(upload_file, engine="openpyxl")
        valdat = pd.concat([valor, data], axis=1)
        valdat["Data"] = valdat["Data"].ffill()

        st.subheader("Pré-visualização do Excel carregado")
        st.dataframe(valor.head())

        if st.button("Guardar Facturação"):
            try:
                # Inserir novos dados
                valdat.to_sql("Valorizada", con=engine, if_exists='append', index=False)

                st.success("Informação carregada com sucesso!")
            
            except Exception as e:
                st.error(f"Ocorreu um erro na importação: {e}")

elif selected == "Extração Valorizada":

    #carregar valorizada
    query = "SELECT * FROM Valorizada"
    valorizada = pd.read_sql(query, engine)
    valorizada['Data'] = pd.to_datetime(valorizada['Data'])
    valorizada['Ano'] = valorizada['Data'].dt.year
    valorizada['Me'] = valorizada['Data'].dt.month
    valorizada['Me'] = valorizada['Me'].astype(str)
    valorizada['Ano'] = valorizada['Ano'].astype(str)
    valorizada['Data'] = valorizada['Data'].dt.date
    valorizada2 = pd.merge(valorizada, refmes, on='Me', how='left')
    #ordenar colunas
    valorizada2 = valorizada2.loc[:,['Ano', 'Mês', 'Unidade', 'Qtd_Docs', 'Lote', 'PIN', 'Ordem', 'CIL', 'Nome', 'Localidade', 'Morada', 'Aviso_Acesso', 
                                     'Nr_Documento', 'Valor', 'Rota', 'Roteiro', 'Itinerario', 'Regiao', 'UC']]

    # ORDEM CORRETA DOS MESES
    ordem_meses = [
        "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
        "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
    ]

    left_column, right_column = st.columns(2)
    with left_column:
        
        st.subheader("Definir Periodo de Leitura")

        # Ano
        ano = st.multiselect(
            "Definir Ano: ",
            options=valorizada2['Ano'].unique()
        )

        geralit = valorizada2.query("`Ano` == @ano")

        # ORGANIZAR MESES
        meses_disponiveis = geralit['Mês'].unique().tolist()
        meses_ordenados = [m for m in ordem_meses if m in meses_disponiveis]

        # Selecionar mês
        mes = st.multiselect(
            "Definir Mês:",
            options=meses_ordenados
        )

        geralit2 = geralit.query("`Mês` == @mes")
    
    #ordenar colunas
    geralit2 = geralit2.loc[:,['Unidade', 'Qtd_Docs', 'Lote', 'PIN', 'Ordem', 'CIL', 'Nome', 'Localidade', 'Morada', 'Aviso_Acesso', 
                                     'Nr_Documento', 'Valor', 'Rota', 'Roteiro', 'Itinerario', 'Regiao', 'UC']]
    geralit2["Itinerario"] = pd.to_numeric(geralit2["Itinerario"], errors="coerce").astype("Int64")

    geralit2.set_index('Unidade', inplace=True)
    st.dataframe(geralit2, use_container_width=True)

    #opção de download dos dados em excel
    @st.cache_data
    def convert_df(df):
        #conversão do dado
        return df.to_csv(sep=';', decimal=',', index=False).encode('utf-8-sig')
    
    csv = convert_df(geralit2)

    st.download_button(
        label="Download Valorizada em CSV",
        data=csv,
        file_name='Valorizada.csv',
        mime='text/csv'
            )
    
