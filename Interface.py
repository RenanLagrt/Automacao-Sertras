import os
import re
import io
import sys
import time
import psutil 
import subprocess
import pytesseract
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from pdf2image import convert_from_path
from datetime import datetime
from dotenv import load_dotenv
from selenium import webdriver
from openpyxl import load_workbook
from itertools import zip_longest
from datetime import datetime, timedelta
from openpyxl.cell.cell import MergedCell
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from Automação_Sertras import AutomaçãoSertras, RelatórioSertras, Envio_Sertras



load_dotenv()

st.set_page_config(layout="wide")

email = os.getenv("EMAIL")
senha = os.getenv("SENHA")

data_atual = datetime.now().strftime("%d-%m-%Y") 

documentos_rh = ["IDENTIFICAÇÃO",
                "CTPS",
                "CONTRATO DE TRABALHO",
                "FICHA DE REGISTRO",
                "CNH"]
documentos_QSMS = ["ASO", 
                "EPI", 
                "NR10", 
                "NR11", 
                "NR12", 
                "NR33", 
                "NR35", 
                "CERTIFICADO DE CLASSE"]

diretorio_base_rh = os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - SERTRAS ARQUIVO PESSOAL")
diretorio_base_qsms =  os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", "Central de Arquivos - QSMS", "000 ATUAL - OBRA 186 - INHAÚMA", "Documentação Funcionários")  

mapeamento_para_documentos = {
        "ELETRICISTA DE REPARO DE REDE DE SANEAMENTO": {
            "IDENTIFICAÇÃO": '//*[@id="edtRequisito_Valor_1"]',
            "CTPS": '//*[@id="edtRequisito_Valor_2"]',
            "FICHA DE REGISTRO": '//*[@id="edtRequisito_Valor_3"]',
            "CONTRATO DE TRABALHO": '//*[@id="edtRequisito_Valor_4"]',
            "ASO": '//*[@id="edtRequisito_Valor_6"]',
            "EPI": '//*[@id="edtRequisito_Valor_8"]',
            "CERTIFICADO DE CLASSE": '//*[@id="edtRequisito_Valor_9"]',
            "NR10": '//*[@id="edtRequisito_Valor_11"]',
            "NR33": '//*[@id="edtRequisito_Valor_13"]',
            "NR35": '//*[@id="edtRequisito_Valor_15"]',
        },
        "OPERADOR DE ESCAVADEIRA": {
            "IDENTIFICAÇÃO": '//*[@id="edtRequisito_Valor_1"]',
            "CTPS": '//*[@id="edtRequisito_Valor_2"]',
            "FICHA DE REGISTRO": '//*[@id="edtRequisito_Valor_3"]',
            "CONTRATO DE TRABALHO": '//*[@id="edtRequisito_Valor_4"]',
            "ASO": '//*[@id="edtRequisito_Valor_6"]',
            "EPI": '//*[@id="edtRequisito_Valor_8"]',
            "CNH": '//*[@id="edtRequisito_Valor_10"]',
            "NR11": '//*[@id="edtRequisito_Valor_12"]',
        },
        "OUTRAS": {
            "IDENTIFICAÇÃO": '//*[@id="edtRequisito_Valor_1"]',
            "CTPS": '//*[@id="edtRequisito_Valor_2"]',
            "FICHA DE REGISTRO": '//*[@id="edtRequisito_Valor_3"]',
            "CONTRATO DE TRABALHO": '//*[@id="edtRequisito_Valor_4"]',
            "ASO": '//*[@id="edtRequisito_Valor_6"]',
            "EPI": '//*[@id="edtRequisito_Valor_8"]',
            "CERTIFICADO DE CLASSE": '//*[@id="edtRequisito_Valor_9"]',
            "NR10": '//*[@id="edtRequisito_Valor_11"]',
            "NR12": '//*[@id="edtRequisito_Valor_10"]',
            "NR33": '//*[@id="edtRequisito_Valor_12"]',
            "NR35": '//*[@id="edtRequisito_Valor_14"]',
        },
    }

mapeamento_para_datas = {
                "ELETRICISTA DE REPARO DE REDE DE SANEAMENTO" : {
                "IDENTIFICAÇÃO" : '//*[@id="edtRequisito_Valor_0"]',
                "ASO" : '//*[@id="edtRequisito_Valor_5"]',
                "EPI" : '//*[@id="edtRequisito_Valor_7"]',
                "NR10" : '//*[@id="edtRequisito_Valor_10"]',
                "NR33" : '//*[@id="edtRequisito_Valor_12"]',
                "NR35" : '//*[@id="edtRequisito_Valor_14"]',
            },
            "OPERADOR DE ESCAVADEIRA" : {
                "IDENTIFICAÇÃO" : '//*[@id="edtRequisito_Valor_0"]',
                "ASO" : '//*[@id="edtRequisito_Valor_5"]',
                "EPI" : '//*[@id="edtRequisito_Valor_7"]',
                "CNH" : '//*[@id="edtRequisito_Valor_9"]',
                "NR11" : '//*[@id="edtRequisito_Valor_11"]',
            },
            "OUTRAS" : {  
                "IDENTIFICAÇÃO" : '//*[@id="edtRequisito_Valor_0"]',
                "ASO" : '//*[@id="edtRequisito_Valor_5"]',
                "EPI" : '//*[@id="edtRequisito_Valor_7"]',
                "NR11" : '//*[@id="edtRequisito_Valor_11"]',
                "NR12" : '//*[@id="edtRequisito_Valor_9"]',
                "NR33" : '//*[@id="edtRequisito_Valor_11"]',
                "NR35" : '//*[@id="edtRequisito_Valor_13"]',
    },
}

automacao = Envio_Sertras(email, senha)

if "executado" not in st.session_state:
    st.session_state["executado"] = False  

if "dados_processados" not in st.session_state:
    st.session_state["dados_processados"] = None

# Exibição do cabeçalho
logo_concrejato = os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", 
                            "Central de Arquivos - QSMS", "000 ATUAL - OBRA 201 - SÃO GONÇALO", 
                            "LOGO CONCREJATO.png")

logo_consorcio = os.path.join(os.path.expanduser("~"), "CONSORCIO CONCREJATOEFFICO LOTE 1", 
                            "Central de Arquivos - QSMS", "000 ATUAL - OBRA 186 - INHAÚMA", 
                            "Logo Consórcio.jpg")

col1, col2, col3 = st.columns([1, 4, 1])  

with col1:
    if os.path.exists(logo_concrejato):
        st.image(logo_concrejato, width=220)
        
    else:
        st.warning("Logotipo não encontrado!")

with col2:
    st.markdown(
        "<h1 style='text-align: center; color: #004080; font-size: 50px;'>AUTOMAÇÃO SERTRAS</h1>", 
        unsafe_allow_html=True
    )  

with col3:
    if os.path.exists(logo_consorcio):
        st.image(logo_consorcio, width=220)
    else:
        st.warning("Logotipo não encontrado!")

# Linha Separadora 
st.markdown("<hr style='border: 1px solid #004080;'>", unsafe_allow_html=True)

placeholder_botao = st.empty()

if not st.session_state["executado"]:
    col_empty1, col_button, col_empty2 = st.columns([2, 2, 2])
    with col_button:
        # 🟢 Botão Azul Nativo do Streamlit
        if st.button("🚀 Executar Automação", key="rodar_automacao", help="Clique para executar a automação", 
                     use_container_width=True, type="primary"):  
            with st.spinner("Executando automação..."):
                tabela, documentos_não_encontrados, documentos_encontrados, documentos_enviados, datas_extraidas, vencimentos_projetados, vencimentos_enviados, erro_envio, documentos_atualizados, documentos_nao_atualizados, datas_modificacao = automacao.run_automation(documentos_rh, documentos_QSMS, diretorio_base_rh, diretorio_base_qsms, mapeamento_para_documentos, mapeamento_para_datas)

                st.session_state["dados_processados"] = {
                    "tabela": tabela,
                    "documentos_não_encontrados": documentos_não_encontrados,
                    "documentos_encontrados": documentos_encontrados,
                    "documentos_nao_atualizados": documentos_nao_atualizados,
                    "documentos_atualizados": documentos_atualizados,
                    "erro_envio": erro_envio,
                    "datas_extraidas": datas_extraidas,
                    "datas_modificacao" : datas_modificacao,
                    "vencimentos_projetados": vencimentos_projetados,
                    "documentos_enviados": documentos_enviados,
                    "vencimentos_enviados": vencimentos_enviados
                }

                st.session_state["executado"] = True
                st.rerun()

if st.session_state["dados_processados"]:
    dados = st.session_state["dados_processados"]
    
    df_sertras = dados["tabela"]

    df_documentos = pd.DataFrame(list(zip_longest(dados["documentos_não_encontrados"], dados["documentos_encontrados"], dados["datas_modificacao"],
                                        dados["documentos_nao_atualizados"], dados["documentos_atualizados"], fillvalue="---")),

                                columns=["DOCUMENTOS NÃO ENCONTRADOS", "DOCUMENTOS ENCONTRADOS", "DATAS MODIFICAÇÃO", "DOCUMENTOS NÃO ATUALIZADOS", "DOCUMENTOS ATUALIZADOS"])

    df_relatorio = pd.DataFrame(list(zip_longest(dados["erro_envio"], dados["datas_extraidas"], dados["vencimentos_projetados"], 
                                       dados["documentos_enviados"], dados["vencimentos_enviados"], fillvalue="---")),

                                columns=["DOCUMENTOS SEM DATA EXTRAÍDA","DATAS EXTRAÍDAS", "VENCIMENTOS PROJETADOS", "DOCUMENTOS ENVIADOS", "VENCIMENTOS ENVIADOS"])
    
    @st.cache_data
    def to_excel_cached(df, sheet_name):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return output.getvalue()

    excel_sertras = to_excel_cached(df_sertras, "Pendencias_Sertras")
    excel_documentos = to_excel_cached(df_documentos, "Relacao_Documentos")
    excel_relatorio = to_excel_cached(df_relatorio, "Relatorio_Execução")

    centered_style = [
        {"selector": "thead th", "props": [("background-color", "blue"), ("color", "white"), ("font-weight", "bold"), ("text-align", "center")]},
        {"selector": "tbody td", "props": [("text-align", "center")]}
    ]

    df_sertras_html = df_sertras.style.set_table_styles(centered_style).hide(axis="index").to_html()
    df_documentos_html = df_documentos.style.set_table_styles(centered_style).hide(axis="index").to_html()
    df_relatorio_html = df_relatorio.style.set_table_styles(centered_style).hide(axis="index").to_html()

    def exibir_tabela(titulo, df, arquivo_excel, nome_arquivo):
        col1, col2, col3 = st.columns([0.5, 5, 0.5])  
        with col2:
            st.markdown(f"## 📋 {titulo}")
            st.markdown(df,unsafe_allow_html=True)
        with col3:
            st.download_button(
                data=arquivo_excel,
                label="⬇️",
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.markdown("<br><br>", unsafe_allow_html=True)

    exibir_tabela("RELATÓRIO PENDÊNCIAS SERTRAS", df_sertras_html, excel_sertras, f"PENDÊNCIA_SERTRAS {data_atual}.xlsx")
    exibir_tabela("RELAÇÃO DOCUMENTOS", df_documentos_html, excel_documentos, f"RELAÇÃO_DOCUMENTOS {data_atual}.xlsx")
    exibir_tabela("RELATÓRIO EXECUÇÃO", df_relatorio_html, excel_relatorio, f"RELATÓRIO_EXECUÇÃO {data_atual}.xlsx")


streamlit_rodando = False

for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
    try:
        if proc.info['cmdline'] and any("streamlit" in cmd for cmd in proc.info['cmdline']):
            streamlit_rodando = True
            break  
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        pass

if not streamlit_rodando:
    subprocess.Popen([sys.executable, "-m", "streamlit", "run", "teste.py"], shell=True)
