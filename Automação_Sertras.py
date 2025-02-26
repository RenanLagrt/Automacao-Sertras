import os
import re
import io
import sys
import time
import psutil 
import subprocess
import pytesseract
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from selenium import webdriver
from itertools import zip_longest
from pdf2image import convert_from_path
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC

# Configurações globais
poppler_path = os.path.join(os.path.expanduser("~"),"Downloads","Release-22.04.0-0","poppler-22.04.0","Library","bin")
tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = tesseract_path

def initialize_driver():
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()))

def login_sertras(driver, email, senha):
    driver.get("https://gestaodeterceiros.sertras.com/escolha-um-contrato")
    driver.maximize_window()

    campo_email = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="edtLoginInfo"]')))
    campo_email.send_keys(email)

    campo_senha = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="edtLoginSenha"]')))
    campo_senha.send_keys(senha)

    botão_enter = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="btnLogin"]/div[2]')))
    botão_enter.click()

    fechar_janela = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="notificationPopup"]/div/div/div[1]/button/span')))
    fechar_janela.click()

def tratar_tabela(Relatorio_Sertras):
    tabela = pd.read_excel(Relatorio_Sertras)

    tabela = tabela.drop(["COMENTÁRIO ANALISTA", "PRAZO SLA"], axis=1)

    Status = ["Pendente", "Pendente Correção", "Vencido"]
    tabela = tabela[tabela["STATUS"].isin(Status)]

    return tabela

def interacao_interface_recursos(driver):
    botão_recursos = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="sidebar-menu"]/div/ul/li[8]/a/span[1]')))
    driver.execute_script("arguments[0].scrollIntoView();", botão_recursos)
    botão_recursos.click()

    botão_recursos_pessoas = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="sidebar-menu"]/div/ul/li[8]/ul/li[1]/a')))
    driver.execute_script("arguments[0].scrollIntoView();", botão_recursos_pessoas)
    botão_recursos_pessoas.click()

def interacao_interface_envio(driver, nome):
    campo_nome = driver.find_element(By.XPATH, '//*[@id="filtro_nome"]')  
    campo_nome.clear()
    campo_nome.send_keys(nome)

    botão_filtrar_nome = driver.find_element(By.XPATH, '//*[@id="dashboard-v1"]/div[4]/div/div/div[2]/form/div[6]/button[1]')
    botão_filtrar_nome.click()

    botão_eventos = driver.find_element(By.XPATH, '//*[@id="data-tables2"]/tbody/tr/td[9]/a')
    botão_eventos.click()

    abas = driver.window_handles
    driver.switch_to.window(abas[-1])

    # Bloco para garantir o envio de documentos de funcionários demitidos, visto que o xpath do botão de documentação desses são alterados
    try:   
        botão_documentação = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="data-tables2"]/tbody/tr[4]/td[4]/ul/li/a')))

    except TimeoutException:
        try:
            botão_documentação = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="data-tables2"]/tbody/tr[5]/td[4]/ul/li/a')))

        except TimeoutException:
            print(f"Nenhum botão de documentação encontrado para {nome}")
            st.error(f"Nenhum botão de documentação encontrado para {nome}")
            return

    driver.execute_script("arguments[0].scrollIntoView();", botão_documentação)
    botão_documentação.click()

    abas = driver.window_handles
    driver.switch_to.window(abas[-1])

def extrair_texto_ocr(caminho_arquivo, poppler_path):
    paginas_imagem = convert_from_path(caminho_arquivo, poppler_path=poppler_path)
    return " ".join(pytesseract.image_to_string(pagina) for pagina in paginas_imagem)

def extrair_datas(texto, padrao):
    return re.findall(padrao, texto)

def calcular_vencimento(data_str, anos=1):
    data_obj = datetime.strptime(data_str, "%d/%m/%Y")
    return (data_obj.replace(year=data_obj.year + anos)).strftime("%d/%m/%Y")

def ler_aso(caminho_arquivo, poppler_path):
    texto = extrair_texto_ocr(caminho_arquivo, poppler_path)
    datas = extrair_datas(texto, r'\b\d{2}/\d{2}/\d{4}\b')
    if len(datas) > 1:
        return datas[1], calcular_vencimento(datas[1])  
    return None, None  

def ler_epi(caminho_arquivo, poppler_path):
    texto = extrair_texto_ocr(caminho_arquivo, poppler_path)
    datas = extrair_datas(texto, r'\b\d{2}/\d{2}/\d{2,4}\b')
    if datas:
        data = datas[-1]
        if len(data.split('/')[2]) == 2:
            data = data[:6] + '20' + data[6:]
        return data, calcular_vencimento(data)
    return None, None

def ler_Nrs(caminho_arquivo, poppler_path, documento):
    texto = extrair_texto_ocr(caminho_arquivo, poppler_path)
    datas = extrair_datas(texto, r'(\d{1,2}/\d{1,2}/\d{4})|(\d{1,2} de [a-zà-ú]+ de \d{4})')

    meses = {'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04',
             'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08',
             'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12'}
    
    for data in datas:
        data = data[0] or data[1]
        if 'de' in data:
            partes = data.split(' de ')
            if len(partes) == 3:
                data = f"{partes[0].zfill(2)}/{meses.get(partes[1])}/{partes[2]}"
        
        anos = 2 if documento in ["NR35", "NR10"] else 1
        return data, calcular_vencimento(data, anos)  

    return None, None

def extrair_vencimento(caminho_arquivo, poppler_path, documento):
    if documento == "ASO":
        return ler_aso(caminho_arquivo, poppler_path)
    elif documento == "EPI":
        return ler_epi(caminho_arquivo, poppler_path)
    else:
        return ler_Nrs(caminho_arquivo, poppler_path, documento)

def processar_documento(documento, nome, arquivo,caminho_base, poppler_path):
    caminho_arquivo = os.path.join(os.path.expanduser("~"), caminho_base, nome, f"{arquivo}.pdf")
    if not os.path.exists(caminho_arquivo):
        return None, None, None
    
    data_extraida, data_vencimento = extrair_vencimento(caminho_arquivo, poppler_path, documento)
    
    return caminho_arquivo, data_extraida, data_vencimento

def obter_data_modificacao(caminho_arquivo):
    return datetime.fromtimestamp(os.path.getmtime(caminho_arquivo)) if os.path.exists(caminho_arquivo) else None

def verificar_atualizacao(status, data_analise, data_envio, caminho_arquivo):
    data_modificacao = obter_data_modificacao(caminho_arquivo)

    if status == "Pendente Correção":
        data_analise = datetime.strptime(data_analise, "%d/%m/%Y")
        return data_modificacao >= data_analise
    
    else:
        data_envio = datetime.strptime(data_envio, "%d/%m/%Y %H:%M")
        data_envio = data_envio.replace(hour=0, minute=0, second=0, microsecond=0)
        return data_modificacao >= data_envio

def enviar_documento(driver, arquivo, documento, caminho_arquivo, data_vencimento, mapeamento_documentos, mapeamento_datas, vencimentos_enviados, documentos_enviados):
    if documento in mapeamento_datas:
        xpath_data = mapeamento_datas[documento]
        campo_data = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath_data)))
        driver.execute_script("arguments[0].scrollIntoView();", campo_data)
        campo_data.clear()
        campo_data.send_keys(data_vencimento)
        vencimentos_enviados.append(data_vencimento)

    if documento in mapeamento_documentos:
        xpath_documento = mapeamento_documentos[documento]
        botao_upload = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath_documento)))
        driver.execute_script("arguments[0].scrollIntoView();", botao_upload)
        botao_upload.send_keys(caminho_arquivo)
        documentos_enviados.append(arquivo)

    time.sleep(1)

    botao_envio = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnFuncaoRequisitoValores"]')))
    driver.execute_script("arguments[0].scrollIntoView();", botao_envio)
    #botao_envio.click()

    time.sleep(2)
    abas = driver.window_handles
    driver.switch_to.window(abas[-1])
    driver.close()
    driver.switch_to.window(abas[-2])
    driver.close()
    driver.switch_to.window(abas[0])

def run_automation(email, senha, Relatorio_Sertras, documentos_rh, documentos_QSMS, diretorio_base_rh, diretorio_base_qsms, mapeamento_para_documentos, mapeamento_para_datas):
    driver = initialize_driver()
    login_sertras(driver, email, senha)
    tabela = tratar_tabela(Relatorio_Sertras)
    interacao_interface_recursos(driver)

    documentos_enviados = []
    erro_envio = []
    documentos_não_encontrados = []
    documentos_encontrados = []
    documentos_atualizados = []
    documentos_nao_atualizados = []
    datas_extraidas = []
    vencimentos_projetados = []
    vencimentos_enviados = []


    NRs = ["NR10", "NR11", "NR12", "NR33", "NR35"]
    
    for nome, grupo in tabela.groupby("NOME"):
        for _, linha in grupo.iterrows():
            status, documento, funcao = linha["STATUS"], linha["DOCUMENTO"], linha["FUNÇÃO"]
            caminho_base = diretorio_base_rh if documento in documentos_rh else diretorio_base_qsms
            arquivo = f"{documento} - {nome}"
            
            caminho_arquivo, data_extraida, data_vencimento = processar_documento(documento, nome, arquivo, caminho_base, poppler_path)

            if not caminho_arquivo:
                documentos_não_encontrados.append(arquivo)
                continue

            documentos_encontrados.append(arquivo)

            if status in ["Pendente Correção","Vencido"]:
                if not verificar_atualizacao(status, linha["DATA ANÁLISE"], linha["DATA ENVIO"], caminho_arquivo):
                    documentos_nao_atualizados.append(arquivo)
                    continue
                
                else:
                    documentos_atualizados.append(arquivo)
            
            if not data_vencimento:
                erro_envio.append(arquivo)
                continue

            if isinstance(data_vencimento, (list, tuple)):
                data_vencimento = data_vencimento[0] if data_vencimento else None

            try:
                data_vencimento = datetime.strptime(data_vencimento, "%d/%m/%Y")

            except (ValueError, TypeError):
                print(f"Erro ao converter a data do documento {arquivo}: Pulando para o próximo.")
                erro_envio.append(arquivo)
                continue

            if status == "Pendente Correção":
                data_vencimento += timedelta(days=1) 

            data_vencimento = data_vencimento.strftime('%d/%m/%Y')
            datas_extraidas.append(data_extraida)
            vencimentos_projetados.append(data_vencimento)

            if arquivo in documentos_encontrados:
                interacao_interface_envio(driver, nome)
                enviar_documento(driver, arquivo, documento, caminho_arquivo, data_vencimento, 
                                        mapeamento_para_documentos.get(funcao, mapeamento_para_documentos["OUTRAS"]), 
                                        mapeamento_para_datas.get(funcao, mapeamento_para_datas["OUTRAS"]), 
                                        vencimentos_enviados, documentos_enviados)
  
    driver.quit()

    return tabela, documentos_não_encontrados, documentos_encontrados, documentos_enviados, datas_extraidas, vencimentos_projetados, vencimentos_enviados, erro_envio, documentos_atualizados, documentos_nao_atualizados  





load_dotenv()

email = os.getenv("EMAIL")
senha = os.getenv("SENHA")

data_atual = datetime.now().strftime("%d-%m-%Y")
Relatorio_Sertras = f"RELATÓRIO_SERTRAS 20-02-2025.xlsx"

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




st.set_page_config(layout="wide")

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
                tabela, documentos_não_encontrados, documentos_encontrados, documentos_enviados, datas_extraidas, vencimentos_projetados, vencimentos_enviados, erro_envio, documentos_atualizados, documentos_nao_atualizados = run_automation(
                    email, senha, Relatorio_Sertras, documentos_rh, documentos_QSMS, diretorio_base_rh, diretorio_base_qsms, mapeamento_para_documentos, mapeamento_para_datas
                )

                st.session_state["dados_processados"] = {
                    "tabela": tabela,
                    "documentos_não_encontrados": documentos_não_encontrados,
                    "documentos_encontrados": documentos_encontrados,
                    "documentos_nao_atualizados": documentos_nao_atualizados,
                    "documentos_atualizados": documentos_atualizados,
                    "erro_envio": erro_envio,
                    "datas_extraidas": datas_extraidas,
                    "vencimentos_projetados": vencimentos_projetados,
                    "documentos_enviados": documentos_enviados,
                    "vencimentos_enviados": vencimentos_enviados
                }

                st.session_state["executado"] = True
                st.rerun()

if st.session_state["dados_processados"]:
    dados = st.session_state["dados_processados"]
    
    df_sertras = dados["tabela"]
    df_documentos = pd.DataFrame(list(zip_longest(dados["documentos_não_encontrados"], dados["documentos_encontrados"], 
                                        dados["documentos_nao_atualizados"], dados["documentos_atualizados"], fillvalue="---")),

                                columns=["DOCUMENTOS NÃO ENCONTRADOS", "DOCUMENTOS ENCONTRADOS", "DOCUMENTOS NÃO ATUALIZADOS", "DOCUMENTOS ATUALIZADOS"])

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
    subprocess.Popen([sys.executable, "-m", "streamlit", "run", "Automação_Sertras.py"], shell=True)




