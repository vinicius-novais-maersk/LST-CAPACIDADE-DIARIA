from datetime import date, timedelta
import os
import re
import time
import unicodedata

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select


def normalizar_texto(texto):
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").lower()


def selecionar_opcao_por_trecho(elemento, trecho):
    select = Select(elemento)
    trecho_normalizado = normalizar_texto(trecho)

    for opcao in select.options:
        if trecho_normalizado in normalizar_texto(opcao.text):
            select.select_by_visible_text(opcao.text)
            return

    raise ValueError(f"Opcao nao encontrada no combo: {trecho}")


class EcargoRelatorioBot:
    relOS = "https://saovw087/e-cargo/PERMISSAOACESSOIMG.ASP?FUNCAO_ID=2318&nome_funcao=RELOSPESQ.asp"
    dinamico = "https://saovw087/e-cargo/PERMISSAOACESSOIMG.ASP?FUNCAO_ID=80712&nome_funcao=BUSCARELDINAMICOUSUARIO.ASP"
    monitor_ce = "https://saovw087/e-cargo/PERMISSAOACESSOIMG.ASP?FUNCAO_ID=2255&nome_funcao=MonitCEFCLSel.asp"
    salvar_planilha = r"C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\OPC Reports\Relatório automático autoGUI - Bruno\Capacidades Citrix"

    colunas = [
        "Nº OS", "ST", "Tipo OS", "Provedor", "Tipo Serviço", "Valor", "Qtde",
        "Booking", "Emissão", "Data Prog.", "Embarcador", "Destinatário",
        "Cliente Proposta", "Aut. de Coleta/Entrega", "Nota Fiscal", "Navio",
        "Origem/Destino", "Container", "Tipo", "Tara", "Lacre 1", "Lacre 2",
        "Lacre 3", "Faturamento", "Peso da carga (sem a tara)", "Nº Agendamento",
        "Nro. Apólice", "Status Averbação", "Número Averbação", "Data Averbação",
        "EDI", "Data EDI", "Val. EDI Alt. Usuário", "Empresa", "Centro de Custo",
        "Diferenciador Proposta", "Cliente Proposta"
    ]

    def __init__(self, user=r"VINISILV", password="Maersk@2027"):
        self.user = user
        self.password = password
        self.driver = None
        self.relatorio = pd.DataFrame()

    def setup_driver(self):
        from urllib.parse import quote

        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--log-level=1")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)
        chrome_options.add_experimental_option(
            "prefs",
            {
                "credentials_enable_service": False,
                "profile.password_manager_enabled": False,
            },
        )

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.driver.set_page_load_timeout(120)
        self.driver.set_script_timeout(120)
        self.driver.maximize_window()

        user_encoded = quote(self.user, safe="")
        pass_encoded = quote(self.password, safe="")
        self.driver.get(f"https://{user_encoded}:{pass_encoded}@saovw087/e-cargo/novoindex.asp")
        time.sleep(3)

        try:
            botao = self.driver.find_element(By.NAME, "Login")
            if botao.is_displayed():
                botao.click()
                print("Logon clicado")
                time.sleep(3)
        except Exception:
            pass

    def transforma_datas(self):
        hoje = date.today()
        return [
            [hoje, hoje],
            [hoje + timedelta(days=1), hoje + timedelta(days=1)],
            [hoje + timedelta(days=2), hoje + timedelta(days=2)],
            [hoje + timedelta(days=3), hoje + timedelta(days=3)],
            [hoje + timedelta(days=4), hoje + timedelta(days=16)],
        ]

    def executar(self):
        print("Iniciando extração de relatórios...")
        inicio = time.time()

        try:
            self.setup_driver()

            aux = RelatorioAuxiliar(self.driver, self.dinamico, self.monitor_ce)
            self.driver = aux.executar()

            print("Extraindo Relatórios: OS Emitidas")
            listas = self.transforma_datas()
            extractor = RelatorioExtractor(self.driver, self.relOS, self.colunas)

            for indice, datas in enumerate(listas, start=1):
                print(f"Extraindo parte {indice}/5...")
                relatorio_parcial = extractor.extrair_relatorio(datas)
                self.relatorio = (
                    pd.concat([self.relatorio, relatorio_parcial])
                    if not self.relatorio.empty
                    else relatorio_parcial
                )

            self.exportar_excel()
        finally:
            if self.driver is not None:
                try:
                    self.driver.quit()
                except Exception:
                    pass

        duracao = time.time() - inicio
        print(f"Tempo total: {int(duracao // 60)} min e {int(duracao % 60)} s.")

    def exportar_excel(self):
        print("Exportando relatório para Excel...")

        temp_path = os.path.join(os.environ["TEMP"], "ROE_EXP_py.xlsx")
        self.relatorio.to_excel(temp_path, index=False)
        time.sleep(5)

        destino = os.path.join(self.salvar_planilha, "ROE_EXP_py.xlsx")
        os.replace(temp_path, destino)

        print("Planilha salva com sucesso!")


class RelatorioExtractor:
    def __init__(self, driver, url, colunas):
        self.driver = driver
        self.url = url
        self.colunas = colunas

    @staticmethod
    def formatar_datas(datas):
        return [data.strftime("%d%m%Y") for data in datas]

    def extrair_relatorio(self, datas):
        datas = self.formatar_datas(datas)
        self.driver.get(self.url)

        campo_data_ini = self.driver.find_element(By.ID, "txtColuna8")
        campo_data_ini.click()
        campo_data_ini.send_keys(datas[0])

        campo_data_fim = self.driver.find_element(By.ID, "txtColunaFim8")
        campo_data_fim.click()
        campo_data_fim.send_keys(datas[1])

        self.driver.find_element(By.XPATH, "//tr[16]/td[2]/input[2]").click()

        campo_tipo_serv = self.driver.find_element(By.NAME, "txtColuna10")
        campo_tipo_serv.click()
        selecionar_opcao_por_trecho(campo_tipo_serv, "Transporte Rodoviário")

        self.driver.find_element(value="btnPesquisar").click()

        tabela_pag = self.driver.find_element(By.CLASS_NAME, "TIPG")
        paginas = []

        for linha in tabela_pag.find_elements(By.CSS_SELECTOR, "tr"):
            for cell in linha.find_elements(By.TAG_NAME, "td"):
                paginas.append(cell.text)

        match_paginas = re.search(r"Página\s+\d+\s+de\s+(\d+)", paginas[0])
        if not match_paginas:
            raise ValueError(f"Nao foi possivel identificar o total de paginas: {paginas[0]}")

        total_pages = int(match_paginas.group(1))
        print("Total de Páginas: " + str(total_pages))

        dados = []
        for _ in range(total_pages):
            soup = BeautifulSoup(self.driver.page_source, "html.parser")
            linhas = soup.select("table.TITB tbody tr")

            for linha in linhas:
                cols = [col.text for col in linha.find_all("td")]
                if cols and cols[0][:1].isdigit():
                    dados.append(cols)

            self.driver.get("https://saovw087/e-cargo/exibRELOSPESQ.ASP?TIPO_PAGINA=PROXIMA")

        df = pd.DataFrame(dados, columns=self.colunas)
        return df[df["Nº OS"] != 0]


class RelatorioAuxiliar:
    def __init__(self, driver, url_dinamico, url_monitor):
        self.driver = driver
        self.url_dinamico = url_dinamico
        self.url_monitor = url_monitor
        print("Extraindo Relatórios: Dinâmico e Monitor")

    def executar(self):
        self.relatorio_dinamico()
        self.monitor_ce()
        return self.driver

    def relatorio_dinamico(self):
        hoje = date.today().strftime("%d%m%Y")

        self.driver.get(self.url_dinamico)
        time.sleep(3)
        selecionar_opcao_por_trecho(
            self.driver.find_element(By.NAME, "cboRelatorios"),
            "Status de Averba",
        )
        self.driver.find_element(By.XPATH, "//input[@value=' Pesquisar ']").click()
        time.sleep(3)
        data_dinamico = self.driver.find_element(By.ID, "txt0,6")
        data_dinamico.click()
        data_dinamico.send_keys(hoje)
        self.driver.find_element(By.XPATH, "//table[3]/tbody/tr/td[2]/input").click()
        self.driver.find_element(By.XPATH, "//input[@value=' Pesquisar ']").click()
        time.sleep(3)

    def monitor_ce(self):
        hoje = date.today().strftime("%d%m%Y")
        fim = (date.today() + timedelta(days=15)).strftime("%d%m%Y")

        self.driver.get(self.url_monitor)
        time.sleep(3)
        data_ini = self.driver.find_element(By.ID, "txtdata_agend_ini")
        data_ini.click()
        data_ini.send_keys(hoje)

        data_fim = self.driver.find_element(By.ID, "txtdata_agend_fim")
        data_fim.click()
        data_fim.send_keys(fim)
        selecionar_opcao_por_trecho(
            self.driver.find_element(By.NAME, "txtTab_Tipo_Relatorio_id"),
            "Receber Planilha Excel por e-mail",
        )
        self.driver.find_element(By.ID, "btnPesquisar").click()
        time.sleep(3)


if __name__ == "__main__":
    bot = EcargoRelatorioBot()
    bot.executar()
