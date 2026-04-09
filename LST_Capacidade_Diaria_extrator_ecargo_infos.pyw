from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from datetime import date, timedelta
import os
import time
import pandas as pd
import threading

class EcargoRelatorioBot:
    relOS = "https://saovw087/e-cargo/PERMISSAOACESSOIMG.ASP?FUNCAO_ID=2318&nome_funcao=RELOSPESQ.asp"
    Dinamico = "https://saovw087/e-cargo/PERMISSAOACESSOIMG.ASP?FUNCAO_ID=80712&nome_funcao=BUSCARELDINAMICOUSUARIO.ASP"
    MonitorCE = "https://saovw087/e-cargo/PERMISSAOACESSOIMG.ASP?FUNCAO_ID=2255&nome_funcao=MonitCEFCLSel.asp"
    salvarPlan = r"C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\OPC Reports\Relatório automático autoGUI - Bruno\Capacidades Citrix"
    
    heads = [
        'Nº OS', 'ST', 'Tipo OS', 'Provedor', 'Tipo Serviço', 'Valor', 'Qtde',
        'Booking', 'Emissão', 'Data Prog.', 'Embarcador', 'Destinatário',
        'Cliente Proposta', 'Aut. de Coleta/Entrega', 'Nota Fiscal', 'Navio',
        'Origem/Destino', 'Container', 'Tipo', 'Tara', 'Lacre 1', 'Lacre 2',
        'Lacre 3', 'Faturamento', 'Peso da carga (sem a tara)', 'Nº Agendamento',
        'Nro. Apólice', 'Status Averbação', 'Número Averbação', 'Data Averbação',
        'EDI', 'Data EDI', 'Val. EDI Alt. Usuário', 'Empresa', 'Centro de Custo',
        'Diferenciador Proposta', 'Cliente Proposta'
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
        chrome_options.add_experimental_option('useAutomationExtension', False)
        prefs = {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False
        }
        chrome_options.add_experimental_option("prefs", prefs)
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.driver.maximize_window()

        # Codifica as credenciais para URL — converte '\' em '%5C', '@' em '%40', etc.
        user_encoded = quote(self.user, safe='')
        pass_encoded = quote(self.password, safe='')
        self.driver.get(f"https://{user_encoded}:{pass_encoded}@saovw087/e-cargo/novoindex.asp")
        time.sleep(3)

        # Clicar no botão Logon se a página ainda exigir
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
            [hoje + timedelta(days=4), hoje + timedelta(days=16)]
        ]

    def executar(self):
        print("Iniciando extração de relatórios...")
        inicio = time.time()

        self.setup_driver()

        aux = RelatorioAuxiliar(self.driver, self.Dinamico, self.MonitorCE)
        self.driver = aux.executar()

        print("Extraindo Relatórios: OS Emitidas")
        listas = self.transforma_datas()
        extractor = RelatorioExtractor(self.driver, self.relOS, self.heads)

        for i, datas in enumerate(listas):
            print(f"Extraindo parte {i+1}/5...")
            rel = extractor.extrair_relatório(datas)
            self.relatorio = pd.concat([self.relatorio, rel]) if not self.relatorio.empty else rel

        self.exportar_excel()
        self.driver.quit()

        duracao = time.time() - inicio
        print(f"Tempo total: {int(duracao//60)} min e {int(duracao%60)} s.")

    def exportar_excel(self):
        print("Exportando relatório para Excel...")

        # Caminho temporário local
        temp_path = os.path.join(os.environ["TEMP"], "ROE_EXP_py.xlsx")

        # Salva primeiro na pasta temporária
        self.relatorio.to_excel(temp_path, index=False)

        # Aguarda alguns segundos para garantir que o arquivo foi fechado
        time.sleep(5)

        # Move para a pasta do SharePoint
        destino = os.path.join(self.salvarPlan, 'ROE_EXP_py.xlsx')
        os.replace(temp_path, destino)

        print("Planilha salva com sucesso!")

class RelatorioExtractor:
    def __init__(self, driver, url, colunas):
        self.driver = driver
        self.url = url
        self.colunas = colunas

    def formatar_datas(self, datas):
        return [d.strftime('%d%m%Y') for d in datas]

    def extrair_relatório(self, datas):
        datas = self.formatar_datas(datas)
        self.driver.get(self.url)
        # Insere Data 1
        Campo_Data_Ini = self.driver.find_element(by=By.ID, value="txtColuna8")
        Campo_Data_Ini.click()
        Campo_Data_Ini.send_keys(datas[0])

        # Insere Data 2
        Campo_Data_Fim = self.driver.find_element(By.ID, "txtColunaFim8")
        Campo_Data_Fim.click()
        Campo_Data_Fim.send_keys(datas[1])

        # Versão Expandida
        self.driver.find_element(By.XPATH, "//tr[16]/td[2]/input[2]").click()

        # Transporte Rodoviário
        Campo_TipoServ = self.driver.find_element(By.NAME, "txtColuna10")
        Campo_TipoServ.click()
        Select( Campo_TipoServ ).select_by_visible_text(u"Transporte Rodoviário")

        # Clique Pesquisar
        self.driver.find_element(value="btnPesquisar").click()

        tabela_pag = self.driver.find_element(by=By.CLASS_NAME, value="TIPG")
        pageNb =      tabela_pag.find_elements(by=By.CSS_SELECTOR, value='tr')
        Pages = []
        
        for L in pageNb:
            for cell in L.find_elements(by=By.TAG_NAME, value='td'):
                Pages.append(cell.text)
        
        total_pages = int(Pages[0][len("Página 1 de "):len("Página 1 de ")+3])
        print( "Total de Páginas: " + str(total_pages) )
        
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
        return df[df['Nº OS'] != 0]


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
        hoje = date.today().strftime('%d%m%Y')

        self.driver.get(self.url_dinamico)
        time.sleep(3)
        Select(self.driver.find_element(By.NAME, "cboRelatorios")).select_by_visible_text("Relatório de Status de Averbaç")
        self.driver.find_element(By.XPATH, "//input[@value=' Pesquisar ']").click()
        time.sleep(3)
        DataDin = self.driver.find_element(By.ID, "txt0,6")
        DataDin.click()
        DataDin.send_keys(hoje)
        self.driver.find_element(By.XPATH, "//table[3]/tbody/tr/td[2]/input").click()
        self.driver.find_element(By.XPATH, "//input[@value=' Pesquisar ']").click()
        time.sleep(3)

    def monitor_ce(self):
        hoje = date.today().strftime('%d%m%Y')
        fim = (date.today() + timedelta(days=15)).strftime('%d%m%Y')

        self.driver.get(self.url_monitor)
        time.sleep(3)
        # Data Inicial
        DataIni_mon = self.driver.find_element(By.ID, "txtdata_agend_ini")
        DataIni_mon.click()
        DataIni_mon.send_keys(hoje)

        # Data Final
        DataFim_mon = self.driver.find_element(By.ID, "txtdata_agend_fim")
        DataFim_mon.click()
        DataFim_mon.send_keys(fim)
        Select(self.driver.find_element(By.NAME, "txtTab_Tipo_Relatorio_id")).select_by_visible_text("Receber Planilha Excel por e-mail")
        self.driver.find_element(By.ID, "btnPesquisar").click()
        time.sleep(3)


if __name__ == "__main__":
    bot = EcargoRelatorioBot()
    bot.executar()
