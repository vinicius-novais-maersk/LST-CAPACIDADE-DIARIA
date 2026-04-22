from datetime import date, timedelta
import os
import re
import time
import unicodedata

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import (
    InvalidSessionIdException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


WAIT_TIMEOUT = 45
MAX_TENTATIVAS_DRIVER = 3
MAX_TENTATIVAS_PARTE = 3
ERROS_RECUPERAVEIS = (
    InvalidSessionIdException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
    ValueError,
)


def normalizar_texto(texto):
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").lower()


class BaseEcargoPage:
    def __init__(self, driver):
        self.driver = driver

    def aguardar_documento_pronto(self, timeout=WAIT_TIMEOUT):
        WebDriverWait(self.driver, timeout, poll_frequency=1).until(
            lambda navegador: navegador.execute_script("return document.readyState") == "complete"
        )

    def navegar(self, url, locator_esperado=None, timeout=WAIT_TIMEOUT):
        self.driver.get(url)
        self.aguardar_documento_pronto(timeout=timeout)

        if locator_esperado is not None:
            self.aguardar_presenca(*locator_esperado, timeout=timeout)

    def aguardar_presenca(self, by, value, timeout=WAIT_TIMEOUT):
        return WebDriverWait(
            self.driver,
            timeout,
            poll_frequency=1,
            ignored_exceptions=(StaleElementReferenceException,),
        ).until(EC.presence_of_element_located((by, value)))

    def aguardar_visivel(self, by, value, timeout=WAIT_TIMEOUT):
        return WebDriverWait(
            self.driver,
            timeout,
            poll_frequency=1,
            ignored_exceptions=(StaleElementReferenceException,),
        ).until(EC.visibility_of_element_located((by, value)))

    def clicar(self, by, value, timeout=WAIT_TIMEOUT):
        elemento = WebDriverWait(
            self.driver,
            timeout,
            poll_frequency=1,
            ignored_exceptions=(StaleElementReferenceException,),
        ).until(EC.element_to_be_clickable((by, value)))
        elemento.click()
        return elemento

    def preencher(self, by, value, texto, timeout=WAIT_TIMEOUT):
        campo = self.aguardar_visivel(by, value, timeout=timeout)
        campo.click()

        try:
            campo.clear()
        except Exception:
            pass

        campo.send_keys(Keys.CONTROL, "a")
        campo.send_keys(Keys.BACKSPACE)
        campo.send_keys(texto)
        return campo

    def selecionar_opcao_por_trecho(self, by, value, trecho, timeout=WAIT_TIMEOUT, tentativas=3):
        trecho_normalizado = normalizar_texto(trecho)
        ultimo_erro = None

        for _ in range(tentativas):
            try:
                select = Select(self.aguardar_visivel(by, value, timeout=timeout))

                for opcao in select.options:
                    if trecho_normalizado in normalizar_texto(opcao.text):
                        select.select_by_visible_text(opcao.text)
                        return

                raise ValueError(f"Opcao nao encontrada no combo: {trecho}")
            except StaleElementReferenceException as exc:
                ultimo_erro = exc
                time.sleep(1)

        if ultimo_erro is not None:
            raise ultimo_erro

        raise ValueError(f"Opcao nao encontrada no combo: {trecho}")


class EcargoRelatorioBot:
    relOS = "https://saovw087/e-cargo/PERMISSAOACESSOIMG.ASP?FUNCAO_ID=2318&nome_funcao=RELOSPESQ.asp"
    dinamico = "https://saovw087/e-cargo/PERMISSAOACESSOIMG.ASP?FUNCAO_ID=80712&nome_funcao=BUSCARELDINAMICOUSUARIO.ASP"
    monitor_ce = "https://saovw087/e-cargo/PERMISSAOACESSOIMG.ASP?FUNCAO_ID=2255&nome_funcao=MonitCEFCLSel.asp"
    salvar_planilha = (
        "C:\\Users\\VNO024\\OneDrive - Maersk Group\\Inland Execution Brazil - OPC - OPC (also BPS)"
        "\\OPC Reports\\Relat\u00f3rio autom\u00e1tico autoGUI - Bruno\\Capacidades Citrix"
    )

    colunas = [
        "N\u00ba OS",
        "ST",
        "Tipo OS",
        "Provedor",
        "Tipo Servi\u00e7o",
        "Valor",
        "Qtde",
        "Booking",
        "Emiss\u00e3o",
        "Data Prog.",
        "Embarcador",
        "Destinat\u00e1rio",
        "Cliente Proposta",
        "Aut. de Coleta/Entrega",
        "Nota Fiscal",
        "Navio",
        "Origem/Destino",
        "Container",
        "Tipo",
        "Tara",
        "Lacre 1",
        "Lacre 2",
        "Lacre 3",
        "Faturamento",
        "Peso da carga (sem a tara)",
        "N\u00ba Agendamento",
        "Nro. Ap\u00f3lice",
        "Status Averba\u00e7\u00e3o",
        "N\u00famero Averba\u00e7\u00e3o",
        "Data Averba\u00e7\u00e3o",
        "EDI",
        "Data EDI",
        "Val. EDI Alt. Usu\u00e1rio",
        "Empresa",
        "Centro de Custo",
        "Diferenciador Proposta",
        "Cliente Proposta",
    ]

    def __init__(self, user=r"VINISILV", password="Maersk@2027"):
        self.user = user
        self.password = password
        self.driver = None
        self.relatorio = pd.DataFrame()

    def fechar_driver(self):
        if self.driver is None:
            return

        try:
            self.driver.quit()
        except Exception:
            pass
        finally:
            self.driver = None

    def setup_driver(self):
        from urllib.parse import quote

        ultimo_erro = None

        for tentativa in range(1, MAX_TENTATIVAS_DRIVER + 1):
            self.fechar_driver()

            try:
                chrome_options = Options()
                chrome_options.add_argument("--headless=new")
                chrome_options.add_argument("--disable-gpu")
                chrome_options.add_argument("--window-size=1920,1080")
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
                self.driver.set_page_load_timeout(60)
                self.driver.set_script_timeout(60)
                self.driver.set_window_size(1920, 1080)
                self.driver.execute_cdp_cmd(
                    "Page.addScriptToEvaluateOnNewDocument",
                    {
                        "source": """
                            Object.defineProperty(navigator, 'webdriver', {
                                get: () => undefined
                            });
                        """
                    },
                )

                user_encoded = quote(self.user, safe="")
                pass_encoded = quote(self.password, safe="")
                helper = BaseEcargoPage(self.driver)
                helper.navegar(
                    f"https://{user_encoded}:{pass_encoded}@saovw087/e-cargo/novoindex.asp",
                    timeout=60,
                )

                try:
                    helper.clicar(By.NAME, "Login", timeout=8)
                    print("Logon clicado")
                    helper.aguardar_documento_pronto(timeout=30)
                except TimeoutException:
                    pass

                return
            except (TimeoutException, WebDriverException) as exc:
                ultimo_erro = exc
                print(
                    f"Falha ao iniciar o navegador ({tentativa}/{MAX_TENTATIVAS_DRIVER}): "
                    f"{type(exc).__name__}: {exc}"
                )
                time.sleep(3)

        raise RuntimeError("Nao foi possivel iniciar a sessao do e-Cargo.") from ultimo_erro

    def preparar_sessao(self):
        self.setup_driver()
        aux = RelatorioAuxiliar(self.driver, self.dinamico, self.monitor_ce)
        aux.executar()

    def executar_parte_com_retry(self, extractor, indice, datas):
        ultimo_erro = None

        for tentativa in range(1, MAX_TENTATIVAS_PARTE + 1):
            try:
                return extractor.extrair_relatorio(datas)
            except ERROS_RECUPERAVEIS as exc:
                ultimo_erro = exc
                print(
                    f"Parte {indice}/5 falhou na tentativa {tentativa}/{MAX_TENTATIVAS_PARTE}: "
                    f"{type(exc).__name__}: {exc}"
                )

                if tentativa == MAX_TENTATIVAS_PARTE:
                    break

                print("Reiniciando sessao do navegador e repetindo a parte...")
                self.preparar_sessao()
                extractor.driver = self.driver
                time.sleep(2)

        raise RuntimeError(
            f"Nao foi possivel concluir a parte {indice} apos {MAX_TENTATIVAS_PARTE} tentativas."
        ) from ultimo_erro

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
        print("Iniciando extracao de relatorios...")
        inicio = time.time()

        try:
            self.preparar_sessao()
            print("Extraindo Relatorios: OS Emitidas")
            listas = self.transforma_datas()
            extractor = RelatorioExtractor(self.driver, self.relOS, self.colunas)

            for indice, datas in enumerate(listas, start=1):
                print(f"Extraindo parte {indice}/5...")
                relatorio_parcial = self.executar_parte_com_retry(extractor, indice, datas)
                self.relatorio = (
                    pd.concat([self.relatorio, relatorio_parcial])
                    if not self.relatorio.empty
                    else relatorio_parcial
                )

            self.exportar_excel()
        finally:
            self.fechar_driver()

        duracao = time.time() - inicio
        print(f"Tempo total: {int(duracao // 60)} min e {int(duracao % 60)} s.")

    def exportar_excel(self):
        print("Exportando relatorio para Excel...")

        temp_path = os.path.join(os.environ["TEMP"], "ROE_EXP_py.xlsx")
        self.relatorio.to_excel(temp_path, index=False)
        time.sleep(2)

        destino = os.path.join(self.salvar_planilha, "ROE_EXP_py.xlsx")
        os.replace(temp_path, destino)

        print("Planilha salva com sucesso!")


class RelatorioExtractor(BaseEcargoPage):
    regex_paginas = re.compile(r"pagina\s+(\d+)\s+de\s+(\d+)")
    url_proxima_pagina = "https://saovw087/e-cargo/exibRELOSPESQ.ASP?TIPO_PAGINA=PROXIMA"

    def __init__(self, driver, url, colunas):
        super().__init__(driver)
        self.driver = driver
        self.url = url
        self.colunas = colunas

    @staticmethod
    def formatar_datas(datas):
        return [data.strftime("%d%m%Y") for data in datas]

    @classmethod
    def extrair_paginacao(cls, html):
        match_paginas = cls.regex_paginas.search(normalizar_texto(html))

        if not match_paginas:
            return None

        return int(match_paginas.group(1)), int(match_paginas.group(2))

    def aguardar_resultado(self, pagina_esperada=1, timeout=WAIT_TIMEOUT):
        def condicao(driver):
            html = driver.page_source
            paginacao = self.extrair_paginacao(html)

            if paginacao is None:
                return False

            pagina_atual, total_paginas = paginacao
            if pagina_atual != pagina_esperada:
                return False

            soup = BeautifulSoup(html, "html.parser")
            if not soup.select("table.TITB tr"):
                return False

            return pagina_atual, total_paginas, soup

        return WebDriverWait(self.driver, timeout, poll_frequency=1).until(condicao)

    def extrair_linhas(self, soup):
        dados = []

        for linha in soup.select("table.TITB tr"):
            cols = [col.get_text(strip=True) for col in linha.find_all("td")]
            if not cols or not cols[0][:1].isdigit():
                continue

            if len(cols) < len(self.colunas):
                cols.extend([""] * (len(self.colunas) - len(cols)))
            elif len(cols) > len(self.colunas):
                cols = cols[: len(self.colunas)]

            dados.append(cols)

        return dados

    def extrair_relatorio(self, datas):
        datas = self.formatar_datas(datas)
        self.navegar(self.url, locator_esperado=(By.ID, "txtColuna8"))
        self.preencher(By.ID, "txtColuna8", datas[0])
        self.preencher(By.ID, "txtColunaFim8", datas[1])
        self.clicar(By.XPATH, "//tr[16]/td[2]/input[2]")
        self.selecionar_opcao_por_trecho(By.NAME, "txtColuna10", "Transporte Rodoviario")
        self.clicar(By.ID, "btnPesquisar")

        _, total_pages, soup = self.aguardar_resultado(pagina_esperada=1, timeout=60)
        print("Total de Paginas: " + str(total_pages))

        dados = self.extrair_linhas(soup)
        for pagina in range(2, total_pages + 1):
            self.navegar(self.url_proxima_pagina, timeout=60)
            _, _, soup = self.aguardar_resultado(pagina_esperada=pagina, timeout=60)
            dados.extend(self.extrair_linhas(soup))

        if not dados:
            raise ValueError("Nenhuma linha valida foi extraida do relatorio.")

        df = pd.DataFrame(dados, columns=self.colunas)
        coluna_os = self.colunas[0]
        return df[df[coluna_os].astype(str).str.strip() != "0"]


class RelatorioAuxiliar(BaseEcargoPage):
    def __init__(self, driver, url_dinamico, url_monitor):
        super().__init__(driver)
        self.driver = driver
        self.url_dinamico = url_dinamico
        self.url_monitor = url_monitor
        print("Extraindo Relatorios: Dinamico e Monitor")

    def executar(self):
        self.relatorio_dinamico()
        self.monitor_ce()
        return self.driver

    def relatorio_dinamico(self):
        hoje = date.today().strftime("%d%m%Y")

        self.navegar(self.url_dinamico, locator_esperado=(By.NAME, "cboRelatorios"))
        self.selecionar_opcao_por_trecho(By.NAME, "cboRelatorios", "Status de Averba")
        self.clicar(By.XPATH, "//input[@value=' Pesquisar ']")
        self.preencher(By.ID, "txt0,6", hoje)
        self.clicar(By.XPATH, "//table[3]/tbody/tr/td[2]/input")
        self.clicar(By.XPATH, "//input[@value=' Pesquisar ']")
        self.aguardar_documento_pronto()

    def monitor_ce(self):
        hoje = date.today().strftime("%d%m%Y")
        fim = (date.today() + timedelta(days=15)).strftime("%d%m%Y")

        self.navegar(self.url_monitor, locator_esperado=(By.ID, "txtdata_agend_ini"))
        self.preencher(By.ID, "txtdata_agend_ini", hoje)
        self.preencher(By.ID, "txtdata_agend_fim", fim)
        self.selecionar_opcao_por_trecho(
            By.NAME,
            "txtTab_Tipo_Relatorio_id",
            "Receber Planilha Excel por e-mail",
        )
        self.clicar(By.ID, "btnPesquisar")
        self.aguardar_documento_pronto()


if __name__ == "__main__":
    bot = EcargoRelatorioBot()
    bot.executar()
