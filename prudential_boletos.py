import os
import time
import shutil
import pandas as pd
from glob import glob
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC

def zoom():
    driver.execute_script("document.body.style.zoom='60%'")
def zim():
    driver.execute_script("document.body.style.zoom='20%'")

# ─── CONFIGURAÇÕES ────────────────────────────────────────────────────────
PERFIL_AUTOMACAO = r"C:\Users\gabriel.mendes\ChromeAutomacao\Prudential"
PASTA_DOWNLOAD   = r"C:\Users\gabriel.mendes\ChromeAutomacao\Downloads"
PASTA_LOGIN      = r"C:\Users\gabriel.mendes\Assertiv Corretora e Administradora de Seguro\ASSERTIV Sharepoint - OPERAÇÕES\Login e Senha Tech\Login_Prudential.xlsx"

os.makedirs(PERFIL_AUTOMACAO, exist_ok=True)
os.makedirs(PASTA_DOWNLOAD,   exist_ok=True)

anomes = datetime.now().strftime("%m%Y")

# Ler a planilha de clientes
cliente_df = pd.read_excel(PASTA_LOGIN)
lista_clientes = cliente_df["Cliente"].dropna().unique().tolist()

print(f"Total de clientes a processar: {len(lista_clientes)}")
print("Clientes:", lista_clientes)


# ─── CHROME OPTIONS ───────────────────────────────────────────────────────
chrome_options = Options()
chrome_options.add_argument(f"--user-data-dir={PERFIL_AUTOMACAO}")
chrome_options.add_argument("--no-first-run")
chrome_options.add_argument("--no-default-browser-check")
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--start-maximized")

prefs = {
    "download.default_directory": PASTA_DOWNLOAD,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,
    "profile.default_content_settings.popups": 0
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)


def verificar_erro_e_recarregar(driver):
    try:
        # Verifica pelo ID
        if driver.find_elements(By.ID, "bodyError"):
            print("⚠ Erro identificado pelo ID bodyError")
            driver.refresh()
            print("🔄 Página atualizada. Aguardando 1 minuto...")
            time.sleep(60)
            return True

        # Verifica pelo XPATH específico
        xpath_erro = "/html/body/div[1]/div/div/div/section/div[2]/div/span/span/div/div/div/div/div/h4"
        elementos = driver.find_elements(By.XPATH, xpath_erro)

        if elementos:
            texto = elementos[0].text

            if (
                "Opsss! Isso não era para acontecer!" in texto
                or "Bad Request" in texto
                or "An error has occurred" in texto
            ):
                print("⚠ Mensagem de erro identificada na tela:")
                print(texto)
                driver.refresh()
                print("🔄 Página atualizada. Aguardando 1 minuto...")
                time.sleep(60)
                return True

        # Verificação geral no HTML da página (extra segurança)
        page_source = driver.page_source

        if (
            "Opsss! Isso não era para acontecer!" in page_source
            or "Bad Request" in page_source
            or '{"message":"An error has occurred."}' in page_source
        ):
            print("⚠ Erro identificado via page_source")
            driver.refresh()
            print("🔄 Página atualizada. Aguardando 60 segundos...")
            time.sleep(60)
            return True

        return False
    except Exception as e:
        print(f"❌ Erro ao verificar falha na página: {e}")
        return False
    time.sleep(10)
    try:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), 'Serviços')]")
        )).click()
        print("Clicou em Serviços")
    except Exception as e:
        print(f"ERRO Serviços:")
        time.sleep(10)
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Serviços')]")
            )).click()
            print("Clicou em Serviços (2ª tentativa)")
        except Exception as e:
            print(f"ERRO Serviços (2ª tentativa): {e}")
    time.sleep(5)
    try:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), 'Consulta de Apólices')]")
        )).click()
        time.sleep(10)
        print("Clicou em Consulta de Apólices")
    except Exception as e:
        print(f"ERRO Consulta de Apólices:")
        time.sleep(10)
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Serviços')]")
            )).click()
            print("Clicou em Serviços (2ª tentativa)")
        except Exception as e:
            print(f"ERRO Serviços (2ª tentativa): {e}")
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Consulta de Apólices')]")
            )).click()
            print("Clicou em Consulta de Apólices (2ª tentativa)")
        except Exception as e:
            print(f"ERRO Consulta de Apólices (2ª tentativa): {e}")


    time.sleep(10)


    # ─── PESQUISA CLIENTE ─────────────────────────────────────────────────
    try:
        campo = wait.until(EC.presence_of_element_located(
            (By.NAME, "nm_cliente_filtro"))
        )
        campo.clear()
        campo.send_keys(nome_cliente)
        print("Cliente inserido")
    except Exception as e:
        print(f"ERRO ao inserir cliente: {e}")
        time.sleep(10)
        try:
            campo = wait.until(EC.presence_of_element_located(
                (By.NAME, "nm_cliente_filtro"))
            )
            campo.clear()
            campo.send_keys(nome_cliente)
            print("Cliente inserido")
        except Exception as e:
            print(f"ERRO ao inserir cliente: {e}")

    try:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), 'Pesquisar')]")
        )).click()
        time.sleep(10)
        print("Pesquisa executada")
    except Exception as e:
        print(f"ERRO ao pesquisar: ")
        time.sleep(10)
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Pesquisar')]")
            )).click()
            print("Pesquisa executada (2ª tentativa)")
        except Exception as e:
            print(f"ERRO ao pesquisar (2ª tentativa): {e}")
    time.sleep(20)
    zoom()
    try:   
        select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "registro_filtro"))
        )
        
        # cria o objeto Select
        select = Select(select_element)
        
        # seleciona pelo VALUE
        select.select_by_value("100")
        time.sleep(5)
        print("Clicando em 100")
    except:
        print("Não precisou de 100")




def semsub():

    sempendentes = 1
    semboleto = 1
    dados_endosso = []

    time.sleep(10)

    try:
        tabela_endosso = wait.until(
            EC.presence_of_element_located((By.ID, "tabelaLista"))
        )

        linhas_endosso = tabela_endosso.find_elements(By.TAG_NAME, "tr")

        for linha in linhas_endosso[1:]:

            colunas = linha.find_elements(By.TAG_NAME, "td")

            if len(colunas) >= 9:

                registro = {
                    "#": colunas[0].text.strip(),
                    "Nº Endosso": colunas[1].text.strip(),
                    "Emissão": colunas[2].text.strip(),
                    "Competencia": colunas[3].text.strip(),
                    "Inicio da Vigencia": colunas[4].text.strip(),
                    "Fim da Vigencia": colunas[5].text.strip(),
                    "Tipo de Endosso": colunas[6].text.strip(),
                    "Vencimento": colunas[7].text.strip(),
                    "Situacao": colunas[8].text.strip(),
                }

                dados_endosso.append(registro)

        print("Tabela de endosso lida")

    except Exception as e:

        print(f"ERRO ao ler endossos: {e}")
        print("Estabilizando página e tentando novamente...")

        time.sleep(15)

        try:

            tabela_endosso = wait.until(
                EC.presence_of_element_located((By.ID, "tabelaLista"))
            )

            linhas_endosso = tabela_endosso.find_elements(By.TAG_NAME, "tr")

            for linha in linhas_endosso[1:]:

                colunas = linha.find_elements(By.TAG_NAME, "td")

                if len(colunas) >= 9:

                    registro = {
                        "#": colunas[0].text.strip(),
                        "Nº Endosso": colunas[1].text.strip(),
                        "Emissão": colunas[2].text.strip(),
                        "Competencia": colunas[3].text.strip(),
                        "Inicio da Vigencia": colunas[4].text.strip(),
                        "Fim da Vigencia": colunas[5].text.strip(),
                        "Tipo de Endosso": colunas[6].text.strip(),
                        "Vencimento": colunas[7].text.strip(),
                        "Situacao": colunas[8].text.strip(),
                    }

                    dados_endosso.append(registro)

            print("Tabela de endosso lida (2ª tentativa)")

        except Exception as e:

            print(f"ERRO ao ler endossos (2ª tentativa): {e}")
            sempendentes = 0
            semboleto = 0
            return sempendentes, semboleto

    time.sleep(5)

    # ─── BUSCAR PENDENTE ─────────────────────────────

    pendente = next(
        (item for item in dados_endosso if item["Situacao"] == "Pendente"),
        None
    )

    if not pendente:

        print("Não tem pendentes na tabela")

        verificar_erro_e_recarregar(driver)

        sempendentes = 0
        semboleto = 0

        return sempendentes, semboleto


    indice = pendente["#"].strip()

    print(f"Clicando no endosso pendente #{indice}")

    try:

        wait.until(
            EC.element_to_be_clickable((By.ID, f"TDLINK_{indice}"))
        ).click()

    except Exception as e:

        print(f"ERRO ao clicar no endosso: {e}")

        sempendentes = 0
        semboleto = 0

        return sempendentes, semboleto


    time.sleep(15)

    # ─── PARCELAS ─────────────────────────────

    try:

        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Parcelas')]")
            )
        ).click()

        print("Clicou em Parcelas")

    except Exception as e:

        print(f"ERRO ao clicar Parcelas: {e}")

        try:

            wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(), 'Parcelas')]")
                )
            ).click()

            print("Clicou em Parcelas (2ª tentativa)")

        except Exception:

            print("Falha ao abrir Parcelas")

            sempendentes = 1
            semboleto = 0

            return sempendentes, semboleto


    time.sleep(10)

    # ─── BOTÃO BOLETO ─────────────────────────────

    try:

        wait.until(
            EC.element_to_be_clickable(
                (By.ID, "TRBTNICON_115000742")
            )
        ).click()

        print("Boleto gerado - download iniciado")

        time.sleep(15)

    except Exception:
        try:
        
            wait.until(
                EC.element_to_be_clickable(
                    (By.ID, "TRBTNICON_115000742")
                )
            ).click()
        
            print("Boleto gerado - download iniciado")
        
            time.sleep(15)
        except Exception:
            
            print("ERRO ao gerar boleto")
    
            sempendentes = 1
            semboleto = 0
    
            return sempendentes, semboleto


    # ─── SUBESTIPULANTES ─────────────────────────────

    try:

        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Subestipulantes')]")
            )
        ).click()

        print("Clicou em Subestipulantes")

    except Exception as e:

        print(f"ERRO Subestipulantes: {e}")


    verificar_erro_e_recarregar(driver)

    time.sleep(5)

    print(f"Finalizado apólice: {nome_cliente}")

    zoom()

    pasta_destino = fr"C:\Users\gabriel.mendes\Assertiv Corretora e Administradora de Seguro\ASSERTIV Sharepoint - OPERAÇÕES\Login e Senha Tech\Prudential\{anomes}\{nome_cliente}"

    os.makedirs(pasta_destino, exist_ok=True)

    arquivos_pdf = glob(os.path.join(PASTA_DOWNLOAD, "*.pdf"))

    if not arquivos_pdf:

        print("Nenhum PDF encontrado para boleto")

    else:

        ultimo_pdf = max(arquivos_pdf, key=os.path.getmtime)

        novo_nome = f"Boleto_{datetime.now().strftime('%m%Y')}.pdf"

        caminho_renomeado = os.path.join(PASTA_DOWNLOAD, novo_nome)

        os.rename(ultimo_pdf, caminho_renomeado)

        destino_final = os.path.join(pasta_destino, novo_nome)

        shutil.move(caminho_renomeado, destino_final)

        print(f"Arquivo movido: {destino_final}")

    return sempendentes, semboleto







# Limpar pasta de downloads inicialmente
for nome in os.listdir(PASTA_DOWNLOAD):
    caminho_completo = os.path.join(PASTA_DOWNLOAD, nome)
    try:
        if os.path.isfile(caminho_completo) or os.path.islink(caminho_completo):
            os.unlink(caminho_completo)
        elif os.path.isdir(caminho_completo):
            shutil.rmtree(caminho_completo)
    except Exception as e:
        print(f"-----------------------------------------")
        print(f"Erro ao apagar {caminho_completo}: {e}")
        print(f"-----------------------------------------")

print("Pasta limpa com sucesso.")









################################################ INCIO DO CÓDIGO ################################################
################################################ INCIO DO CÓDIGO ################################################
################################################ INCIO DO CÓDIGO ################################################

# ─── INICIAR DRIVER ───────────────────────────────────────────────────────
try:
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 65)
    print("Chrome iniciado com sucesso.")
except Exception as e:
    print(f"ERRO ao iniciar Chrome: {e}")
    raise
    


# ─── LOGIN ────────────────────────────────────────────────────────────────
try:
    driver.get('https://ssologin.prudential.com/app/i4pro/Login.fcc')
    print("Acessou página de login")
except Exception as e:
    print(f"ERRO ao acessar login: {e}")
    

time.sleep(5)

try:
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//*[contains(text(), 'Continuar')]")
    )).click()
    print("Login concluído")
except Exception as e:
    print(f"ERRO ao clicar em Continuar: {e}")

time.sleep(20)
print("Estabilizando pagina em 20s")
zoom()
time.sleep(20)


# ─── LOOP PRINCIPAL: CADA CLIENTE ─────────────────────────────────────────
for idx, nome_cliente in enumerate(lista_clientes, 0):
    print(f"\n{'='*70}")
    print(f"Processando cliente {idx+1}/{len(lista_clientes)}")
    print(f"{'='*70}\n")
    # idx = 1
    # nome_cliente = lista_clientes[idx]
    # print(nome_cliente)

    # Limpar pasta de downloads antes de cada cliente
    for nome in os.listdir(PASTA_DOWNLOAD):
        caminho_completo = os.path.join(PASTA_DOWNLOAD, nome)
        try:
            if os.path.isfile(caminho_completo) or os.path.islink(caminho_completo):
                os.unlink(caminho_completo)
            elif os.path.isdir(caminho_completo):
                shutil.rmtree(caminho_completo)
        except Exception as e:
            print(f"Erro ao apagar {caminho_completo}: {e}")
    print("Pasta limpa com sucesso.")
    time.sleep(1)
    zoom()
    time.sleep(10)

    # ─── NAVEGAÇÃO PARA CONSULTA DE APÓLICES ──────────────────────────────
    try:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), 'Serviços')]")
        )).click()
        print("Clicou em Serviços")
    except Exception as e:
        print(f"ERRO Serviços:")
        time.sleep(10)
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Serviços')]")
            )).click()
            print("Clicou em Serviços (2ª tentativa)")
        except Exception as e:
            print(f"ERRO Serviços (2ª tentativa): {e}")
            continue
    time.sleep(5)
    try:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), 'Consulta de Apólices')]")
        )).click()
        time.sleep(10)
        print("Clicou em Consulta de Apólices")
    except Exception as e:
        print(f"ERRO Consulta de Apólices:")
        time.sleep(10)
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Serviços')]")
            )).click()
            print("Clicou em Serviços (2ª tentativa)")
        except Exception as e:
            print(f"ERRO Serviços (2ª tentativa): {e}")
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Consulta de Apólices')]")
            )).click()
            print("Clicou em Consulta de Apólices (2ª tentativa)")
        except Exception as e:
            print(f"ERRO Consulta de Apólices (2ª tentativa): {e}")
            continue

    time.sleep(10)


    # ─── PESQUISA CLIENTE ─────────────────────────────────────────────────
    try:
        campo = wait.until(EC.presence_of_element_located(
            (By.NAME, "nm_cliente_filtro"))
        )
        campo.clear()
        campo.send_keys(nome_cliente)
        print("Cliente inserido")
    except Exception as e:
        print(f"ERRO ao inserir cliente: {e}")
        time.sleep(10)
        try:
            campo = wait.until(EC.presence_of_element_located(
                (By.NAME, "nm_cliente_filtro"))
            )
            campo.clear()
            campo.send_keys(nome_cliente)
            print("Cliente inserido")
        except Exception as e:
            print(f"ERRO ao inserir cliente: {e}")
            continue

    try:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(), 'Pesquisar')]")
        )).click()
        time.sleep(10)
        print("Pesquisa executada")
    except Exception as e:
        print(f"ERRO ao pesquisar: ")
        time.sleep(10)
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Pesquisar')]")
            )).click()
            print("Pesquisa executada (2ª tentativa)")
        except Exception as e:
            print(f"ERRO ao pesquisar (2ª tentativa): {e}")
            continue
        
        
    print("Primeira etapa finalizada, (Pesquisou o cliente)")
    
    time.sleep(15)
    
    #################################### FIM DA PRIMEIRA ETAPAAA
    #################################### FIM DA PRIMEIRA ETAPAAA
    #################################### FIM DA PRIMEIRA ETAPAAA



    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import TimeoutException
    modulo_encontrado = False
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[contains(translate(text(), 'MODULO', 'modulo'), 'modulo')]")
        ))
        modulo_encontrado = True
    except TimeoutException:
        pass
    
    if modulo_encontrado:
        print("Tem 'Modulo' na tela → aguardando 5 segundos")
        time.sleep(15)
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Endossos')]")
            )).click()
            print("Clicou em Endossos")
            verificar_erro_e_recarregar(driver)
            print("Não tem subestipulantes")
            print("Vai inciar o semsub()")
            time.sleep(5)
            zoom()
            time.sleep(1)
            sempendentes, semboleto = semsub()
            if sempendentes == 0:
                continue
            if semboleto == 0:
                continue
            try:
                time.sleep(5)
                continue
            except:
                continue

        except Exception as e:
            print(f"ERRO ao clicar Endossos: tentando novamente")
            time.sleep(10)
            try:
                wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(), 'Endossos')]")
                )).click()
                print("Clicou em Endossos")
                verificar_erro_e_recarregar(driver)
                print("Não tem subestipulantes")
                print("Vai inciar o semsub()")
                time.sleep(5)
                sempendentes, semboleto = semsub()
                if sempendentes == 0:
                    continue
                if semboleto == 0:
                    continue
                try:
                    time.sleep(5)
                    continue
                except:
                    continue
            except Exception as e:
                print(f"ERRO ao clicar Endossos (2ª tentativa): {e}")
                verificar_erro_e_recarregar(driver)
                continue
            
            
    print("NÂO TEM MODULO!!!!!!!!!!!!!!!!!!!")
    try:   
        select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "registro_filtro"))
        )
        
        # cria o objeto Select
        select = Select(select_element)
        
        # seleciona pelo VALUE
        select.select_by_value("100")
        time.sleep(5)
        print("Clicando em 100")
    except:
        print("Não precisou de 100")
        
    try:
        tabela = wait.until(
            EC.presence_of_element_located((By.ID, "tabelaLista"))
        )
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        print("Tabela principal carregada")
        time.sleep(10)
    except Exception as e:
        print(f"ERRO ao carregar tabela principal: {e}")
        linhas = []
        continue

    dados_tabela = []
    for linha in linhas:
        colunas = linha.find_elements(By.TAG_NAME, "td")
        if colunas:
            dados_tabela.append([col.text.strip() for col in colunas])

    if not dados_tabela:
        print("Nenhuma linha encontrada na tabela.")
        continue

    df_tabela = pd.DataFrame(dados_tabela)
    df_tabela[2] = (
    df_tabela[2].astype(str).str.strip() + "_" +
    df_tabela[3].astype(str)
      .str.replace(r'[./-]', '', regex=True)      # remove . / - 
      .str.replace(r'\s+', '', regex=True)         # remove espaços
      .str.strip())

# 3. (Opcional) Remove a coluna 3 original, já que não precisamos mais dela
    df_tabela = df_tabela.drop(columns=[3])
    
    df_ativos = df_tabela[df_tabela.iloc[:, 3] == 'Ativo'].copy()

    # df_tabela = df_ativos
    

    # df_indice_2 = df_tabela[[2]].copy()
    # print(df_indice_2)
    # df_indice_2.iloc[:, 0] = df_indice_2.iloc[:, 0].str.replace(" ", "_", regex=False)
    # print(df_indice_2)
    # time.sleep(1)
    # zoom()
    # time.sleep(1)
    # # ─── LOOP DAS APÓLICES ────────────────────────────────────────────────
    # links = driver.find_elements(By.XPATH, "//*[starts-with(@id, 'TDLINK_')]")
    # total = len(links)
    # print(f"Total de códigos: {total}")



    df_tabela = df_ativos.copy()  # ← aqui já são só os ativos

    df_indice_2 = df_tabela[[2]].copy()
    df_indice_2.iloc[:, 0] = df_indice_2.iloc[:, 0].str.replace(" ", "_", regex=False)
    
    time.sleep(1)
    zoom()
    time.sleep(1)
    
    # ─── LOOP DAS APÓLICES (agora total baseado nos ativos) ────────────────────
    # total = len(driver.find_elements(By.XPATH, "//*[starts-with(@id, 'TDLINK_')]"))  # ← COMENTADO: contava todos
    
    total = len(df_tabela)   # ← Agora conta apenas os ativos do DataFrame
    print(f"Total de códigos (apenas ativos): {total}")



    for i in range(1, total + 1):

        # Limpar pasta antes de cada apólice
        for nome in os.listdir(PASTA_DOWNLOAD):
            caminho_completo = os.path.join(PASTA_DOWNLOAD, nome)
            try:
                if os.path.isfile(caminho_completo) or os.path.islink(caminho_completo):
                    os.unlink(caminho_completo)
                elif os.path.isdir(caminho_completo):
                    shutil.rmtree(caminho_completo)
            except Exception as e:
                print(f"Erro ao apagar {caminho_completo}: {e}")
        print("Pasta limpa com sucesso.")
        
        
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Subestipulantes')]")
            )).click()
            print("Clicou em Subestipulantes")
            time.sleep(15)
        except Exception as e:
            print(f"ERRO ao clicar Subestipulantes:")
            time.sleep(10)
            try:
                wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(), 'Subestipulantes')]")
                )).click()
                print("Clicou em Subestipulantes (2ª tentativa)")
                time.sleep(10)
            except Exception as e:
                print(f"ERRO ao clicar Subestipulantes (2ª tentativa): {e}")
                verificar_erro_e_recarregar(driver)
                
        try:   
            select_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "registro_filtro"))
            )
            
            # cria o objeto Select
            select = Select(select_element)
            
            # seleciona pelo VALUE
            select.select_by_value("100")
            time.sleep(5)
            print("Clicando em 100")
        except:
            print("Não precisou de 100")  
            
            # Não há registros cadastrados

        time.sleep(5)
        try:
            link = wait.until(
                EC.element_to_be_clickable((By.ID, f"TDLINK_{i}"))
            )
            link.click()
            print(f"Clicou no TDLINK_{i}")
        except Exception as e:
            print(f"ERRO ao clicar TDLINK_{i}:")
            time.sleep(10)
            try:
                link = wait.until(
                    EC.element_to_be_clickable((By.ID, f"TDLINK_{i}"))
                )
                link.click()
                print(f"Clicou no TDLINK_{i} (2ª tentativa)")
            except Exception as e:
                print(f"ERRO ao clicar TDLINK_{i} (2ª tentativa): {e}")
                verificar_erro_e_recarregar(driver)
                continue

        time.sleep(15)
        zoom()
        time.sleep(20)


        # ENDOSSOS
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Endossos')]")
            )).click()
            print("Clicou em Endossos")
            verificar_erro_e_recarregar(driver)

        except Exception as e:
            print(f"ERRO ao clicar Endossos:")
            time.sleep(15)
            try:
                wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(), 'Endossos')]")
                )).click()
                print("Clicou em Endossos (2ª tentativa)")
            except Exception as e:
                print(f"ERRO ao clicar Endossos (2ª tentativa): {e}")
                verificar_erro_e_recarregar(driver)
                continue

        time.sleep(15)
        
        
        
        # LER TABELA ENDOSSO
        try:
            tabela_endosso = wait.until(
                EC.presence_of_element_located((By.ID, "tabelaLista"))
            )
            linhas_endosso = tabela_endosso.find_elements(By.TAG_NAME, "tr")
            dados_endosso = []
            for linha in linhas_endosso[1:]:
                colunas = linha.find_elements(By.TAG_NAME, "td")
                if len(colunas) >= 9:
                    registro = {
                        "#": colunas[0].text.strip(),
                        "Nº Endosso": colunas[1].text.strip(),
                        "Emissão": colunas[2].text.strip(),
                        "Competencia": colunas[3].text.strip(),
                        "Inicio da Vigencia": colunas[4].text.strip(),
                        "Fim da Vigencia": colunas[5].text.strip(),
                        "Tipo de Endosso": colunas[6].text.strip(),
                        "Vencimento": colunas[7].text.strip(),
                        "Situacao": colunas[8].text.strip(),
                    }
                    dados_endosso.append(registro)
            print("Tabela de endosso lida")
        except Exception as e:
            print(f"ERRO ao ler endossos: {e}")
            time.sleep(10)
            try:
                tabela_endosso = wait.until(
                    EC.presence_of_element_located((By.ID, "tabelaLista"))
                )
                linhas_endosso = tabela_endosso.find_elements(By.TAG_NAME, "tr")
                dados_endosso = []
                for linha in linhas_endosso[1:]:
                    colunas = linha.find_elements(By.TAG_NAME, "td")
                    if len(colunas) >= 9:
                        registro = {
                            "#": colunas[0].text.strip(),
                            "Nº Endosso": colunas[1].text.strip(),
                            "Emissão": colunas[2].text.strip(),
                            "Competencia": colunas[3].text.strip(),
                            "Inicio da Vigencia": colunas[4].text.strip(),
                            "Fim da Vigencia": colunas[5].text.strip(),
                            "Tipo de Endosso": colunas[6].text.strip(),
                            "Vencimento": colunas[7].text.strip(),
                            "Situacao": colunas[8].text.strip(),
                        }
                        dados_endosso.append(registro)
                print("Tabela de endosso lida (2ª tentativa)")
            except Exception as e:
                print(f"ERRO ao ler endossos (2ª tentativa): {e}")
                continue

        time.sleep(10)


        # ─── BUSCAR PENDENTE ──────────────────────────────────────────────
        try:
            pendente = next(
                (item for item in dados_endosso if item["Situacao"] == "Pendente"),
                None
            )

            if not pendente:
                print(f"-------------------------------------------------------------")
                print(f"Apólice {i}: Sem endosso pendente → continuando para próxima")
                print(f"-------------------------------------------------------------")
                time.sleep(2)
                print("Clicando em Subestipulantes")
                try:
                    wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//*[contains(text(), 'Subestipulantes')]")
                    )).click()
                    print("Clicou em Subestipulantes")
                    time.sleep(15)
                except Exception as e:
                    print(f"ERRO ao clicar Subestipulantes:")
                    time.sleep(10)
                    try:
                        wait.until(EC.element_to_be_clickable(
                            (By.XPATH, "//*[contains(text(), 'Subestipulantes')]")
                        )).click()
                        print("Clicou em Subestipulantes (2ª tentativa)")
                        time.sleep(10)
                    except Exception as e:
                        print(f"ERRO ao clicar Subestipulantes (2ª tentativa): {e}")
                        verificar_erro_e_recarregar(driver)
                        continue

                continue

            indice = pendente["#"].strip()
            print("Tem pendentes!")
            print(f"Apólice {i}:Clicando no endosso pendente #{indice}")

            wait.until(EC.element_to_be_clickable((By.ID, f"TDLINK_{indice}"))).click()

        except Exception as e:
            print(f"ERRO grave ao buscar/processar pendente na apólice {i}: {e}")
            continue

        time.sleep(15)


        # PARCELAS
        try:
            time.sleep(10)
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Parcelas')]")
            )).click()
            print("Clicou em Parcelas")
        except Exception as e:
            print(f"ERRO ao clicar Parcelas: {e}")
            time.sleep(10)
            try:
                wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(), 'Parcelas')]")
                )).click()
                print("Clicou em Parcelas (2ª tentativa)")
            except Exception as e:
                print(f"ERRO ao clicar Parcelas (2ª tentativa): {e}")
                verificar_erro_e_recarregar(driver)
                continue

        time.sleep(25)
        

        
        # ─── BOTÃO BOLETO ─────────────────────────────

        try:

            wait.until(
                EC.element_to_be_clickable(
                    (By.ID, "TRBTNICON_115000742")
                )
            ).click()

            print("Boleto gerado - download iniciado")

            time.sleep(15)

        except Exception:
            try:
            
                wait.until(
                    EC.element_to_be_clickable(
                        (By.ID, "TRBTNICON_115000742")
                    )
                ).click()
            
                print("Boleto gerado - download iniciado")
            
                time.sleep(15)
            except Exception:
                
                print("ERRO ao gerar boleto")
                print("Clicando em Subestipulantes")
                try:
                    wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//*[contains(text(), 'Subestipulantes')]")
                    )).click()
                    print("Clicou em Subestipulantes")
                    time.sleep(15)
                except Exception as e:
                    print(f"ERRO ao clicar Subestipulantes:")
                    time.sleep(10)
                    try:
                        wait.until(EC.element_to_be_clickable(
                            (By.XPATH, "//*[contains(text(), 'Subestipulantes')]")
                        )).click()
                        print("Clicou em Subestipulantes (2ª tentativa)")
                        time.sleep(10)
                    except Exception as e:
                        print(f"ERRO ao clicar Subestipulantes (2ª tentativa): {e}")
                        verificar_erro_e_recarregar(driver)
                        continue


        pasta_destino = fr"C:\Users\gabriel.mendes\Assertiv Corretora e Administradora de Seguro\ASSERTIV Sharepoint - OPERAÇÕES\Login e Senha Tech\Prudential\{anomes}\{nome_cliente}"
    
        os.makedirs(pasta_destino, exist_ok=True)
    
        arquivos_pdf = glob(os.path.join(PASTA_DOWNLOAD, "*.pdf"))
    
        if not arquivos_pdf:
    
            print("Nenhum PDF encontrado para boleto")
    
        else:
    
            ultimo_pdf = max(arquivos_pdf, key=os.path.getmtime)
    
            novo_nome = f"Boleto_{datetime.now().strftime('%m%Y')}.pdf"
    
            caminho_renomeado = os.path.join(PASTA_DOWNLOAD, novo_nome)
    
            os.rename(ultimo_pdf, caminho_renomeado)
    
            destino_final = os.path.join(pasta_destino, novo_nome)
    
            shutil.move(caminho_renomeado, destino_final)
    
            print(f"Arquivo movido: {destino_final}")



print("\n" + "="*80)
print("PROCESSO FINALIZADO")
print("="*80)

driver.quit()
