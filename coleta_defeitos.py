# === IMPORTAÇÕES ===
import time
import os
import pyautogui
import gspread
import json
import sys
import traceback

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
from datetime import datetime, timedelta

# === CONFIGURAÇÕES INICIAIS ===
CAMINHO_OS_TEMP = "os_feitas.json"
load_dotenv()
USUARIO = os.getenv("USUARIO")
SENHA = os.getenv("SENHA")

# === FUNÇÕES DE CAMINHO ===
def caminho_absoluto(relativo):
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(__file__)
    return os.path.join(base, relativo)

def caminho_img(nome_arquivo):
    return caminho_absoluto(os.path.join("img", nome_arquivo))

# === GERENCIAMENTO DE ORDENS DE SERVIÇO (TEMPORÁRIO) ===
def carregar_os_processados():
    if os.path.exists(CAMINHO_OS_TEMP):
        with open(CAMINHO_OS_TEMP, 'r', encoding='utf-8') as f:
            try:
                return set(json.load(f))
            except json.JSONDecodeError:
                return set()
    return set()

def salvar_os_processados(os_numero):
    os_processadas = carregar_os_processados()
    os_processadas.add(os_numero)
    with open(CAMINHO_OS_TEMP, "w", encoding='utf-8') as f:
        json.dump(list(os_processadas), f)

def limpar_arquivo_temporario():
    if os.path.exists(CAMINHO_OS_TEMP):
        os.remove(CAMINHO_OS_TEMP)

def fechar_popup():
    time.sleep(10)
    print("❎ Fechando pop-up...")
    pyautogui.moveTo(x=415, y=105)
    time.sleep(5)
    pyautogui.click()
    print("✅ Pop-up fechado.")

def localizar_e_clicar_ver():
    print("👁️ Procurando botão 'Ver' na tela...")
    for tentativa in range(10):
        btn_pos = pyautogui.locateOnScreen(caminho_absoluto("img/ver.png"), confidence=0.7)
        if btn_pos:
            pyautogui.click(btn_pos)
            print("✅ Botão 'Ver' clicado.")
            return
        print(f"⏳ Tentativa {tentativa + 1}/10 – botão ainda não visível.")
        time.sleep(1)
    print("❌ Botão 'Ver' NÃO encontrado após múltiplas tentativas.")
    raise RuntimeError("Botão 'Ver' não encontrado após 10 segundos.")

def preencher_campo_com_pyautogui():
    print("⌨️ Preenchendo campo com 'Todos TT's'...")
    time.sleep(5)
    pyautogui.press('backspace')
    time.sleep(2)
    pyautogui.write("Todos TT's")
    time.sleep(2)
    pyautogui.press('enter')
    print("✅ Campo preenchido.")

# === CONEXÃO COM PLANILHA GOOGLE SHEETS ===
def conectar_planilha():
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            caminho_absoluto(os.path.join("boot-456117-9322a3608591.json")),
            scope
        )
        client = gspread.authorize(creds)
        sheet = client.open_by_url(
            "https://docs.google.com/spreadsheets/d/1BAN52GhyN2gP8mNAbX7T6HLOlyYW9eeHeGEPabTdqqM/edit?gid=0"
        ).sheet1
        return sheet
    except Exception as e:
        print("Erro ao conectar à planilha:", e)
        sys.exit(1)

# === MANIPULAÇÃO DE DATAS ===
def obter_ultima_data(sheet):
    col_datas = sheet.col_values(6)
    datas_validas = [d for d in col_datas if d.strip() and d.strip() != "-"]
    if not datas_validas:
        raise ValueError("Nenhuma data válida encontrada na coluna 6")
    ultima_data_str = datas_validas[-1]
    return datetime.strptime(ultima_data_str, "%d/%m/%Y")

def datas_faltantes(ultima_data):
    hoje = datetime.now()
    proxima_data = ultima_data + timedelta(days=1)
    datas = []
    while proxima_data.date() <= hoje.date():
        datas.append(proxima_data)
        proxima_data += timedelta(days=1)
    return datas

# === EXTRAÇÃO DE DADOS DO CLIENTE ===
def dados_cliente(driver, data_selecionada):
    try:
        time.sleep(3)
        nome_completo = driver.find_element(
            By.CSS_SELECTOR, "div[data-ofsc-role='page-description-text']"
        ).text.strip()
        nome_tecnico = nome_completo.split(",")[0]
        time.sleep(3)

        nome_cliente = driver.find_element(
            By.CSS_SELECTOR, "div[data-label='cname']"
        ).text.strip()
        ordem_servico = driver.find_element(
            By.CSS_SELECTOR, "div[data-label='appt_number']"
        ).text.strip()
        time.sleep(3)

        pyautogui.scroll(-1000)
        time.sleep(2)

        btn_info = pyautogui.locateOnScreen(
            caminho_img("informacoes_cliente.png"), confidence=0.8
        )
        if not btn_info:
            return None
        pyautogui.click(btn_info)
        time.sleep(2)

        celular = None
        telefone = None
        try:
            celular = driver.find_element(
                By.CSS_SELECTOR, "a[data-label='cmobile']"
            ).text.strip()
        except Exception:
            pass
        try:
            telefone = driver.find_element(
                By.CSS_SELECTOR, "a[data-label='cphone']"
            ).text.strip()
        except Exception:
            pass

        numero_final = celular or telefone

        # Fechar info e voltar
        time.sleep(5)
        btn_fechar = pyautogui.locateOnScreen(
            caminho_img("fechar_info_cliente.png"), confidence=0.6
        )
        if btn_fechar:
            pyautogui.click(btn_fechar)
            time.sleep(3)

        time.sleep(5)
        btn_voltar = pyautogui.locateOnScreen(
            caminho_img("botao_voltar_linhas.png"), confidence=0.6
        )
        if btn_voltar:
            pyautogui.click(btn_voltar)
            time.sleep(3)

        if not numero_final:
            return None

        return {
            "nome": nome_cliente,
            "telefone": numero_final,
            "tecnico": nome_tecnico,
            "ordem": ordem_servico,
            "data": data_selecionada.strftime('%d/%m/%Y')
        }
    except Exception as e:
        print("Erro ao extrair dados do cliente:", e)
        return None

# === SALVAMENTO DE DADOS NA PLANILHA (PULA COLUNA 4) ===
def salvar_dados(dados, sheet):
    try:
        todas_linhas = sheet.get_all_values()
        ordens_registradas = {linha[4] for linha in todas_linhas[1:] if len(linha) >= 5}
        if dados["ordem"] in ordens_registradas:
            return  # já existe
        proxima_linha = len(todas_linhas) + 1
        sheet.update_cell(proxima_linha, 1, dados["nome"] )      # Coluna 1
        sheet.update_cell(proxima_linha, 2, dados["telefone"])   # Coluna 2
        sheet.update_cell(proxima_linha, 3, dados["tecnico"])    # Coluna 3
        # Coluna 4 pulada
        sheet.update_cell(proxima_linha, 5, dados["ordem"])      # Coluna 5
        sheet.update_cell(proxima_linha, 6, dados["data"])       # Coluna 6
    except Exception as e:
        print("Erro ao salvar dados na planilha:", e)

# === AUXILIARES PARA DETECÇÃO DE CLIQUES ===
def largura_bloco(x_inicial, y, cor_bloco=(121, 182, 235)):
    largura = 0
    try:
        while True:
            if pyautogui.pixel(x_inicial + largura, y) == cor_bloco:
                largura += 1
            else:
                break
    except Exception:
        pass
    return largura

def foi_clicado(x, posicoes_clicadas):
    return any(pos in posicoes_clicadas for pos in range(x, x + 10))

# === REGISTRAR DEFEITOS E DADOS ===
def registrar_defeitos(driver, sheet, data):
    ordens_processadas = carregar_os_processados()
    coordenadas_linhas = [
        (331, 364), (308, 392), (358, 419), (338, 448),
        (312, 474), (316, 505), (316, 532), (326, 555),
        (314, 561),
    ]
    for linha_x, linha_y in coordenadas_linhas:
        x = 300
        posicoes_clicadas = set()
        passos = 0
        while x < 1500 and passos < 300:
            try:
                if pyautogui.pixel(x, linha_y) == (121, 182, 235) and not foi_clicado(x, posicoes_clicadas):
                    pyautogui.moveTo(x, linha_y, duration=0.1)
                    pyautogui.doubleClick()
                    time.sleep(2)
                    dados = dados_cliente(driver, data)
                    if dados and dados["ordem"] not in ordens_processadas:
                        salvar_os_processados(dados["ordem"])
                        ordens_processadas.add(dados["ordem"])
                        salvar_dados(dados, sheet)
                    largura = largura_bloco(x, linha_y)
                    for offset in range(largura + 5):
                        posicoes_clicadas.add(x + offset)
                    x += largura + 5
                else:
                    x += 5
                passos += 1
            except Exception as e:
                print("Erro em registrar_defeitos loop:", e)
                x += 5
                passos += 1

# === PROCESSAMENTO DAS DATAS ===
def processar_datas(driver, sheet):
    limpar_arquivo_temporario()
    try:
        ultima = obter_ultima_data(sheet)
        faltantes = datas_faltantes(ultima)
        for data in faltantes:

            # Marcos
            time.sleep(10)
            btn_marcos = pyautogui.locateOnScreen(
                caminho_img("marcos.png"), confidence=0.9
            )
            if not btn_marcos:
                return None
            pyautogui.click(btn_marcos)
            time.sleep(2)
            abrir_calendario(driver)
            selecionar_data(driver, data)
            time.sleep(1)
            pyautogui.moveTo(x=415, y=105)
            time.sleep(3)
            localizar_e_clicar_ver()
            time.sleep(5)
            preencher_campo_com_pyautogui()
            time.sleep(3)
            pyautogui.moveTo(x=415, y=105)
            time.sleep(5)
            registrar_defeitos(driver, sheet, data)
            time.sleep(8)

            # Sérgio
            time.sleep(10)
            btn_sergio = pyautogui.locateOnScreen(
                caminho_img("sergio.png"), confidence=0.9
            )
            if not btn_sergio:
                return None
            pyautogui.click(btn_sergio)
            time.sleep(2)
            abrir_calendario(driver)
            selecionar_data(driver, data)
            time.sleep(1)
            pyautogui.moveTo(x=415, y=105)
            time.sleep(3)
            localizar_e_clicar_ver()
            time.sleep(1)
            preencher_campo_com_pyautogui()
            time.sleep(3)
            pyautogui.moveTo(x=415, y=105)
            time.sleep(5)
            registrar_defeitos(driver, sheet, data)
            time.sleep(8)

            # Anastacio
            time.sleep(10)
            btn_anastacio = pyautogui.locateOnScreen(
                caminho_img("anastacio.png"), confidence=0.9
            )
            if not btn_anastacio:
                return None
            pyautogui.click(btn_anastacio)
            time.sleep(2)
            abrir_calendario(driver)
            selecionar_data(driver, data)
            time.sleep(1)
            pyautogui.moveTo(x=415, y=105)
            time.sleep(3)
            localizar_e_clicar_ver()
            time.sleep(1)
            preencher_campo_com_pyautogui()
            time.sleep(3)
            pyautogui.moveTo(x=415, y=105)
            time.sleep(5)
            registrar_defeitos(driver, sheet, data)
            time.sleep(8)

            # Lazaro
            time.sleep(10)
            btn_lazaro = pyautogui.locateOnScreen(
                caminho_img("lazaro.png"), confidence=0.8
            )
            if not btn_lazaro:
                return None
            pyautogui.click(btn_lazaro)
            time.sleep(2)
            abrir_calendario(driver)
            selecionar_data(driver, data)
            time.sleep(1)
            pyautogui.moveTo(x=415, y=105)
            time.sleep(3)
            localizar_e_clicar_ver()
            time.sleep(1)
            preencher_campo_com_pyautogui()
            time.sleep(3)
            pyautogui.moveTo(x=415, y=105)
            time.sleep(5)
            registrar_defeitos(driver, sheet, data)
            time.sleep(8)

    except Exception as e:
        print("Erro em processar_datas:", e)
        traceback.print_exc()  # <<< ADICIONE ISSO para ver o erro detalhado
    finally:
        limpar_arquivo_temporario()

# === FUNÇÕES DE NAVEGAÇÃO E LOGIN ===
def abrir_chrome():
    options = Options()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.get("https://gvt.fs.ocs.oraclecloud.com/")
    return driver

def login(driver):
    time.sleep(3)
    driver.find_element(By.ID, "username").send_keys(USUARIO)
    driver.find_element(By.ID, "password").send_keys(SENHA)
    driver.find_element(By.XPATH, "//span[text()='Conectar']").click()

def verificação(driver):
    time.sleep(3)
    campo_senha = driver.find_element(By.ID, "password")
    campo_senha.send_keys(SENHA)
    checkbox = driver.find_element(By.ID, "delsession")
    if not checkbox.is_selected():
        checkbox.click()
    driver.find_element(By.XPATH, "//span[text()='Conectar']").click()
    time.sleep(5)

def menu():
    time.sleep(3)
    btn = pyautogui.locateOnScreen(caminho_img("menu.png"), confidence=0.8)
    if btn: pyautogui.click(btn)
    time.sleep(1)
    btn = pyautogui.locateOnScreen(caminho_img("atividades.png"), confidence=0.7)
    if btn: pyautogui.click(btn)

# === CALENDÁRIO ===
def abrir_calendario(driver):
    try:
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "ko-hint-overlay"))
        )
        spans = WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((
                By.XPATH,
                "//span[contains(@class, 'app-button-title') and contains(text(), '/')]"
            ))
        )
        btn = spans[0].find_element(By.XPATH, "./ancestor::button[1]")
        driver.execute_script("arguments[0].scrollIntoView(true);", btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", btn)
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "div.kd-date-picker-popup"))
        )
    except Exception as e:
        print("Erro ao abrir calendário:", e)

def selecionar_data(driver, data):
    try:
        dia = data.day
        mes = data.month - 1
        ano = data.year
        seletor = f"td[data-day='{dia}'][data-month='{mes}'][data-year='{ano}']"
        elemento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, seletor))
        )
        elemento.click()
        time.sleep(1)
    except Exception as e:
        print("Erro ao selecionar data:", e)

# === EXECUÇÃO PRINCIPAL ===
if __name__ == '__main__':
    sheet = conectar_planilha()
    driver = abrir_chrome()
    login(driver)
    verificação(driver)
    fechar_popup()
    menu()
    processar_datas(driver, sheet)
