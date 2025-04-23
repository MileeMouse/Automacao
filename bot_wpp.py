import time
import urllib.parse
import gspread
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from oauth2client.service_account import ServiceAccountCredentials

def conectar_planilha():
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            os.path.abspath(os.path.join("boot-456117-9322a3608591.json")),
            scope
        )
        client = gspread.authorize(creds)
        sheet = client.open_by_url(
            "https://docs.google.com/spreadsheets/d/1BAN52GhyN2gP8mNAbX7T6HLOlyYW9eeHeGEPabTdqqM/edit?gid=0"
        ).sheet1
        return sheet
    except Exception as e:
        print("Erro ao conectar à planilha:", e)

def enviar_mensagens_whatsapp(datas_escolhidas, supervisores, tecnicos):
    import pandas as pd
    from datetime import datetime

    sheet = conectar_planilha()
    if not sheet:
        print("Erro: planilha não carregada.")
        return

    dados = sheet.get_all_records()
    df = pd.DataFrame(dados)
    df.columns = df.columns.str.strip().str.lower()

    for col in ['data', 'hype', 'irla', 'usuario', 'telefone']:
        if col not in df.columns:
            print(f"Coluna esperada não encontrada: {col}")
            return

    print("Dados da planilha carregados.")
    print(df[['usuario', 'telefone', 'irla', 'hype', 'data']].head())

    datas_formatadas = []
    for data in datas_escolhidas:
        try:
            dt = datetime.strptime(data.strip(), "%d/%m/%Y")
            datas_formatadas.append(dt.strftime("%d/%m/%Y"))
        except ValueError:
            print(f"Data inválida ignorada: {data}")
            continue

    supervisores = [s.strip().lower() for s in supervisores]
    tecnicos = [t.strip().lower() for t in tecnicos]

    df['data'] = df['data'].astype(str).str.strip()
    df['hype'] = df['hype'].astype(str).apply(lambda x: x.replace('\xa0', ' ').strip().lower())
    df['irla'] = df['irla'].astype(str).apply(lambda x: x.replace('\xa0', ' ').strip().lower())

    df_filtrado = df[
        df['data'].isin(datas_formatadas) &
        df['hype'].isin(supervisores) &
        df['irla'].isin(tecnicos)
    ]

    if df_filtrado.empty:
        print("Nenhum dado encontrado com os filtros escolhidos.")
        return

    print(f"{len(df_filtrado)} registros encontrados após filtros.")

    # Dicionário com telefones dos técnicos
    telefones_tecnicos = {
        "erick souza de carvalho": "41 91234-5678",
        "edimar marcondes loubaque": "41 99876-5432",
        "rafael nascimento ribeiro": "41 91111-2222",
        "diogo spelier de castro pereira": "41 98888-0000"
    }

    navegador = webdriver.Chrome()
    navegador.get("https://web.whatsapp.com")

    while len(navegador.find_elements(By.ID, 'side')) < 1:
        time.sleep(1)
    time.sleep(2)

    for linha in df_filtrado.itertuples():
        cliente = getattr(linha, 'usuario')
        telefone = getattr(linha, 'telefone')
        tecnico = getattr(linha, 'irla')
        supervisor = getattr(linha, 'hype')

        telefone_tecnico = telefones_tecnicos.get(tecnico.lower(), "(telefone não informado)")

        mensagem = f'''
Bom dia/Boa tarde,

Tipo de atividade: Reparo
Nome do cliente: {cliente}
Telefone do técnico: {tecnico.title()} - {telefone_tecnico}
Telefone do supervisor: {supervisor}

Sua garantia para o reparo é de 30 dias, válida diretamente com o técnico que realizou a instalação. Se precisar de algo ou tiver qualquer dúvida, pode entrar em contato comigo por este número.

Ah, quase me esqueci! Você receberá um pedido de avaliação do meu atendimento por e-mail. Se puder deixar sua avaliação com 5 estrelas, vai ajudar muito e eu ficarei muito grato!

Estou à disposição para o que precisar.'''

        texto = urllib.parse.quote(mensagem)
        link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto}"
        navegador.get(link)

        while len(navegador.find_elements(By.ID, 'side')) < 1:
            time.sleep(1)
        time.sleep(5)

        try:
            time.sleep(15)
            navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[2]/button/span').click()
            print(f"Mensagem enviada para {cliente} ({telefone})")
        except:
            print(f"Não foi possível enviar para {cliente} ({telefone})")
        time.sleep(2)

    print("Todos os envios concluídos. Encerrando navegador.")
    navegador.quit()
    return
