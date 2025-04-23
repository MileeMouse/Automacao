import tkinter as tk
import tkinter.ttk as ttk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox, simpledialog, Listbox
import json
import threading
from dotenv import set_key
from coleta_defeitos import conectar_planilha, abrir_chrome, login, verificação, fechar_popup, menu, processar_datas
from bot_wpp import enviar_mensagens_whatsapp
import subprocess
import time
import os

ARQUIVO_USUARIOS = 'usuarios.json'
ARQUIVO_ENV = '.env'

def carregar_usuarios():
    try:
        with open(ARQUIVO_USUARIOS, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def salvar_usuarios(usuarios):
    with open(ARQUIVO_USUARIOS, 'w', encoding='utf-8') as f:
        json.dump(usuarios, f, indent=4)

def atualizar_credencial(chave, novo_valor):
    novo_valor = novo_valor.strip('"').strip("'")
    with open(ARQUIVO_ENV, 'r+', encoding='utf-8') as f:
        linhas = f.readlines()
        f.seek(0)
        for linha in linhas:
            if linha.startswith(chave.upper()):
                f.write(f"{chave.upper()}={novo_valor}\n")
            else:
                f.write(linha)
        f.truncate()

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Coleta de Defeitos")
        self.root.geometry("1280x768")
        self.root.resizable(False, False)
        self.usuario_logado = None
        self.tecnicos_confirmados = []
        self.ultimo_supervisores_utilizados = []

        self.usuarios = carregar_usuarios()

        style = tb.Style()
        style.configure("TNotebook.Tab", padding=10)
        style.configure("TNotebook", tabposition='n')
        style.map("TNotebook.Tab", background=[("selected", "#007BFF")], foreground=[("selected", "white")])

        self.criar_tela_login()

    def criar_tela_login(self):
        self.frame_login = tb.Frame(self.root, padding=30)
        self.frame_login.place(relx=0.5, rely=0.5, anchor='center')

        tb.Label(self.frame_login, text="Login de Usuário", font=("Segoe UI", 16, "bold"), bootstyle="primary").pack(pady=(0, 20))
        tb.Label(self.frame_login, text="Usuário:", font=("Segoe UI", 10)).pack(anchor='w')
        self.entry_usuario = tb.Entry(self.frame_login, width=30)
        self.entry_usuario.pack(pady=(0, 10))
        tb.Label(self.frame_login, text="Senha:", font=("Segoe UI", 10)).pack(anchor='w')
        self.entry_senha = tb.Entry(self.frame_login, width=30, show="*")
        self.entry_senha.pack(pady=(0, 15))

        tb.Button(self.frame_login, text="Entrar", bootstyle="primary", width=30, command=self.verificar_login).pack()

    def verificar_login(self):
        usuario = self.entry_usuario.get()
        senha = self.entry_senha.get()
        if usuario in self.usuarios and self.usuarios[usuario]['senha'] == senha:
            self.usuario_logado = usuario
            self.empresa_logada = self.usuarios[usuario]['empresa']
            self.frame_login.destroy()
            self.criar_tela_principal()
        else:
            messagebox.showerror("Erro de Login", "Usuário ou senha incorretos.")

    def criar_tela_principal(self):
        self.tabs = tb.Notebook(self.root, bootstyle="primary")

        self.frame_coleta = tb.Frame(self.tabs, padding=20)
        self.frame_coleta_inner = tb.Frame(self.frame_coleta)
        self.frame_coleta_inner.pack(expand=True)

        self.tabs.add(self.frame_coleta, text="Coleta de Defeitos")

        tb.Label(self.frame_coleta_inner, text=f"Usuário: {self.usuario_logado}", font=("Segoe UI", 12)).pack(pady=5, anchor="center")
        tb.Label(self.frame_coleta_inner, text=f"Empresa: {self.empresa_logada}", font=("Segoe UI", 12)).pack(pady=5, anchor="center")

        self.criar_linha_botao(self.frame_coleta_inner, "Iniciar Coleta", self.iniciar_bot, "Aperte para iniciar a coleta dos defeitos")
        self.criar_linha_botao(self.frame_coleta_inner, "Encerrar Sistema", self.root.quit, "Fechar o aplicativo")
        self.criar_linha_botao(self.frame_coleta_inner, "Alterar Usuário do Bot", self.trocar_usuario, "Aperte para mudar o login de acesso no oracle")
        self.criar_linha_botao(self.frame_coleta_inner, "Alterar Senha do Bot", self.mudar_senha, "Aperte para mudar o login de acesso no oracle")

        self.frame_wpp = tb.Frame(self.tabs, padding=20)
        self.tabs.add(self.frame_wpp, text="Automacao Wpp")

        tb.Label(self.frame_wpp, text="Selecione as datas:", font=("Segoe UI", 10)).pack(anchor='center')
        tb.Button(self.frame_wpp, text="Selecionar Datas", bootstyle="primary", width=30, command=self.abrir_calendario).pack(pady=(0, 10), anchor='center')

        self.lista_datas = Listbox(self.frame_wpp, height=4)
        self.lista_datas.pack(pady=(0, 10), fill='x', padx=40)

        tb.Label(self.frame_wpp, text="Selecione os supervisores:", font=("Segoe UI", 10)).pack(anchor='center')
        self.lista_supervisores = Listbox(self.frame_wpp, selectmode='multiple', height=5)
        self.lista_supervisores.pack(pady=5, fill='x', padx=40)

        tb.Button(self.frame_wpp, text="Filtrar Técnicos", bootstyle="info", command=self.atualizar_tecnicos).pack(pady=5, anchor='center')

        tb.Label(self.frame_wpp, text="Selecione os técnicos:", font=("Segoe UI", 10)).pack(anchor='center')
        self.lista_tecnicos = Listbox(self.frame_wpp, selectmode='multiple', height=12)
        self.lista_tecnicos.pack(pady=10, fill='x', padx=40)

        frame_botoes_tecnicos = tb.Frame(self.frame_wpp)
        frame_botoes_tecnicos.pack(pady=5, anchor='center')

        tb.Button(frame_botoes_tecnicos, text="Selecionar Todos os Técnicos", bootstyle="success", width=25,
                  command=self.selecionar_todos_tecnicos).pack(side='left', padx=5)

        tb.Button(frame_botoes_tecnicos, text="Desmarcar Todos os Técnicos", bootstyle="danger", width=25,
                  command=self.desmarcar_todos_tecnicos).pack(side='left', padx=5)

        tb.Button(self.frame_wpp, text="Confirmar Técnicos Selecionados", bootstyle="secondary", width=30,
                  command=self.confirmar_tecnicos).pack(pady=5, anchor='center')

        tb.Button(self.frame_wpp, text="Iniciar WhatsApp", bootstyle="primary", width=30, command=self.iniciar_automacao_wpp).pack(pady=10, anchor='center')

        self.carregar_supervisores_e_tecnicos()
        self.tabs.pack(expand=1, fill="both")

    def criar_linha_botao(self, parent, texto_botao, comando, descricao):
        linha = tb.Frame(parent)
        linha.pack(pady=10, anchor='center')
        btn = tb.Button(linha, text=texto_botao, bootstyle="primary", width=25, command=comando)
        btn.pack(pady=2)
        label = tb.Label(linha, text=descricao, font=("Segoe UI", 9), anchor='center')
        label.pack()

    def abrir_calendario(self):
        subprocess.run(["python", "seletor_datas.py"])
        time.sleep(0.5)
        if os.path.exists("datas_selecionadas.json"):
            with open("datas_selecionadas.json", "r", encoding="utf-8") as f:
                datas = json.load(f)
            self.lista_datas.delete(0, 'end')
            for data in datas:
                self.lista_datas.insert('end', data)

    def carregar_supervisores_e_tecnicos(self):
        with open("supervisores_tecnicos.json", "r", encoding="utf-8") as f:
            self.dados_usuarios = json.load(f)
        self.lista_supervisores.delete(0, 'end')
        for supervisor in self.dados_usuarios:
            self.lista_supervisores.insert('end', supervisor)

    def atualizar_tecnicos(self, event=None):
        selecionados = self.lista_supervisores.curselection()
        supervisores = [self.lista_supervisores.get(i) for i in selecionados]
        self.ultimo_supervisores_utilizados = supervisores
        tecnicos_set = set()
        for sup in supervisores:
            tecnicos_set.update(self.dados_usuarios.get(sup, []))
        self.lista_tecnicos.delete(0, 'end')
        for tec in tecnicos_set:
            self.lista_tecnicos.insert('end', tec)

    def iniciar_automacao_wpp(self):
        datas = list(self.lista_datas.get(0, 'end'))
        supervisores = self.ultimo_supervisores_utilizados
        tecnicos = self.tecnicos_confirmados

        if not datas:
            messagebox.showwarning("Faltam dados", "Você precisa selecionar ao menos uma data.")
            return
        if not supervisores:
            messagebox.showwarning("Faltam dados", "Você precisa ter filtrado técnicos com ao menos um supervisor.")
            return
        if not tecnicos:
            messagebox.showwarning("Faltam dados", "Você precisa confirmar os técnicos antes de continuar.")
            return

        threading.Thread(target=self.executar_envio_mensagens, args=(datas, supervisores, tecnicos), daemon=True).start()

    def executar_envio_mensagens(self, datas, supervisores, tecnicos):
        try:
            enviar_mensagens_whatsapp(datas, supervisores, tecnicos)
            messagebox.showinfo("Sucesso", "Mensagens enviadas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao enviar mensagens: {str(e)}")

    def iniciar_bot(self):
        try:
            sheet = conectar_planilha()
            driver = abrir_chrome()
            login(driver)
            verificação(driver)
            fechar_popup()
            menu()
            processar_datas(driver, sheet)
            messagebox.showinfo("Bot Finalizado", "Processamento de dados finalizado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao iniciar o bot de coleta: {str(e)}")

    def trocar_usuario(self):
        novo_usuario = simpledialog.askstring("Novo Usuário", "Digite o novo nome de usuário:")
        if novo_usuario:
            atualizar_credencial("USUARIO", novo_usuario)
            messagebox.showinfo("Alteração", "Usuário atualizado no .env com sucesso!")

    def mudar_senha(self):
        nova_senha = simpledialog.askstring("Nova Senha", "Digite a nova senha:", show='*')
        if nova_senha:
            atualizar_credencial("SENHA", nova_senha)
            messagebox.showinfo("Alteração", "Senha atualizada no .env com sucesso!")

    def selecionar_todos_tecnicos(self):
        self.lista_tecnicos.select_set(0, 'end')

    def desmarcar_todos_tecnicos(self):
        self.lista_tecnicos.select_clear(0, 'end')

    def confirmar_tecnicos(self):
        tecnicos = [self.lista_tecnicos.get(i) for i in self.lista_tecnicos.curselection()]
        if not tecnicos:
            messagebox.showwarning("Nenhum Técnico Selecionado", "Selecione ao menos um técnico antes de confirmar.")
            return
        self.tecnicos_confirmados = tecnicos
        messagebox.showinfo("Técnicos Confirmados", f"{len(tecnicos)} técnico(s) confirmados com sucesso.")

if __name__ == "__main__":
    root = tb.Window(themename="darkly")
    app = App(root)
    root.mainloop()
