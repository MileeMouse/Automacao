
import tkinter as tk
from tkinter import Listbox, Button, Label
from tkcalendar import Calendar
import json

def main():
    datas_selecionadas = []

    def adicionar_data():
        data = cal.get_date()
        if data not in datas_selecionadas:
            datas_selecionadas.append(data)
            listbox.insert('end', data)

    def confirmar():
        with open("datas_selecionadas.json", "w", encoding="utf-8") as f:
            json.dump(datas_selecionadas, f)
        janela.destroy()

    janela = tk.Tk()
    janela.title("Selecionar Datas")
    janela.geometry("300x400")

    Label(janela, text="Selecione a(s) data(s):").pack(pady=5)
    cal = Calendar(janela, date_pattern='dd/mm/yyyy')
    cal.pack(pady=5)

    Button(janela, text="Adicionar Data", command=adicionar_data).pack(pady=5)

    listbox = Listbox(janela, height=6)
    listbox.pack(padx=10, pady=5, fill='both', expand=True)

    Button(janela, text="Confirmar Seleção", command=confirmar).pack(pady=5)

    janela.mainloop()

if __name__ == "__main__":
    main()
