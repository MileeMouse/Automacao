
import tkinter as tk
from tkinter import messagebox, Listbox
from tkcalendar import Calendar

def adicionar_data():
    data = cal.get_date()
    if data not in listbox.get(0, 'end'):
        listbox.insert('end', data)

root = tk.Tk()
root.title("Selecionar Datas")
root.geometry("300x400")

tk.Label(root, text="Selecione a(s) data(s):", font=("Segoe UI", 10)).pack(pady=(10, 0))

cal = Calendar(root, date_pattern='dd/mm/yyyy')
cal.pack(pady=10)

tk.Button(root, text="Adicionar Data", command=adicionar_data).pack(pady=(0, 10))

listbox = Listbox(root, height=6)
listbox.pack(padx=10, pady=10, fill='both', expand=True)

tk.Button(root, text="Mostrar Selecionadas", command=lambda: messagebox.showinfo("Datas", ", ".join(listbox.get(0, 'end')))).pack(pady=5)

root.mainloop()
