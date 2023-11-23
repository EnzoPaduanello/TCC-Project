import os
import openpyxl
import tkinter as tk
from tkinter import messagebox

def validar_ra(ra, arquivo_xlsx):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    arquivo_xlsx_path = os.path.join(script_dir, arquivo_xlsx + ".xlsx")

    workbook = openpyxl.load_workbook(arquivo_xlsx_path)
    sheet = workbook.active

    for row in sheet.iter_rows(values_only=True):
        if row[0] == ra:
            workbook.close()
            return True

    workbook.close()
    return False

def verificar_ra():
    ra = entrada_ra.get()
    if ra.lower() == "exit":
        root.quit()
    elif validar_ra(ra, arquivo_xlsx):
        messagebox.showinfo("Resultado", "RA validado com sucesso!")
    else:
        messagebox.showerror("Resultado", "RA inválido!")

# Nome do arquivo XLSX com os RAs válidos (sem a extensão)
arquivo_xlsx = "Ras"

# Cria a janela principal
root = tk.Tk()
root.title("Validação de RA")
root.geometry("300x300")

# Rótulo e entrada para digitar o RA
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

label_ra = tk.Label(frame, text="Digite o RA ou 'exit' para sair:")
label_ra.pack()

entrada_ra = tk.Entry(frame)
entrada_ra.pack()

# Vincula o evento <Return> (ou <Enter>) ao Entry para chamar verificar_ra
entrada_ra.bind("<Return>", verificar_ra)

botao_validar = tk.Button(frame, text="Validar RA", command=verificar_ra)
botao_validar.pack()

root.mainloop()