import os
import openpyxl
import tkinter as tk
from tkinter import messagebox
from tkinter import PhotoImage
from PIL import Image, ImageTk

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
    ra = ra_entry.get()

    if ra.lower() == "exit":
        root.quit()
    else:
        if validar_ra(ra, arquivo_xlsx):
            mostrar_mensagem_valido()
        else:
            mostrar_mensagem_invalido()

def ativar_verificacao_automatica():
    ra = ra_entry.get().strip()  # Remove espaços em branco do início e do final
    if ra:  # Verifica se o campo não está vazio
        verificar_ra()  # Chama a função verificar_ra() se o campo não estiver vazio
    root.after(5000, ativar_verificacao_automatica)


def mostrar_mensagem_valido():
    mensagem = tk.Tk()
    mensagem.title("Resultado da Validação")

    fonte_personalizada = ("Arial", 20)

    # Abre a imagem
    imagem_pillow = Image.open("C:\\Users\\ENZOVITORIANOPUTTINI\\Desktop\\Projeto_TCC\\img\\valido.png")

    # Converte a imagem para um formato suportado pelo tkinter
    imagem_tk = ImageTk.PhotoImage(imagem_pillow)

    # Cria o rótulo e associa a imagem convertida a ele
    label = tk.Label(root, image=imagem_tk)
    label.pack()

    textLabel = tk.Label(mensagem, text="Valido!", font=fonte_personalizada)
    textLabel.pack()

    mensagem.after(3000, mensagem.destroy)

def mostrar_mensagem_invalido():
    mensagem = tk.Tk()
    mensagem.title("Resultado da Validação")

    fonte_personalizada = ("Arial", 20)

    imagem = PhotoImage(file="img\\erro.png")
    imgLabel = tk.Label(mensagem, image=imagem)
    imgLabel.pack()

    textLabel = tk.Label(mensagem, text="Invalido!", font=fonte_personalizada)
    textLabel.pack()

    mensagem.after(3000, mensagem.destroy)
    

# Nome do arquivo XLSX com os RAs válidos (sem a extensão)
arquivo_xlsx = "Ras"

# Cria a janela principal
root = tk.Tk()
root.title("Validação de RA")

# Rótulo e entrada para digitar o RA
ra_label = tk.Label(root, text="Digite o RA ou 'exit' para sair:")
ra_label.pack()
ra_entry = tk.Entry(root)
ra_entry.pack()

# Botão para verificar o RA
verificar_button = tk.Button(root, text="Verificar", command=validar_ra)
verificar_button.pack()

# Ativar verificação automática após 0,5 segundos
root.after(500, ativar_verificacao_automatica)

root.mainloop()
