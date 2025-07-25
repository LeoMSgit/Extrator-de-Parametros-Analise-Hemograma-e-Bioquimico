import os
import pandas as pd
import PyPDF2
import re
from openpyxl import load_workbook
from tkinter import Tk, Label, Button, filedialog, messagebox, StringVar
from tkinter.ttk import Progressbar, Style, Frame
from concurrent.futures import ThreadPoolExecutor
import webbrowser


# Selecionar a pasta
def selecionar_pasta():
    caminho = filedialog.askdirectory()
    if caminho:
        pasta_selecionada.set(caminho)

# Verificar se o caminho da pasta é válido
if not os.path.exists(caminho_pasta):
    print(f"O caminho fornecido '{caminho_pasta}' não é válido.")
    exit()

# Lista de parâmetros
parametros = [
    "ERITROCITOS", "HEMOGLOBINA", "HEMATÓCRITO", "V.C.M", "H.C.M", "C.H.C.M",
    "PLAQUETAS", "LEUCÓCITOS TOTAIS", "BASTONETES%", "SEGMENTADOS%", "LINFÓCITOS%",
    "MONÓCITOS%", "EOSINÓFILOS%", "BASÓFILOS%", "BASTONETES", "SEGMENTADOS", "LINFÓCITOS",
    "MONÓCITOS", "EOSINÓFILOS", "BASÓFILOS", "ALBUMINA", "BILIRRUBINA DIRETA",
    "BILIRRUBINA TOTAL", "CK", "CREATININA", "FOSFATASE ALCALINA", "GGT",
    "PROTEINA TOTAL", "AST", "ALT", "UREIA", "BILIRRUBINA INDIRETA"
]

# DataFrame para armazenar os parâmetros na primeira coluna
df_final = pd.DataFrame({'PARÂMETROS': parametros})

# Processar todos os PDFs da pasta
arquivos_pdf = [f for f in os.listdir(caminho_pasta) if f.endswith('.pdf')]


# Função para extrair valores com regex
def extrair_valor(regex, texto):
    match = re.search(regex, texto, re.IGNORECASE)
    return match.group(1).replace('%', '') if match else None

# Função específica para extrair valor absoluto após a porcentagem
def extrair_valor_absoluto(label, texto):
    regex = rf"{label}\s+\d+\s+%\s+([\d.,]+)"
    return extrair_valor(regex, texto)

# Loop para processar cada arquivo PDF
for nome_pdf in arquivos_pdf:
    caminho_pdf = os.path.join(caminho_pasta, nome_pdf)

    with open(caminho_pdf, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text1 = reader.pages[0].extract_text() or ""
        text2 = reader.pages[1].extract_text() if len(reader.pages) > 1 else ""

    resultados = []

    # Extraindo valores da primeira página
    resultados.append(extrair_valor(r"ERITROCITOS?\s+([\d.,]+)\s+m", text1))
    resultados.append(extrair_valor(r"HEMOGLOBINA\s+([\d.,]+)\s+g/dL", text1))
    resultados.append(extrair_valor(r"HEMATÓCRITO\s+([\d.,]+)\s+%", text1))
    resultados.append(extrair_valor(r"V\.C\.M\s+([\d.,]+)\s+fl", text1))
    resultados.append(extrair_valor(r"H\.C\.M\s+([\d.,]+)\s+pg", text1))
    resultados.append(extrair_valor(r"C\.H\.C\.M\s+([\d.,]+)\s+%", text1))
    resultados.append(extrair_valor(r"PLAQUETAS\s+([\d.,]+)\s+µL", text1))

    # Leucócitos totais
    leuc_tot = extrair_valor(r"LEUCÓCITOS TOTAIS\s+([\d.,]+)\s+/mm³", text1)
    resultados.append(leuc_tot)
    resultados.append(extrair_valor(r"BASTONETES\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"SEGMENTADOS\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"LINFÓCITOS\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"MONÓCITOS\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"EOSINÓFILOS\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"BASÓFILOS\s+([\d.,]+)\s+(%|/mm³)", text1))
    # Convertendo para int para possível uso posterior (se necessário)
    leuc_tot_int = int(leuc_tot.replace('.', '').replace(',', '')) if leuc_tot else 0

    # Valores absolutos do leucograma
    resultados.append(extrair_valor_absoluto("BASTONETES", text1))
    resultados.append(extrair_valor_absoluto("SEGMENTADOS", text1))
    resultados.append(extrair_valor_absoluto("LINFÓCITOS", text1))
    resultados.append(extrair_valor_absoluto("MONÓCITOS", text1))
    resultados.append(extrair_valor_absoluto("EOSINÓFILOS", text1))
    resultados.append(extrair_valor_absoluto("BASÓFILOS", text1))

    # Extraindo resultados da segunda página
    resultados_second_page = re.findall(r"RESULTADO\.+:\s+([\d,.]+)", text2)

    for i, resultado in enumerate(resultados_second_page):
        if i < len(parametros) - 22:
            resultados.append(resultado)

    # Adicionar os resultados ao DataFrame
    df_final[nome_pdf] = pd.Series(resultados)

# Processar os PDFs da pasta selecionada
def processar_pdfs():
    caminho_pasta = pasta_selecionada.get()
    if not os.path.exists(caminho_pasta):
        messagebox.showerror("Erro", f"O caminho fornecido '{caminho_pasta}' não é válido.")
        return

    arquivos_pdf = [f for f in os.listdir(caminho_pasta) if f.endswith('.pdf')]
    if not arquivos_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF encontrado na pasta selecionada.")
        return

    # Atualizar barra de progresso e mensagem de status
    progresso['maximum'] = len(arquivos_pdf)
    status_label.config(text="Processando PDFs...")

    df_final = pd.DataFrame({'PARÂMETROS': parametros})

    with ThreadPoolExecutor() as executor:
        for i, (nome_pdf, resultados) in enumerate(zip(arquivos_pdf, executor.map(processar_pdf,
                                                                                  [os.path.join(caminho_pasta, f) for f
                                                                                   in arquivos_pdf],
                                                                                  [parametros] * len(arquivos_pdf)))):
            df_final[nome_pdf] = pd.Series(resultados)
            progresso['value'] = i + 1
            root.update_idletasks()

    # Nome do arquivo Excel com base no nome da pasta
    nome_pasta = os.path.basename(os.path.normpath(caminho_pasta))
    nome_arquivo_excel = f'{nome_pasta}_resultados.xlsx'

    # Salvar o DataFrame em um arquivo Excel com os dados na vertical
    df_final.to_excel(nome_arquivo_excel, index=False)

    # Ajustar a largura da coluna A (PARÂMETROS)
    wb = load_workbook(nome_arquivo_excel)
    ws = wb.active
    ws.column_dimensions['A'].width = 20
    wb.save(nome_arquivo_excel)

    status_label.config(text="Processamento concluído!")
    messagebox.showinfo("Concluído", f"Arquivo Excel '{nome_arquivo_excel}' gerado com sucesso!")


# Abrir o LinkedIn no navegador
def abrir_linkedin(event):
    webbrowser.open("https://www.linkedin.com/in/leomsantos/")

# Abrir o GitHub no navegador
def abrir_github(event):
    webbrowser.open("https://github.com/LeoMSgit")


# Criação da interface gráfica
root = Tk()
root.title("Extrator de Parâmetros de PDFs")
root.geometry("440x380")
root.configure(bg='#f0f0f0')

# Estilo Clam
style = Style()
style.theme_use('clam')
style.configure('TLabel', font=('Arial', 12))
style.configure('TButton', font=('Arial', 10), padding=6)
style.configure('TProgressbar', thickness=20)

# Variáveis
pasta_selecionada = StringVar()

# Frame principal
frame = Frame(root, padding="50")
frame.grid(row=0, column=0, padx=20, pady=20)

# Componentes da interface
Label(frame, text="Caminho da pasta com PDFs:").grid(row=0, column=0, sticky="w")
Button(frame, text="Selecionar Pasta", command=selecionar_pasta).grid(row=0, column=1, padx=20)
Label(frame, textvariable=pasta_selecionada, relief="sunken", width=40).grid(row=1, column=0, columnspan=2, pady=10)
Button(frame, text="Iniciar Processamento", command=processar_pdfs).grid(row=2, column=0, columnspan=2, pady=10)

# Barra de progresso
progresso = Progressbar(frame, orient="horizontal", length=300, mode="determinate")
progresso.grid(row=3, column=0, columnspan=2, pady=10)

# Rótulo de status
status_label = Label(frame, text="", relief="sunken", anchor="w")
status_label.grid(row=4, column=0, columnspan=2, sticky="we", pady=10)

# Marca d'água
Label(root, text="Criado por Leonardo Miguel dos Santos", font=('Arial', 10)).grid(row=1, column=0, pady=(0, 2))
links_label = Label(root, text="LinkedIn: in/leomsantos | GitHub: @LeoMSgit",
                    fg="blue", cursor="hand2", font=('Arial', 10, 'underline'))
links_label.bind("<Button-1>", lambda e: abrir_linkedin(e) if e.x < 145 else abrir_github(e))
links_label.grid(row=2, column=0, pady=(0, 10))

# Iniciar interface
root.mainloop()
