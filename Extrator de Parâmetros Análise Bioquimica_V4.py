import os
import pandas as pd
import PyPDF2
import re
from openpyxl import load_workbook
from tkinter import Tk, Label, Button, filedialog, messagebox, StringVar
from tkinter.ttk import Progressbar
from concurrent.futures import ProcessPoolExecutor


# Função para selecionar a pasta
def selecionar_pasta():
    caminho = filedialog.askdirectory()
    if caminho:
        pasta_selecionada.set(caminho)


# Função para extrair valores com regex
def extrair_valor(regex, texto):
    match = re.search(regex, texto, re.IGNORECASE)
    return match.group(1).replace(',', '.') if match else None


# Função para processar um único PDF (para paralelismo)
def processar_pdf(caminho_pdf, parametros):
    resultados = []

    with open(caminho_pdf, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        page1 = reader.pages[0]
        text1 = page1.extract_text() or ""
        text2 = reader.pages[1].extract_text() if len(reader.pages) > 1 else ""

    # Extraindo valores da primeira página
    resultados.append(extrair_valor(r"ERITROCITOS?\s+([\d.,]+)\s+m", text1))
    resultados.append(extrair_valor(r"HEMOGLOBINA\s+([\d.,]+)\s+g/dL", text1))
    resultados.append(extrair_valor(r"HEMATÓCRITO\s+([\d.,]+)\s+%", text1))
    resultados.append(extrair_valor(r"V\.C\.M\s+([\d.,]+)\s+fl", text1))
    resultados.append(extrair_valor(r"H\.C\.M\s+([\d.,]+)\s+pg", text1))
    resultados.append(extrair_valor(r"C\.H\.C\.M\s+([\d.,]+)\s+%", text1))
    resultados.append(extrair_valor(r"PLAQUETAS\s+([\d.,]+)\s+µL", text1))
    resultados.append(extrair_valor(r"LEUCÓCITOS TOTAIS\s+([\d.,]+)\s+/mm³", text1))
    resultados.append(extrair_valor(r"BASTONETES\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"SEGMENTADOS\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"LINFÓCITOS\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"MONÓCITOS\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"EOSINÓFILOS\s+([\d.,]+)\s+(%|/mm³)", text1))
    resultados.append(extrair_valor(r"BASÓFILOS\s+([\d.,]+)\s+(%|/mm³)", text1))

    # Extraindo valores da segunda página
    resultados_second_page = re.findall(r"RESULTADO\.+:\s+([\d,.]+)", text2)
    for i, resultado in enumerate(resultados_second_page):
        if i < len(parametros) - 14:  # ajuste para os últimos parâmetros
            resultados.append(resultado.replace(',', '.'))

    return resultados


# Função para processar os PDFs da pasta selecionada
def processar_pdfs():
    caminho_pasta = pasta_selecionada.get()
    if not os.path.exists(caminho_pasta):
        messagebox.showerror("Erro", f"O caminho fornecido '{caminho_pasta}' não é válido.")
        return

    arquivos_pdf = [f for f in os.listdir(caminho_pasta) if f.endswith('.pdf')]
    if not arquivos_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF encontrado na pasta selecionada.")
        return

    # Atualizar barra de progresso
    progresso['maximum'] = len(arquivos_pdf)

    df_final = pd.DataFrame({'PARÂMETROS': parametros})

    with ProcessPoolExecutor() as executor:
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

    messagebox.showinfo("Concluído", f"Arquivo Excel '{nome_arquivo_excel}' gerado com sucesso!")


# Parâmetros de exemplo
parametros = [
    "ERITROCITOS", "HEMOGLOBINA", "HEMATÓCRITO", "V.C.M", "H.C.M", "C.H.C.M",
    "PLAQUETAS", "LEUCÓCITOS TOTAIS", "BASTONETES", "SEGMENTADOS", "LINFÓCITOS",
    "MONÓCITOS", "EOSINÓFILOS", "BASÓFILOS", "ALBUMINA", "BILIRRUBINA DIRETA",
    "BILIRRUBINA TOTAL", "CK", "CREATININA", "FOSFATASE ALCALINA", "GGT",
    "PROTEINA TOTAL", "AST", "ALT", "UREIA", "BILIRRUBINA INDIRETA"
]

# Criação da interface gráfica
root = Tk()
root.title("Extrator de Parâmetros de PDFs")

# Variáveis
pasta_selecionada = StringVar()

# Componentes da interface
Label(root, text="Caminho da pasta com PDFs:").grid(row=0, column=0, padx=10, pady=10)
Button(root, text="Selecionar Pasta", command=selecionar_pasta).grid(row=0, column=1, padx=100, pady=100)
Label(root, textvariable=pasta_selecionada).grid(row=1, column=0, columnspan=2, padx=10, pady=10)
Button(root, text="Iniciar Processamento", command=processar_pdfs).grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# Barra de progresso
progresso = Progressbar(root, orient="horizontal", length=300, mode="determinate")
progresso.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# Iniciar interface
root.mainloop()
