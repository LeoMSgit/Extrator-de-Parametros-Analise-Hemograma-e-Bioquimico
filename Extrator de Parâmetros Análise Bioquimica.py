import os
import pandas as pd
import PyPDF2
import re

# Solicitar o caminho da pasta ao usuário
caminho_pasta = input("Por favor, insira o caminho completo da pasta com os arquivos PDF: ")

# Verificar se o caminho da pasta é válido
if not os.path.exists(caminho_pasta):
    print(f"O caminho fornecido '{caminho_pasta}' não é válido.")
    exit()

# Lista de parâmetros
parametros = [
    "ERITROCITOS", "HEMOGLOBINA", "HEMATÓCRITO", "V.C.M", "H.C.M", "C.H.C.M",
    "PLAQUETAS", "LEUCÓCITOS TOTAIS", "BASTONETES", "SEGMENTADOS", "LINFÓCITOS",
    "MONÓCITOS", "EOSINÓFILOS", "BASÓFILOS", "ALBUMINA", "BILIRRUBINA DIRETA",
    "BILIRRUBINA TOTAL", "CK", "CREATININA", "FOSFATASE ALCALINA", "GGT",
    "PROTEINA TOTAL", "AST", "ALT", "UREIA", "BILIRRUBINA INDIRETA"
]

# DataFrame para armazenar os dados
df_final = pd.DataFrame(columns=['Arquivo'] + parametros)

# Processar todos os PDFs da pasta
arquivos_pdf = [f for f in os.listdir(caminho_pasta) if f.endswith('.pdf')]

# Loop para processar cada arquivo PDF
for nome_pdf in arquivos_pdf:
    caminho_pdf = os.path.join(caminho_pasta, nome_pdf)

    # Abrir o arquivo PDF
    with open(caminho_pdf, 'rb') as file:
        reader = PyPDF2.PdfReader(file)

        # Extração de texto da primeira página
        page1 = reader.pages[0]
        text1 = page1.extract_text() or ""

        # Extração de texto da segunda página
        page2 = reader.pages[1]
        text2 = page2.extract_text() or ""

    # Criando um dicionário para armazenar os resultados
    resultados_dict = {parametro: None for parametro in parametros}

    # Extraindo valores da primeira página
    resultados_dict["ERITROCITOS"] = re.search(r"ERITROCITOS?\s+([\d.,]+)\s+m", text1, re.IGNORECASE)
    resultados_dict["HEMOGLOBINA"] = re.search(r"HEMOGLOBINA\s+([\d.,]+)\s+g/dL", text1, re.IGNORECASE)
    resultados_dict["HEMATÓCRITO"] = re.search(r"HEMATÓCRITO\s+([\d.,]+)\s+%", text1, re.IGNORECASE)
    resultados_dict["V.C.M"] = re.search(r"V\.C\.M\s+([\d.,]+)\s+fl", text1, re.IGNORECASE)
    resultados_dict["H.C.M"] = re.search(r"H\.C\.M\s+([\d.,]+)\s+pg", text1, re.IGNORECASE)
    resultados_dict["C.H.C.M"] = re.search(r"C\.H\.C\.M\s+([\d.,]+)\s+%", text1, re.IGNORECASE)
    resultados_dict["PLAQUETAS"] = re.search(r"PLAQUETAS\s+([\d.,]+)\s+µL", text1, re.IGNORECASE)
    resultados_dict["LEUCÓCITOS TOTAIS"] = re.search(r"LEUCÓCITOS TOTAIS\s+([\d.,]+)\s+/mm³", text1, re.IGNORECASE)
    resultados_dict["BASTONETES"] = re.search(r"BASTONETES\s+([\d.,]+)\s+(%|/mm³)", text1, re.IGNORECASE)
    resultados_dict["SEGMENTADOS"] = re.search(r"SEGMENTADOS\s+([\d.,]+)\s+(%|/mm³)", text1, re.IGNORECASE)
    resultados_dict["LINFÓCITOS"] = re.search(r"LINFÓCITOS\s+([\d.,]+)\s+(%|/mm³)", text1, re.IGNORECASE)
    resultados_dict["MONÓCITOS"] = re.search(r"MONÓCITOS\s+([\d.,]+)\s+(%|/mm³)", text1, re.IGNORECASE)
    resultados_dict["EOSINÓFILOS"] = re.search(r"EOSINÓFILOS\s+([\d.,]+)\s+(%|/mm³)", text1, re.IGNORECASE)
    resultados_dict["BASÓFILOS"] = re.search(r"BASÓFILOS\s+([\d.,]+)\s+(%|/mm³)", text1, re.IGNORECASE)

    # Associar valores aos parâmetros do dicionário
    for parametro in resultados_dict.keys():
        if resultados_dict[parametro]:
            resultados_dict[parametro] = resultados_dict[parametro].group(1).replace(',', '.')

    # Extraindo valores da segunda página
    resultados_second_page = re.findall(r"RESULTADO\.+:\s+([\d,.]+)", text2)
    # Associe os resultados da segunda página com os parâmetros
    for i, resultado in enumerate(resultados_second_page):
        if i < len(parametros) - 14:  # ajuste para os últimos 14 parâmetros
            resultados_dict[parametros[i + 14]] = resultado.replace(',', '.')  # começando da posição 14

    # Adicionar nome do arquivo ao dicionário
    resultados_dict['Arquivo'] = nome_pdf

    # Adicionar os dados ao DataFrame
    df_final = pd.concat([df_final, pd.DataFrame([resultados_dict])], ignore_index=True)

# Nome do arquivo Excel com base no nome da pasta
nome_pasta = os.path.basename(os.path.normpath(caminho_pasta))
nome_arquivo_excel = f'{nome_pasta}.xlsx'

# Salvar DataFrame em um arquivo Excel
df_final.to_excel(nome_arquivo_excel, index=False)

print(f"Arquivo Excel '{nome_arquivo_excel}' gerado com sucesso!")
