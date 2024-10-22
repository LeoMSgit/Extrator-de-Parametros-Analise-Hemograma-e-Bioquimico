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

# DataFrame para armazenar os parâmetros na primeira coluna
df_final = pd.DataFrame({'PARÂMETROS': parametros})

# Processar todos os PDFs da pasta
arquivos_pdf = [f for f in os.listdir(caminho_pasta) if f.endswith('.pdf')]


# Função para extrair valores com regex
def extrair_valor(regex, texto):
    match = re.search(regex, texto, re.IGNORECASE)
    return match.group(1).replace(',', '.') if match else None


# Loop para processar cada arquivo PDF
for nome_pdf in arquivos_pdf:
    caminho_pdf = os.path.join(caminho_pasta, nome_pdf)

    # Abrir o arquivo PDF
    with open(caminho_pdf, 'rb') as file:
        reader = PyPDF2.PdfReader(file)

        # Extração de texto da primeira página
        page1 = reader.pages[0]
        text1 = page1.extract_text() or ""

        # Extração de texto da segunda página (se houver)
        if len(reader.pages) > 1:
            page2 = reader.pages[1]
            text2 = page2.extract_text() or ""
        else:
            text2 = ""

    # Criando uma lista para armazenar os resultados
    resultados = []

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

    # Associe os resultados da segunda página com os parâmetros da lista a partir de "ALBUMINA"
    for i, resultado in enumerate(resultados_second_page):
        if i < len(parametros) - 14:  # ajuste para os últimos parâmetros
            resultados.append(resultado.replace(',', '.'))

    # Adicionar os resultados como uma nova coluna ao DataFrame
    df_final[nome_pdf] = pd.Series(resultados)

# Nome do arquivo Excel com base no nome da pasta
nome_pasta = os.path.basename(os.path.normpath(caminho_pasta))
nome_arquivo_excel = f'{nome_pasta}_resultados.xlsx'

# Salvar o DataFrame em um arquivo Excel com os dados na vertical
df_final.to_excel(nome_arquivo_excel, index=False)

print(f"Arquivo Excel '{nome_arquivo_excel}' gerado com sucesso!")
