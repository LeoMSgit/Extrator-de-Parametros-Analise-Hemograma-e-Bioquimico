# **Extrator de Parâmetros Análise de Hemograma e Bioquímico**

## **Descrição**
O **Extrator de Parâmetros Análise de Hemograma e Bioquímico** é uma ferramenta que extrai automaticamente valores de exames laboratoriais de arquivos PDF e os organiza em um arquivo Excel. Ele facilita a coleta de dados de PDFs e apresenta os resultados em um formato organizado, pronto para ser usado.

---

## **Requisitos**
- **Sistema Operacional:** Windows
- Arquivos PDF localizados em uma pasta, todos com resultados laboratoriais que seguem a mesma estrutura.

---

## Instruções de Uso

 - Versão Google Colab (Jupyter) disponível [AQUI](https://colab.research.google.com/github/LeoMSgit/Extrator-de-Parametros-Analise-Hemograma-e-Bioquimico/blob/main/Google%20Colab%20(Cloud)%20-%20Extrator_de_Par%C3%A2metros_An%C3%A1lise_de_Hemograma_e_Bioqu%C3%ADmico.ipynb).
   <br />
   <br />

1. Baixe o arquivo executável do software via GitHub [AQUI](https://github.com/LeoMSgit/Extrator-de-Parametros-Analise-Hemograma-e-Bioquimico/releases/download/release_3/Extrator.de.Parametros.Analise.Bioquimica_V7.exe) ou o arquivo compactado via Google Drive [AQUI](https://drive.google.com/file/d/1KJoWQ1pmojDkjbuWDafSpXSFYcjz5JKC/view?usp=sharing).

2. **Extraia o programa:**
   - Utilize uma ferramente como WinRar para extrair o software.

3. **Abra o programa:**
   - Execute o arquivo `.exe` gerado. Uma janela do console será aberta, solicitando o caminho da pasta com os arquivos PDF.

![](https://raw.githubusercontent.com/LeoMSgit/Extrator-de-Parametros-Analise-Hemograma-e-Bioquimico/refs/heads/main/Print/Screenshot_1.png)

4. **Insira o caminho completo da pasta:**
   - No campo de entrada do console, insira o caminho completo da pasta onde estão localizados os arquivos PDF.
   - **Exemplo:**  
     `C:\Users\Usuario\Documents\ExamesPDF`
   
5. **Escolha o Modelo de Laudo:**
   - Devido à diferentes modelos do PDF utilizados por diferentes laboratórios, resultados imprecisos podem ser encontrados caso um laudo PDF seja processado pelo modelo diferente do estabelecido na concepção do Software

6. **Processamento dos arquivos:**
   - O programa irá automaticamente acessar todos os PDFs dentro da pasta especificada, extrair os valores de interesse e organizá-los no formato necessário.


7. **Resultado do processamento:**
   - O programa criará um arquivo Excel (`.xlsx`) com o nome da pasta que contém os arquivos PDF.
   - O arquivo Excel terá os seguintes campos:
     - **Coluna A:** Lista de parâmetros de exames
         - A versão atual do **Extrator de Parâmetros Análise Bioquimica** suporta a extração dos seguintes parâmetros dos arquivos PDF:

            - ERITROCITOS
            - HEMOGLOBINA
            - HEMATÓCRITO
            - V.C.M
            - H.C.M
            - C.H.C.M
            - PLAQUETAS
            - LEUCÓCITOS TOTAIS
            - BASTONETES
            - SEGMENTADOS
            - LINFÓCITOS
            - MONÓCITOS
            - EOSINÓFILOS
            - BASÓFILOS
            - ALBUMINA
            - BILIRRUBINA DIRETA
            - BILIRRUBINA TOTAL
            - CK
            - CREATININA
            - FOSFATASE ALCALINA
            - GGT
            - PROTEINA TOTAL
            - AST
            - ALT
            - UREIA
            - BILIRRUBINA INDIRETA 
     - **Colunas subsequentes:** Para cada PDF processado, uma nova coluna será gerada, contendo os valores extraídos daquele arquivo PDF.
   
   - A célula A1 será chamada de **"PARÂMETROS"** e todas as colunas subsequentes trarão o nome dos arquivos PDF processados (ex.: A1.pdf, A2.pdf, etc.).
  
   ![Descrição da Imagem](https://i.imgur.com/YCoevYA.png)


8. **Finalização:**
   - Após o processamento, uma mensagem de sucesso será exibida no console e ele será encerrado automaticamente:
     - **Exemplo:** `"Arquivo Excel 'NomeDaPasta.xlsx' gerado com sucesso!"`

9. **Localização do arquivo Excel:**
   - O arquivo Excel será salvo na mesma pasta onde está o programa executável, sob o nome da pasta que você especificou seguido por _resultados.

---

## **Erros Comuns**

1. **Caminho inválido:**
   - Se o caminho fornecido para a pasta estiver incorreto ou não for encontrado, o programa exibirá a seguinte mensagem:  
     `"O caminho fornecido 'caminho_da_pasta' não é válido."`  
   - Verifique se o caminho está correto e tente novamente.

2. **PDFs fora do padrão:**
   - Se os arquivos PDF não seguirem o formato esperado, alguns parâmetros podem não ser corretamente identificados. Verifique se os PDFs têm a estrutura correta antes de processá-los.

3. **Outros arquivos .pdf na pasta**
   - Caso existam outros arquivos .pdf na pasta que não sigam o padrão estipulado nesse manual, seus nomes serão impressos no arquivo Excel, porém não terão dados atrelados a eles.
---

## **Apoio Técnico**
Se houver qualquer dúvida ou problema com o programa, entre em contato com o suporte técnico através do e-mail **leonardo@lmsdevs.com** ou no GitHub em **@leomsgit**.

---

### **Autor**
[Leonardo Miguel dos Santos](https://github.com/LeoMSgit)


