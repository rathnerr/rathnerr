import pandas as pd
from docxtpl import DocxTemplate
import os

# 1. Configurar caminhos (Garante que o Python ache os arquivos na pasta certa)
pasta_raiz = os.path.dirname(os.path.abspath(__file__))
caminho_excel = os.path.join(pasta_raiz, "dados.xlsx")
caminho_modelo = os.path.join(pasta_raiz, "modelo.docx")

def gerar_propostas():
    # 2. Ler os dados do Excel
    try:
        df = pd.read_excel(caminho_excel)
    except Exception as e:
        print(f"Erro ao ler Excel: {e}")
        return

    # 3. Carregar o modelo do Word
    try:
        template = DocxTemplate(caminho_modelo)
    except Exception as e:
        print(f"Erro ao abrir modelo Word: {e}")
        return

    # 4. Loop para processar cada linha do Excel
    for indice, linha in df.iterrows():
        # Transforma a linha do Excel em um "dicionário" para o Word
        contexto = {
            'cliente': linha['cliente'],
            'servico': linha['servico'],
            'valor': linha['valor'],
            'cor': linha['cor'],
            'genero': linha['genero']
        }
        
        # Preenche o modelo
        template.render(contexto)
        
        # Salva um arquivo individual para cada cliente
        nome_arquivo = f"Proposta_{linha['cliente']}.docx"
        caminho_salvamento = os.path.join(pasta_raiz, nome_arquivo)
        template.save(caminho_salvamento)
        
        print(f"✅ Gerada: {nome_arquivo}")

if __name__ == "__main__":
    gerar_propostas()