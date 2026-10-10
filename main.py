import requests
import pandas as pd
from datetime import datetime
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

print("🤖 Iniciando robô de automação e formatação...")

# 1. Puxando dados em tempo real da API
url = "https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL"
response = requests.get(url)

if response.status_code == 200:
    dados = response.json()

    lista_cotacoes = [
        {
            "Moeda": "Dólar (USD)",
            "Valor em BRL": float(dados["USDBRL"]["bid"]),
            "Variação (%)": float(dados["USDBRL"]["pctChange"]),
            "Data/Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        {
            "Moeda": "Euro (EUR)",
            "Valor em BRL": float(dados["EURBRL"]["bid"]),
            "Variação (%)": float(dados["EURBRL"]["pctChange"]),
            "Data/Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        {
            "Moeda": "Bitcoin (BTC)",
            "Valor em BRL": float(dados["BTCBRL"]["bid"]),
            "Variação (%)": float(dados["BTCBRL"]["pctChange"]),
            "Data/Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    ]

    nome_arquivo = "Relatorio_Cotacoes_Hoje.xlsx"

    # 2. Salvando com OpenPyXL para aplicar estilo visual
    df = pd.DataFrame(lista_cotacoes)

    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Cotações')

        # Acessando a aba do Excel para formatar
        worksheet = writer.sheets['Cotações']

        # Estilo para o Cabeçalho (Azul escuro com texto branco e negrito)
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Ajuste automático do tamanho de todas as colunas
        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            # Define a largura com uma folga para não prensar os textos
            worksheet.column_dimensions[col_letter].width = max_len + 5

            # Formata a coluna 'Valor em BRL' como Moeda (R$) e 'Variação' como Porcentagem
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
                row[1].number_format = 'R$ #,##0.00'  # Coluna do Valor
                row[2].number_format = '0.00%'  # Coluna da Variação

    print(f"✅ Sucesso! Planilha '{nome_arquivo}' gerada e formatada com sucesso.")
else:
    print("❌ Erro ao conectar com a API.")
