import requests
import json
import pandas as pd
import os
from bs4 import BeautifulSoup
from urllib.parse import quote
from datetime import datetime

# filename = list(uploaded.keys())[0]
# print(uploaded.keys())

filename = input("Digite o nome do arquivo (com a extensão): ")

df_uploaded = pd.read_excel(filename)
print(f"Programa iniciado as {datetime.now()}")

#display(df_uploaded.head())

rotulos_de_linha = df_uploaded['Rótulos de Linha'].tolist()

parametros_ignorados = ['(vazio)', 'Total Geral']

LISTA_MEDICAMENTOS = [item for item in rotulos_de_linha if item not in parametros_ignorados]

print("\nProcessed list of medications:")
print(LISTA_MEDICAMENTOS)

# URL da API em Graphql
API_URL = "https://qp1crcg3c6.execute-api.us-east-1.amazonaws.com/production/b2c/graphql"

# Headers da API
HEADERS = {
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Content-Type": "application/json",
    "Origin": "https://farmaindex.com",
    "Referer": "https://farmaindex.com/",
    "sec-ch-ua": "\"Not;A=Brand\";v=\"99\", \"Google Chrome\";v=\"139\", \"Chromium\";v=\"139\"",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "\"Linux\"",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "cross-site",
    "user-agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36"
}

# Query GraphQL
GRAPHQL_QUERY = """
query searchPrefix($q: String!) {
  searchPrefix(q: $q) {
    medicamentoid
    medicamento
    apresentacao
    farmaco
    laboratorio
    tipoid
    tipo
    imagem
    imagemUrl
    tarja
    oferta
    url
    preco
    __typename
  }
}
"""

def extrair_dados_medicamento(url: str):
    """
    Busca o HTML de uma URL, encontra os dados estruturados (JSON-LD)
    e extrai as informações do medicamento.
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        script_tag = soup.find('script', type='application/ld+json')
        if not script_tag:
            return None
        dados_json = json.loads(script_tag.string)
        id_alvo = quote(f"{url}#drug", safe='/:#?=')
        # print(f"\nDados puros: {dados_json}")
        for item in dados_json.get('@graph', []):
            # print(f"\nItem puro: {item}")
            if isinstance(item, dict) and item.get('@id') == id_alvo:
                return item
        return None
    except Exception:
        return None

def buscar_medicamento(nome_medicamento: str):
    """
    Função que envia uma requisição para a API GraphQL e retorna os resultados.
    """
    apresentacao_buscada = []
    # print(f"Buscando por: '{nome_medicamento}'...")
    payload = {
        "operationName": "searchPrefix",
        "variables": {"q": nome_medicamento},
        "query": GRAPHQL_QUERY
    }
    try:
        response = requests.post(API_URL, headers=HEADERS, json=payload, timeout=10)
        response.raise_for_status()
        dados = response.json().get('data', {}).get('searchPrefix', [])
        dados_filtrados = []
        for dado in dados:
            apresentacao = dado.get('apresentacao')
            tem_oferta = dado.get('preco') is not None
            if apresentacao not in apresentacao_buscada and tem_oferta:
                dados_filtrados.append(dado)
                apresentacao_buscada.append(apresentacao)
        return dados_filtrados
    except requests.exceptions.RequestException as e:
        print(f"Ocorreu um erro na requisição: {e}")
        return []
    except json.JSONDecodeError:
        print("Erro ao decodificar a resposta JSON.")
        return []


def mostra_medicamentos_encontrados():
    global i, med
    print(f"\nForam encontrados {len(medicamentos_encontrados)} resultados. Exibindo os {limite_exibicao} primeiros:\n")
    for i, med in enumerate(medicamentos_encontrados[:limite_exibicao]):
        print(f"--- Opção {i + 1} ---")
        print(f"  Nome: {med.get('medicamento')}")
        print(f"  Apresentação: {med.get('apresentacao')}")
        print(f"  Laboratório: {med.get('laboratorio')}")
        print("-" * 20)


def salva_excel():
    decisao = "s"
    if decisao.lower().strip().startswith('s'):
        # --- BLOCO DE FORMATAÇÃO E SALVAMENTO ---

        nome_arquivo = "base_farmaindex.xlsx"
        nome_aba = "base"

        df_novo_dado = pd.DataFrame([dados_completos])

        if os.path.exists(nome_arquivo):
            # print(f"\nArquivo '{nome_arquivo}' encontrado. Adicionando dados...")
            df_existente = pd.read_excel(nome_arquivo)
            df_final = pd.concat([df_existente, df_novo_dado], ignore_index=True)
        else:
            # print(f"\nCriando novo arquivo '{nome_arquivo}'...")
            df_final = df_novo_dado

        colunas_numericas = ['Preço Referência', 'Maior Preço Encontrado', 'Menor Preço Encontrado', 'Registro MS']
        for col in colunas_numericas:
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

        # Garante que as colunas de moeda sejam arredondadas para 2 casas decimais
        colunas_moeda = ['Preço Referência', 'Maior Preço Encontrado', 'Menor Preço Encontrado']
        for col in colunas_moeda:
            if col in df_final.columns:
                df_final[col] = df_final[col].round(2)

        # Usa o ExcelWriter para salvar e formatar
        with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, sheet_name=nome_aba, index=False)

            workbook = writer.book
            worksheet = writer.sheets[nome_aba]

            # Cria os formatos de célula
            formato_moeda = workbook.add_format({'num_format': '0.00'})
            formato_ms = workbook.add_format({'num_format': '0'})

            try:
                headers = list(df_final.columns)
                col_preco_ref = headers.index('Preço Referência')
                col_maior_preco = headers.index('Maior Preço Encontrado')
                col_menor_preco = headers.index('Menor Preço Encontrado')
                col_ms = headers.index('Registro MS')

                # Aplica a formatação e define a largura das colunas
                worksheet.set_column(col_preco_ref, col_preco_ref, 18, formato_moeda)
                worksheet.set_column(col_maior_preco, col_maior_preco, 22, formato_moeda)
                worksheet.set_column(col_menor_preco, col_menor_preco, 22, formato_moeda)
                worksheet.set_column(col_ms, col_ms, 18, formato_ms)
                # Ajusta a largura das outras colunas para melhor visualização
                worksheet.set_column(headers.index('Nome'), headers.index('Nome'), 25)
                worksheet.set_column(headers.index('Apresentação'), headers.index('Apresentação'), 40)
                worksheet.set_column(headers.index('Laboratório'), headers.index('Laboratório'), 25)
            except ValueError:
                print("Aviso: Não foi possível formatar uma ou mais colunas.")

        # print("Dados salvos e formatados com sucesso!")

    else:
        print("Operação cancelada. Encerrando.")


def mostra_detalhes_do_medicamento():
    print("\n--- Detalhes do Produto Selecionado ---")
    for chave, valor in dados_completos.items():
        print(f"  {chave}: {valor}")
    print("-" * 35)


# EXECUÇÃO PRINCIPAL DO SCRIPT
if __name__ == "__main__":

    for i,med in enumerate(LISTA_MEDICAMENTOS):
        NOME_DO_MEDICAMENTO_A_BUSCAR = med
        medicamentos_encontrados = buscar_medicamento(NOME_DO_MEDICAMENTO_A_BUSCAR)

        if medicamentos_encontrados:

          # limite_exibicao = 10
          limite_exibicao = len(medicamentos_encontrados)

          # mostra_medicamentos_encontrados()

          try:
              for med_a_salvar in medicamentos_encontrados:

                  med_selecionado = med_a_salvar

                  nome_para_url = med_selecionado.get('medicamento').lower().replace(' ', '-')
                  id_medicamento = str(med_selecionado.get('medicamentoid'))
                  URL_MEDICAMENTO = f"https://farmaindex.com/{nome_para_url}/{id_medicamento}"

                  dados_do_remedio = extrair_dados_medicamento(URL_MEDICAMENTO)

                  dados_completos = {
                      'Nome': med_selecionado.get('medicamento'),
                      'Apresentação': med_selecionado.get('apresentacao'),
                      'Laboratório': med_selecionado.get('laboratorio'),
                      'Preço Referência': med_selecionado.get('preco'),
                      'Registro MS': 'Não encontrado',
                      'Maior Preço Encontrado': 'N/A',
                      'Menor Preço Encontrado': 'N/A'
                  }

                  if dados_do_remedio:
                      identifier = dados_do_remedio.get('identifier', {})
                      registro = identifier.get('value')
                      offers = dados_do_remedio.get('offers', {})
                      maior_preco = offers.get('highPrice')
                      menor_preco = offers.get('lowPrice')

                      dados_completos['Registro MS'] = registro
                      dados_completos['Maior Preço Encontrado'] = maior_preco
                      dados_completos['Menor Preço Encontrado'] = menor_preco

                  # mostra_detalhes_do_medicamento()

                  salva_excel()

          except (ValueError, IndexError) as e:
              print(f"\nErro: {e}. Encerrando o programa.")
        else:
            print(f"\nNão foram encontrados resultados para {NOME_DO_MEDICAMENTO_A_BUSCAR}")
    print(f"\nPrograma finalizado as {datetime.now()}")