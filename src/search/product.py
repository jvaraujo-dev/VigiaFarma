import pandas as pd
import re
import time
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

import numpy as np


def buscar_precos(driver, produto):
    """
    Busca o produto no Google Shopping, extrai os preços, filtra os outliers
    usando desvio padrão e retorna o mínimo e máximo dos preços restantes.
    """
    query = produto.replace(' ', '+')
    url = f"https://www.google.com/search?tbm=shop&q={query}"
    SELETOR_PRECO_ARIA = 'span[aria-label^="Current price:"]'

    try:
        driver.get(url)
        if 'sorry/index' in driver.current_url:
            print("\n🚨 CAPTCHA DETECTADO! 🚨")
            input("Resolva o CAPTCHA no navegador e, DEPOIS, pressione Enter aqui para continuar...")
            print("Continuando a busca...")
            time.sleep(2)

        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, SELETOR_PRECO_ARIA)))

        html_content = driver.page_source
        soup = BeautifulSoup(html_content, 'lxml')
        elementos_preco = soup.select(SELETOR_PRECO_ARIA)

        if not elementos_preco:
            return 'Não encontrado', 'Não encontrado'

        precos = []
        for el in elementos_preco:
            preco_texto = el.get_text(strip=True)
            preco_limpo = re.sub(r'[^\d,]', '', preco_texto).replace(',', '.')
            if preco_limpo:
                precos.append(float(preco_limpo)/100)

        if not precos:
            return 'Não encontrado', 'Não encontrado'

        # --- NOVO BLOCO DE FILTRAGEM ESTATÍSTICA ---

        # 1. Se tivermos poucos preços, não faz sentido calcular o desvio padrão.
        if len(precos) < 3:
            return min(precos), max(precos)

        # 2. Converter a lista de preços para um array NumPy para cálculos eficientes.
        precos_array = np.array(precos)

        # 3. Calcular a média e o desvio padrão.
        media = np.mean(precos_array)
        desvio_padrao = np.std(precos_array)

        print(f"  -> Análise Estatística: Média=R${media:.2f}, Desvio Padrão=R${desvio_padrao:.2f}")

        # 4. Definir os limites (vamos usar 1x o desvio padrão como padrão).
        limite_inferior = media - (desvio_padrao * 1)
        limite_superior = media + (desvio_padrao * 1)

        # 5. Filtrar a lista, mantendo apenas os preços dentro dos limites.
        precos_filtrados = [p for p in precos if limite_inferior <= p <= limite_superior]

        print(f"  -> {len(precos) - len(precos_filtrados)} preço(s) removido(s) como outlier(s).")

        # 6. Se o filtro removeu todos os preços, retornamos os originais como fallback.
        if not precos_filtrados:
            return min(precos), max(precos)

        # 7. Retornar o mínimo e máximo da lista JÁ FILTRADA.
        return min(precos_filtrados), max(precos_filtrados)

    except TimeoutException:
        # ... (código de erro continua o mesmo)
        print(f"Tempo esgotado para '{produto}'. O seletor '{SELETOR_PRECO_ARIA}' não foi encontrado.")
        return 'Seletor não encontrado', 'Seletor não encontrado'
    except Exception as e:
        # ... (código de erro continua o mesmo)
        print(f"Ocorreu um erro inesperado ao buscar '{produto}': {e}")
        return 'Erro inesperado', 'Erro inesperado'


# --- Início do Script Principal ---
# (As configurações iniciais do driver continuam as mesmas)
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--disable-blink-features=AutomationControlled")
# options.add_argument("--headless") # Mantenha comentado para depurar
options.add_argument("--window-size=1920,1080")
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36")

driver = webdriver.Chrome(service=service, options=options)

try:
    # O fluxo principal agora é simples e direto, sem a necessidade de descoberta dinâmica.
    df_entrada = pd.read_excel('../resources/produtos.xlsx')
    lista_produtos = df_entrada['Produto'].tolist()
    dados_finais = []

    for produto in lista_produtos:
        print(f"Buscando preços para: '{produto}'...")
        preco_min, preco_max = buscar_precos(driver, produto)
        dados_finais.append({
            'Produto': produto,
            'Preco_Minimo': preco_min,
            'Preco_Maximo': preco_max
        })
        print(f"  -> Mínimo: R$ {preco_min} | Máximo: R$ {preco_max}")
        # Aumentar um pouco a pausa pode ajudar a evitar CAPTCHAs futuros
        time.sleep(2)

    df_saida = pd.DataFrame(dados_finais)
    df_saida.to_excel('precos_encontrados.xlsx', index=False)
    print("\n🎉 Processo concluído! Resultados salvos em 'precos_encontrados.xlsx'")

finally:
    if driver:
        driver.quit()