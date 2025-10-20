import pandas as pd
import re
import time
from bs4 import BeautifulSoup
import jellyfish

# SELENIUM
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

# Grafico
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker


def geraGrafico(precos_array):
    # Gera índices para o eixo X, apenas para posicionar os pontos
    x_indices = np.arange(len(precos_array))
    # 2. Realizar os mesmos cálculos que o seu script faz
    media = np.mean(precos_array)
    desvio_padrao = np.std(precos_array)
    # Este é o fator multiplicador que você pode ajustar
    fator_desvio = 1.0
    limite_inferior = media - (desvio_padrao * fator_desvio)
    limite_superior = media + (desvio_padrao * fator_desvio)
    # 3. Separar os preços em "mantidos" e "removidos" (outliers)
    precos_mantidos_mask = (precos_array >= limite_inferior) & (
        precos_array <= limite_superior
    )
    outliers_mask = ~precos_mantidos_mask
    # 4. Criar o Gráfico
    plt.style.use("seaborn-v0_8-whitegrid")
    fig, ax = plt.subplots(figsize=(12, 7))
    # Plotar a média e os limites do desvio padrão
    ax.axhline(
        y=media,
        color="orangered",
        linestyle="-",
        linewidth=2,
        label=f"Média: R${media:.2f}",
    )
    ax.axhline(
        y=limite_superior,
        color="gold",
        linestyle="--",
        linewidth=2,
        label=f"Limite Superior (+{fator_desvio}σ): R${limite_superior:.2f}",
    )
    ax.axhline(
        y=limite_inferior,
        color="gold",
        linestyle="--",
        linewidth=2,
        label=f"Limite Inferior (-{fator_desvio}σ): R${limite_inferior:.2f}",
    )
    ax.fill_between(
        x_indices,
        limite_inferior,
        limite_superior,
        color="gold",
        alpha=0.1,
        label="Faixa de Normalidade",
    )
    # Plotar os pontos de preço
    ax.scatter(
        x_indices[precos_mantidos_mask],
        precos_array[precos_mantidos_mask],
        color="royalblue",
        s=80,
        zorder=5,
        label="Preços Mantidos",
    )
    ax.scatter(
        x_indices[outliers_mask],
        precos_array[outliers_mask],
        color="red",
        s=80,
        zorder=5,
        edgecolor="black",
        marker="X",
        label="Outliers Removidos",
    )
    # 5. Melhorar a aparência do gráfico
    ax.set_title("Visualização do Filtro por Desvio Padrão", fontsize=18, pad=20)
    ax.set_xlabel("Produtos Encontrados (ilustrativo)", fontsize=12)
    ax.set_ylabel("Preço (R$)", fontsize=12)
    ax.legend(fontsize=11)
    ax.grid(True, which="both", linestyle="--", linewidth=0.5)
    # Formatar o eixo Y para mostrar o símbolo R$
    formatter = mticker.FormatStrFormatter("R$%.2f")
    ax.yaxis.set_major_formatter(formatter)
    # Oculta os números do eixo X, pois são apenas para posicionamento
    plt.xticks([])
    plt.tight_layout()
    # Exibe o gráfico em uma janela
    plt.show()


def extrair_unidade_e_quantidade(texto_produto):
    """Extrai a unidade de medida e a quantidade de um texto."""
    if not isinstance(texto_produto, str):
        return None, None
    texto_produto = texto_produto.lower()
    match = re.search(r"(\d[\d.,]*)\s?(kg|g|ml|l|caps)\b", texto_produto)
    if match:
        quantidade = float(match.group(1).replace(",", "."))
        unidade = match.group(2)
        if unidade == "kg":
            quantidade *= 1000
            unidade = "g"
        elif unidade == "l":
            quantidade *= 1000
            unidade = "ml"
        return str(int(quantidade)), unidade
    return None, None


def buscar_precos(driver, produto):
    """
    Busca o produto no Google Shopping, extrai os preços de forma dinâmica e robusta,
    filtra por similaridade e outliers, e retorna o mínimo e máximo.
    """
    query = produto.replace(" ", "+")
    url = f"https://www.google.com/search?tbm=shop&q={query}"
    # Seletor genérico para aguardar o carregamento inicial da página
    SELETOR_AGUARDAR = "body"

    try:
        driver.get(url)
        if "sorry/index" in driver.current_url:
            print("\n🚨 CAPTCHA DETECTADO! 🚨")
            input(
                "Resolva o CAPTCHA no navegador e, DEPOIS, pressione Enter aqui para continuar..."
            )
            print("Continuando a busca...")
            time.sleep(2)

        # Espera o corpo da página carregar, é o suficiente
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, SELETOR_AGUARDAR))
        )

        html_content = driver.page_source
        soup = BeautifulSoup(html_content, "lxml")

        # --- NOVA LÓGICA DE EXTRAÇÃO DINÂMICA ---
        produtos_encontrados = []
        # Encontra todos os containers de produtos baseados em um atributo 'title' que contenha texto.
        # Esta é uma abordagem muito mais robusta do que usar classes.
        containers_produtos = soup.select("div[title]:has(img)")

        if not containers_produtos:
            print("  -> Nenhum container de produto encontrado na página.")
            return "Não encontrado", "Não encontrado"

        qtd_original, unidade_original = extrair_unidade_e_quantidade(produto)
        print(
            f"  -> Padrão a ser buscado: Quantidade={qtd_original}, Unidade={unidade_original}"
        )

        for container in containers_produtos:
            nome_produto_encontrado = container.get("title")

            # Dentro do container, encontra o preço buscando pelo texto "R$"
            preco_tag = container.find(
                lambda tag: tag.name in ['span', 'div'] and re.search(r"R\$\s?[\d.,]+", tag.get_text())
            )
            if not preco_tag:
                continue

            # Pega o texto completo do elemento encontrado e limpa para extrair apenas os números
            preco_texto = preco_tag.strip()
            preco_limpo = re.search(r"[\d.,]+", preco_texto)
            if not preco_limpo:
                continue

            preco_float = float(preco_limpo.group().replace(".", "").replace(",", "."))

            if nome_produto_encontrado and preco_float:
                # 1. Validação por Jaro-Winkler
                similaridade = jellyfish.jaro_winkler_similarity(
                    produto.lower(), nome_produto_encontrado.lower()
                )
                limiar = 0.8  # Ajuste este valor conforme a necessidade

                if similaridade >= limiar:
                    # 2. Validação de Quantidade e Unidade
                    (
                        qtd_encontrada,
                        unidade_encontrada,
                    ) = extrair_unidade_e_quantidade(nome_produto_encontrado)

                    # Compara se a quantidade/unidade bate com a do produto original
                    if (
                        qtd_original is None
                        or (qtd_original == qtd_encontrada and unidade_original == unidade_encontrada)
                    ):
                        produtos_encontrados.append(
                            {"nome": nome_produto_encontrado, "preco": preco_float}
                        )
                        print(
                            f"  -> Match: '{nome_produto_encontrado}' (Preço: R${preco_float:.2f}, Similaridade: {similaridade:.2f})"
                        )
                    else:
                        print(
                            f"  -> Ignorado (Qtd/Unid): '{nome_produto_encontrado}' (Encontrado: {qtd_encontrada}{unidade_encontrada})"
                        )
                else:
                    print(
                        f"  -> Ignorado (Nome): '{nome_produto_encontrado}' (Similaridade: {similaridade:.2f})"
                    )

        if not produtos_encontrados:
            return "Não encontrado", "Não encontrado"

        precos = [p["preco"] for p in produtos_encontrados]

        # --- BLOCO DE FILTRAGEM ESTATÍSTICA (sem alterações) ---
        if len(precos) < 3:
            return min(precos), max(precos)

        precos_array = np.array(precos)
        media = np.mean(precos_array)
        desvio_padrao = np.std(precos_array)

        print(
            f"  -> Análise Estatística: Média=R${media:.2f}, Desvio Padrão=R${desvio_padrao:.2f}"
        )

        limite_inferior = media - (desvio_padrao * 1)
        limite_superior = media + (desvio_padrao * 1)
        precos_filtrados = [
            p for p in precos if limite_inferior <= p <= limite_superior
        ]

        print(
            f"  -> {len(precos) - len(precos_filtrados)} preço(s) removido(s) como outlier(s)."
        )

        if not precos_filtrados:
            return min(precos), max(precos)

        return min(precos_filtrados), max(precos_filtrados)

    except TimeoutException:
        print(f"Tempo esgotado para '{produto}'. A página pode não ter carregado a tempo.")
        return "Timeout", "Timeout"
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao buscar '{produto}': {e}")
        return "Erro inesperado", "Erro inesperado"


# --- Início do Script Principal ---
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("--disable-blink-features=AutomationControlled")
# options.add_argument("--headless")  # Executa navegador em segundo plano
options.add_argument("--window-size=1920,1080")
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
)

driver = webdriver.Chrome(service=service, options=options)

try:
    # ATENÇÃO: Verifique o caminho para o seu arquivo Excel.
    df_entrada = pd.read_excel("../resources/produtos.xlsx")
    lista_produtos = df_entrada["Produto"].tolist()
    dados_finais = []

    for produto in lista_produtos:
        print(f"\nBuscando preços para: '{produto}'...")
        preco_min, preco_max = buscar_precos(driver, produto)
        dados_finais.append(
            {"Produto": produto, "Preco_Minimo": preco_min, "Preco_Maximo": preco_max}
        )
        if isinstance(preco_min, float) and isinstance(preco_max, float):
            print(f"  -> Resultado Final: Mínimo: R$ {preco_min:.2f} | Máximo: R$ {preco_max:.2f}")
        else:
            print(f"  -> Resultado Final: Preços não encontrados para '{produto}'")
        time.sleep(2)  # Mantém uma pausa para evitar bloqueios

    df_saida = pd.DataFrame(dados_finais)
    df_saida.to_excel("precos_encontrados.xlsx", index=False)
    print("\n🎉 Processo concluído! Resultados salvos em 'precos_encontrados.xlsx'")

finally:
    if driver:
        driver.quit()