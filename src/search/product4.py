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
    ax.grid(True, which='both', linestyle='--', linewidth=0.5)
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
        else:
            unidade = "u"
        return str(int(quantidade)), unidade
    return None, None


def buscar_precos(driver, produto):
    """
    Busca o produto no Google Shopping, extrai os preços de forma dinâmica,
    filtra os resultados e retorna os preços min/max, além de uma lista detalhada
    de todos os itens encontrados para análise.
    """
    query = produto.replace(" ", "+")
    url = f"https://www.google.com/search?tbm=shop&q={query}"
    SELETOR_PRECO_ARIA = 'div[aria-label^="Current price"]'

    # NOVA LISTA para guardar todos os detalhes
    todos_os_itens_analisados = []

    try:
        driver.get(url)
        if "sorry/index" in driver.current_url:
            print("\n🚨 CAPTCHA DETECTADO! 🚨")
            input(
                "Resolva o CAPTCHA no navegador e, DEPOIS, pressione Enter aqui para continuar..."
            )
            print("Continuando a busca...")
            time.sleep(2)

        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, SELETOR_PRECO_ARIA))
        )

        html_content = driver.page_source
        soup = BeautifulSoup(html_content, "lxml")

        elementos_preco = soup.select(SELETOR_PRECO_ARIA)

        if not elementos_preco:
            print("  -> Nenhum elemento de preço encontrado com o seletor aria-label.")
            # MODIFICAÇÃO: Retorna a lista vazia junto com os outros valores
            return "Não encontrado", "Não encontrado", todos_os_itens_analisados

        qtd_original, unidade_original = extrair_unidade_e_quantidade(produto)
        print(
            f"  -> Padrão a ser buscado: Quantidade={qtd_original}, Unidade={unidade_original}"
        )

        for el_preco in elementos_preco:
            # container_produto = el_preco.find_parent().find_parent().find_parent()
            container_produto = el_preco.find_parent()
            if not container_produto:
                continue

            el_titulo = container_produto.find("div", {"title": True})
            if not el_titulo:
                continue

            nome_produto_encontrado = el_titulo["title"]
            preco_texto = el_preco["aria-label"]
            preco_limpo = re.search(r"[\d.,]+", preco_texto)

            link_produto = ""
            el_link = container_produto.find_parent("a")
            if el_link and "href" in el_link.attrs:
                link_produto = "https://www.google.com" + el_link["href"]

            if not preco_limpo:
                continue

            preco_float = float(preco_limpo.group().replace(".", "").replace(",", ""))/100

            if nome_produto_encontrado and preco_float:
                similaridade = jellyfish.jaro_winkler_similarity(
                    produto.lower(), nome_produto_encontrado.lower()
                )
                limiar = 0.01  # Mantido o limiar baixo para logar mais itens

                qtd_encontrada, unidade_encontrada = extrair_unidade_e_quantidade(nome_produto_encontrado)

                status_calculo = "Incluído"
                motivo_rejeicao = ""

                match_similaridade = (similaridade >= limiar)
                match_unidade = (qtd_original is None or (
                        qtd_original == qtd_encontrada and unidade_original == unidade_encontrada) or unidade_encontrada == "u")

                if not match_similaridade:
                    status_calculo = "Rejeitado (Relevância)"
                    motivo_rejeicao = f"Baixa similaridade com o termo '{produto}' ({similaridade:.2%})"
                elif not match_unidade:
                    status_calculo = "Rejeitado (Relevância)"
                    unidade_esp = f"{qtd_original}{unidade_original}" if qtd_original else "N/A"
                    unidade_enc = f"{qtd_encontrada}{unidade_encontrada}" if qtd_encontrada else "N/A"
                    motivo_rejeicao = f"Unidade/Qtd. divergente (Esperado: {unidade_esp}, Encontrado: {unidade_enc})"

                # Adiciona todos os itens encontrados à lista de análise
                todos_os_itens_analisados.append({
                    "Produto_Pesquisado": produto,
                    "Nome_Encontrado": nome_produto_encontrado,
                    "Preco": preco_float,
                    "Similaridade": f"{similaridade:.2%}",  # Formata como porcentagem
                    "Link": link_produto,
                    "Status_Calculo": status_calculo,
                    "Motivo_Rejeicao": motivo_rejeicao
                })

        # Filtra a lista de produtos que serão usados para o cálculo
        produtos_para_calculo = [p for p in todos_os_itens_analisados if p["Status_Calculo"] == "Incluído"]

        if not produtos_para_calculo:
            return "Não encontrado", "Não encontrado", todos_os_itens_analisados

        precos = [p["Preco"] for p in produtos_para_calculo]

        if len(precos) < 3:
            return min(precos), max(precos), todos_os_itens_analisados

        precos_array = np.array(precos)
        media = np.mean(precos_array)
        desvio_padrao = np.std(precos_array)
        print(f"  -> Análise Estatística: Média=R${media:.2f}, Desvio Padrão=R${desvio_padrao:.2f}")

        limite_inferior = media - (desvio_padrao * 1)
        limite_superior = media + (desvio_padrao * 1)

        precos_filtrados_final = []
        outliers_removidos_count = 0

        for item in produtos_para_calculo:
            preco_item = item["Preco"]
            if limite_inferior <= preco_item <= limite_superior:
                precos_filtrados_final.append(preco_item)
            else:
                # ATUALIZA O STATUS do item na lista principal 'todos_os_itens_analisados'
                item["Status_Calculo"] = "Rejeitado (Outlier)"
                item["Motivo_Rejeicao"] = (
                    f"Preço (R${preco_item:.2f}) fora do desvio padrão "
                    f"(Faixa aceitável: R${limite_inferior:.2f} - R${limite_superior:.2f})"
                )
                outliers_removidos_count += 1

        print(f"  -> {outliers_removidos_count} preço(s) removido(s) como outlier(s).")

        # Se *todos* os preços foram filtrados como outliers, retorna o min/max original (pré-filtro)
        if not precos_filtrados_final:
            print(
                "  -> Aviso: Todos os preços 'incluídos' foram removidos como outliers. Retornando min/max da faixa de relevância.")
            # Os status na lista 'todos_os_itens_analisados' já foram atualizados
            return min(precos), max(precos), todos_os_itens_analisados

        # 3. Retorna o min/max dos preços que *não* são outliers
        return min(precos_filtrados_final), max(precos_filtrados_final), todos_os_itens_analisados

    except TimeoutException:
        print(f"Tempo esgotado para '{produto}'. O seletor '{SELETOR_PRECO_ARIA}' não foi encontrado.")
        return "Timeout", "Timeout", todos_os_itens_analisados
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao buscar '{produto}': {e}")
        return "Erro inesperado", "Erro inesperado", todos_os_itens_analisados


# --- Início do Script Principal ---
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("--disable-blink-features=AutomationControlled")
# options.add_argument("--headless")
options.add_argument("--window-size=1920,1080")
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
)

driver = webdriver.Chrome(service=service, options=options)

try:
    df_entrada = pd.read_excel("../resources/produtos.xlsx")
    lista_produtos = df_entrada["Produto"].tolist()
    dados_resumo = []

    # NOVA LISTA para o arquivo de detalhes
    todos_os_dados_detalhados = []

    for produto in lista_produtos:
        print(f"\nBuscando preços para: '{produto}'...")
        # MODIFICAÇÃO: Captura a lista de itens analisados
        preco_min, preco_max, itens_analisados = buscar_precos(driver, produto)

        dados_resumo.append({"Produto": produto, "Preco_Minimo": preco_min, "Preco_Maximo": preco_max})

        # Adiciona os detalhes da busca atual à lista geral
        todos_os_dados_detalhados.extend(itens_analisados)

        if isinstance(preco_min, float) and isinstance(preco_max, float):
            print(f"  -> Resultado Final: Mínimo: R$ {preco_min:.2f} | Máximo: R$ {preco_max:.2f}")
        else:
            print(f"  -> Resultado Final: Preços não encontrados para '{produto}'")
        time.sleep(2)

    # --- SALVANDO OS DOIS ARQUIVOS EXCEL ---

    # Salva o arquivo de resumo como antes
    df_resumo = pd.DataFrame(dados_resumo)
    df_resumo.to_excel("precos_encontrados.xlsx", index=False)
    print("\n🎉 Processo concluído! Resultados salvos em 'precos_encontrados.xlsx'")

    # Salva o novo arquivo com todos os detalhes
    if todos_os_dados_detalhados:
        df_detalhes = pd.DataFrame(todos_os_dados_detalhados)
        df_detalhes.to_excel("itens_procurados.xlsx", index=False)
        print("🔍 Detalhes da busca salvos em 'itens_procurados.xlsx'")

finally:
    if driver:
        driver.quit()