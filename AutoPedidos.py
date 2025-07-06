import time
import pyautogui
import pyperclip
import pandas

# Lê a planilha
tabela = pandas.read_csv(r"C:\Users\Stella C. Aranha\Downloads\AUTOMAÇÃO\pedidos.csv")

time.sleep(2)

# Define os 3 padrões de posição"58640-000"
posicoes = [
    {  # Posição 1
        'nome': (707, 252),
        'endereco': (722, 266),
        'cidade': (733, 279),
        'estado': (845, 277),
        'cep': (707, 287),
        'valor': (827, 316),
        'qtd': (766, 315),
        'conteudo': (589, 316),
    },
    {  # Posição 2
        'nome': (720, 423),
        'endereco': (719, 434),
        'cidade': (741, 450),
        'estado': (858, 448),
        'cep': (730, 462),
        'valor': (837, 488),
        'qtd': (763, 487),
        'conteudo': (566, 489),
    },
    {  # Posição 3
        'nome': (731, 593),
        'endereco': (741, 607),
        'cidade': (732, 622),
        'estado': (852, 623),
        'cep': (714, 630),
        'valor': (822, 659),
        'qtd': (760, 660),
        'conteudo': (551, 659),
    },
]

# Dicionário de estados
estados = {
    "Acre": "AC", "Alagoas": "AL", "Amapá": "AP", "Amazonas": "AM", "Bahia": "BA",
    "Ceará": "CE", "Distrito Federal": "DF", "Espírito Santo": "ES", "Goiás": "GO",
    "Maranhão": "MA", "Mato Grosso": "MT", "Mato Grosso do Sul": "MS", "Minas Gerais": "MG",
    "Pará": "PA", "Paraíba": "PB", "Paraná": "PR", "Pernambuco": "PE", "Piauí": "PI",
    "Rio de Janeiro": "RJ", "Rio Grande do Norte": "RN", "Rio Grande do Sul": "RS",
    "Rondônia": "RO", "Roraima": "RR", "Santa Catarina": "SC", "São Paulo": "SP",
    "Sergipe": "SE", "Tocantins": "TO"
}

# Contador para nome do arquivo
contador_doc = 1

# Abrir Word no início
pyautogui.click(410, 758)
time.sleep(3)

# Loop das linhas
for linha in tabela.index:

    grupo = linha % 3
    pos = posicoes[grupo]

    # Nome
    nome = tabela.loc[linha, "Nome do destinatário"]
    pyautogui.click(*pos['nome'])
    pyperclip.copy(nome)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Endereço
    endereco = tabela.loc[linha, "Endereço de entrega"].strip()
    pyautogui.click(*pos['endereco'])
    pyperclip.copy(endereco)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Cidade
    cidade = tabela.loc[linha, "Cidade de entrega"]
    pyautogui.click(*pos['cidade'])
    pyperclip.copy(cidade)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Estado
    estado = tabela.loc[linha, "Nome do estado de entrega"]
    estado_sigla = estados.get(estado, '??')
    pyautogui.click(*pos['estado'])
    pyperclip.copy(estado_sigla)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # CEP
    cep = tabela.loc[linha, "Código postal (CEP) de entrega"]
    pyautogui.click(*pos['cep'])
    pyperclip.copy(cep)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Valor
    valor = tabela.loc[linha, "Preço"]
    pyautogui.click(*pos['valor'])
    pyperclip.copy(valor)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Quantidade
    qtd = tabela.loc[linha, "Quant. total do pedido"]
    pyautogui.click(*pos['qtd'])
    pyperclip.copy(qtd)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    # Conteúdo
    # Lê o conteúdo original
    conteudo_original = tabela.loc[linha, "Item"].lower()

    if "primavera-verão" in conteudo_original:
        conteudo = "livros: dia de primavera e um verão entre nós"
    elif "um verão entre nós" in conteudo_original:
        conteudo = "livro: um verão entre nós"
    elif "dia de primavera" in conteudo_original:
        conteudo = "livro: dia de primavera"
    else:
        conteudo = conteudo_original

    pyautogui.click(*pos['conteudo'])
    pyperclip.copy(conteudo)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    # Se for a 3ª posição do grupo (linha % 3 == 2), salvar e abrir novo documento
    if grupo == 2:
        # SALVAR como "declaração {contador_doc}"
        pyautogui.hotkey('ctrl', 's')
        time.sleep(3)
        nome_arquivo = f"declaração {contador_doc}"
        pyperclip.copy(nome_arquivo)
        pyautogui.hotkey('f12')
        time.sleep(3)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(4)
        for _ in range(7):  # Repete 7 vezes
            pyautogui.hotkey('tab')
            time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(3)

        # FECHAR Word
        pyautogui.hotkey('alt', 'f4')
        time.sleep(2)

        # ABRIR NOVO Word
        import os
        # Caminho completo até o arquivo

        caminho_arquivo = r"C:\Users\Stella C. Aranha\Downloads\Nova pasta\modelocont.docx"

        # Abrir o arquivo com o programa padrão (Word)
        os.startfile(caminho_arquivo)
        time.sleep(4)

        contador_doc += 1
