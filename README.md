README.md

# Preenchimento Automatizado 

Este projeto tem como objetivo automatizar o preenchimento de documentos com base em uma planilha de pedidos, 
utilizando Python e integração com o Microsoft Word.

## Funcionalidades

- Leitura automática de dados a partir de uma planilha (Excel).
- Preenchimento dinâmico de documentos do Word (.docx) com os dados da planilha.
- Redução significativa do tempo gasto em tarefas manuais e repetitivas.

## Tecnologias utilizadas

- Python 3.x
- Bibliotecas:
  - pandas: leitura e acesso aos dados da planilha .csv
  - pyautogui: simulação de cliques, atalhos de teclado e interação com a tela
  - pyperclip: cópia de dados para a área de transferência (clipboard)
  - time: controle de pausas entre ações automatizadas
  - os: abertura automática de novos documentos Word com base em um modelo

## ⚠️ Observações

> A planilha real de pedidos *não foi incluída neste repositório* por conter dados sensíveis.  
> Se quiser testar o funcionamento, substitua por uma planilha fictícia com os mesmos campos usados no código.

## 🛠️ Como executar o projeto

1. Clone o repositório:
   ```bash
   git clone https://github.com/StellaCabanelas/Preenchimento-Automatizado

2. Instale as bibliotecas necessárias:
pip install pandas openpyxl python-docx

3. Execute o script principal:
python Autopedidos.py
