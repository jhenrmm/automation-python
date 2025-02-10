import openpyxl
import pyperclip
import pyautogui

workbook = openpyxl.load_workbook(
    r'C:\Users\João Henrique\Desktop\Projetos\Freelancer\2\produtos_ficticios.xlsx')

sheet_produtos = workbook['Produtos']
for linha in sheet_produtos.iter_rows(min_row=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(356, 339, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(348, 423, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(353, 562, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo = linha[3].value
    pyperclip.copy(codigo)
    pyautogui.click(347, 644, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(351, 734, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensao = linha[5].value
    pyperclip.copy(dimensao)
    pyautogui.click(358, 819, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(365, 884, duration=1)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(352, 366, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.click(365, 449, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_val = linha[8].value
    pyperclip.copy(data_de_val)
    pyautogui.click(360, 535, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(366, 615, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(343, 708, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(363, 735, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(373, 774, duration=1)
    else:
        pyautogui.click(364, 801, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(350, 795, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(367, 852, duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(355, 384, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_de_origem = linha[13].value
    pyperclip.copy(pais_de_origem)
    pyautogui.click(352, 464, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(354, 553, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barra = linha[15].value
    pyperclip.copy(codigo_de_barra)
    pyautogui.click(354, 689, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    local_no_armazen = linha[16].value
    pyperclip.copy(local_no_armazen)
    pyautogui.click(351, 773, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(366, 835, duration=1)
    pyautogui.click(1153, 162, duration=1)
    pyautogui.click(992, 613, duration=1)
