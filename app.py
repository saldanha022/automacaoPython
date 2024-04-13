import openpyxl
import pyperclip #usado para copiar com acentos
import pyautogui
from time import sleep #importando comando para pausa

#-Entrando na planilha
planilha = openpyxl.load_workbook('produtos_ficticios.xlsx')
pagina_produtos = planilha['Produtos'] #Acessando pagina da planilha

#-Copiar informação de um campo e colar no seu campo correspondente
for linha in pagina_produtos.iter_rows(min_row=2): #percorrendo as linhas dentro da planilha a partir da linha 2
    
    #Nome Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto) #copiando o valor da tabela
    pyautogui.click(866,216, duration=1) #entrando no campo onde deve ser colado atraves da coordenada e definindo o tempo
    pyautogui.hotkey('ctrl', 'v') #passando um atalho e colando a informação
    pyautogui.hotkey('tab') #tab para ir para o proximo camp

    #Descricao
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Codigo Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.hotkey('ctrl', 'v') #sem tab pois proximo passo exige click


    #Clicando no botao de proximo
    pyautogui.click(857,710)
    sleep(2) #dando uma pausa de 3 segundos
    
    #Preço
    preco = linha[6].value
    pyautogui.hotkey('tab') #entrando no campo
    pyperclip.copy(preco)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')


    #Qtd em Estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.hotkey('ctrl', 'v') #Nao tem o tab pois o tamanhos precisa clicar

    #Abrindo menu tamanhos
    pyautogui.click(851,559, duration= 1)
    tamanho = linha[10].value
    pyperclip.copy(tamanho)

    #Checando o tamanho para selecionar a opçao
    if tamanho == 'Pequeno':
        pyautogui.click(846,574, duration= 1)
        pyautogui.hotkey('tab')
    elif tamanho == 'Grande':
        pyautogui.click(843,600, duration= 1)
        pyautogui.hotkey('tab')
    else:
        pyautogui.click(857,615, duration= 1)
        pyautogui.hotkey('tab')

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.hotkey('ctrl', 'v')

    #Clicando no Botao Proximo
    pyautogui.click(836,688, duration=1)
    sleep(3)

    #Fabricante
    fabricante = linha[12].value
    pyautogui.hotkey('tab') #entrando no campo
    pyperclip.copy(fabricante)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Pais Origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Observaçoes
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Codigo de Barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    #Localização armazem
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.hotkey('ctrl', 'v') #sem tab, encerramento

    #Botao Concluir
    pyautogui.click(855,674, duration=1)

    #Botao Ok
    pyautogui.click(1219,166, duration=1)
    sleep(2)

    #Botao add mais um
    pyautogui.click(1021,470, duration=1)
