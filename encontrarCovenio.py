import pyautogui as py

def funcao_convenio():
    tentativas = 10
    max_tries = tentativas
    tries = 0
    if caixaSelecao.get() == "AMIL":
        global codPlano
        codPlano = 5
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/amil.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "ASSIM":
        codPlano = 7
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/assim.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click('descer')
    elif caixaSelecao.get() == "BR DISTRIBUIDORA":
        codPlano = 1678
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/amil.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "BRADESCO":
        codPlano = 27
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/bradesco.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "CAC":
        codPlano = 10
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/cac.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "CEF":
            codPlano = 1151
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/cef.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
    elif caixaSelecao.get() == "CORPO DE BOMBEIROS DO ESTADO DO RIO DE JANEIRO":
        codPlano = 2879
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/corpobomb.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "FUSMA":
        codPlano = 5383
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/fusma.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "GOLDEN CROSS":
        codPlano = 20
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/golden.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "MEDISERVICE":
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/mediservice.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
                codPlano = 24            
    elif caixaSelecao.get() == "PETROBRAS":
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/petrobras.png')
                py.click(x, y)
                break
            except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
                    codPlano = 26
    elif caixaSelecao.get() == "POSTAL SAUDE":
        codPlano = 4113
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/postal.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "SUL AMERICA":
            codPlano = 31
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/sulamerica.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
    elif caixaSelecao.get() == "UNIMED":
        codPlano = '4131'
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/unimed.png')
                py.click(x, y)
                break
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    else:
        print("Plano de saúde não encontrado")
