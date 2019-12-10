from anp import ANPWebScrap
from anp import captchaSolver,Excel
import time
import pysnooper
import _thread as thread
#from PIL import Image

def inicializar():
    anp = ANPWebScrap('https://preco.anp.gov.br/include/Resumo_Por_Estado_Index.asp')
    #anp.clickLink('ESTADO')    
    return anp
    
def navegarPagina(anp,estado,combustivel,tentativas):
    if tentativas <= MAX_TENTATIVAS:
        estado = estado
        combustivel = combustivel        
        anp.selecionarEstado(estado)    
        #anp.processar()
        #anp.processar()
        anp.selecionarCombustivel(combustivel)
        
        preencherCaptcha(anp,estado,combustivel,'captcha.png',tentativas)
       
    #anp.captcha()
def preencherCaptcha(anp,estado,combustivel,file,tentativas):
    anp.getCaptcha(file)    
    captcha = captchaSolver()
    captcha.read_ocr2(file)
    captcha.trim('screenshot-temp.png')
    captcha.resize('screenshot-temp3.png')    
    text = captcha.solveCaptcha2('screenshot-temp2.png')   
    #text = captcha.solveCaptcha2('screenshot-temp.png')   
    print(tentativas)   
    print(text)    
    anp.escreveCaptcha(text)

    #anp.processar()
    #time.sleep(2)
    print(anp.testaSucessoCaptcha())
    if anp.testaSucessoCaptcha() is False:
        if tentativas < MAX_TENTATIVAS: 
            tentativas = tentativas + 1
            navegarPagina(anp,estado,combustivel,tentativas)
    else:        
        return    
def teste():
    global estado
    global combustivel
    global anp
    global tentativas
    global planilha
    global MAX_TENTATIVAS
    MAX_TENTATIVAS = 10
    anp = inicializar()
    print(anp.testaSucessoCaptcha())
    #tentativas = 0
    #navegarPagina(anp,'Alagoas','Diesel',tentativas)
    #time.sleep(1)
    #anp.voltar() 
#@pysnooper.snoop()
def lerDadosGasolina():
    combustivel = "Gasolina"
    linhas = 3
    linhas2 = 2
    for linha in range(2,estados.max_row + 1):        
                try:
                    navegarPagina(anp,estados['B'+str(linha)].value,combustivel,tentativas) 
                except:
                    continue
                #print('Teste sucesso Captcha = ' + str(anp.testaSucessoCaptcha()))
                if anp.testaSucessoCaptcha() == True:
                    anp.lerTabela()
                    rows = anp.getLinhas()
                    #print(rows) 
                    lines = 0           
                    for row in rows:
                        #cells = row.getColunas()                    
                        
                        lines = lines + 1
                        if lines >= 4:
                            linhas = linhas + 1  
                            #print(row)                  
                            colsCidades = anp.getColunas(row)
                            colunas = [ele.text.strip() for ele in colsCidades]
                            #print(colunas)
                            s.acquire()
                            try:
                                
                                precoRegiao['A'+str(linhas-1)] = combustivel
                                precoRegiao['B'+str(linhas-1)] = estados['B'+str(linha)].value
                                precoRegiao['C'+str(linhas-1)] = str(colunas[0])
                                precoRegiao['D'+str(linhas-1)] = int(colunas[1])
                                precoRegiao['E'+str(linhas-1)] = float(colunas[2])
                                precoRegiao['F'+str(linhas-1)] = float(colunas[3])
                                precoRegiao['G'+str(linhas-1)] = float(colunas[4])
                                precoRegiao['H'+str(linhas-1)] = float(colunas[5])
                                precoRegiao['I'+str(linhas-1)] = float(colunas[6])
                                precoRegiao['J'+str(linhas-1)] = float(colunas[7])
                                precoRegiao['K'+str(linhas-1)] = float(colunas[8])
                                precoRegiao['L'+str(linhas-1)] = float(colunas[9])
                                precoRegiao['M'+str(linhas-1)] = float(colunas[10])
                            except:
                                s.release()
                                pass
                            planilha.salvarPlanilha('anp_saida.xlsx')
                            s.release()
                    lines = 0        
                    for row in rows:
                        lines = lines + 1
                        if lines >=4:
                            colsPostos = anp.getColunas(row)
                            colunas = [ele.text.strip() for ele in colsPostos]
                            try:
                                anp.clickLink(str(colunas[0]))
                            except:
                                continue
                                pass
                            time.sleep(1)
                            anp.lerTabela()
                            try:
                                linhasTabelas = anp.getLinhas()
                            except:
                                continue
                                pass    
                            #print(linhasTabelas)
                            lines1=0
                            for linha2 in linhasTabelas:
                                lines1 = lines1 + 1
                                if lines1 >=2: 
                                    #print(linha2)
                                    cols1 = anp.getColunas(linha2)
                                    #print(cols1)
                                    colunas1 = [ele.text.strip() for ele in cols1]
                                    #lines1 = 0
                                    #print(colunas1)
                                    #for linha1 in linhasTabelas:
                                    
                                    linhas2 = linhas2 + 1
                                    s.acquire() 
                                    try:
                                              
                                        precoPostos['A'+str(linhas2-1)] = combustivel
                                        precoPostos['B'+str(linhas2-1)] = estados['B'+str(linha)].value
                                        precoPostos['C'+str(linhas2-1)] = str(colunas[0])
                                        precoPostos['D'+str(linhas2-1)] = str(colunas1[0])
                                        precoPostos['E'+str(linhas2-1)] = str(colunas1[1])
                                        precoPostos['F'+str(linhas2-1)] = str(colunas1[2])
                                        precoPostos['G'+str(linhas2-1)] = str(colunas1[3])
                                        precoPostos['H'+str(linhas2-1)] = float(colunas1[4])
                                        precoPostos['I'+str(linhas2-1)] = float(colunas1[5])
                                        precoPostos['J'+str(linhas2-1)] = str(colunas1[6])
                                        precoPostos['K'+str(linhas2-1)] = str(colunas1[7])
                                        precoPostos['L'+str(linhas2-1)] = str(colunas1[8])
                                    except:
                                        s.release()
                                        pass
                                    #precoPostos['L'+str(linhas2-1)] = str(colunas1[9])
                                    #precoPostos['M'+str(lines-1)] = str(colunas1[9])
                                    planilha.salvarPlanilha('anp_saida.xlsx')
                                    s.release()
                            anp.voltar()
                    anp.voltar()            
def lerDadosEtanol():
    combustivel = "Etanol"
    linhas = 3
    linhas2 = 2
    for linha in range(2,estados.max_row + 1):        
                try:
                    navegarPagina(anp,estados['B'+str(linha)].value,combustivel,tentativas) 
                except:
                    continue
                #print('Teste sucesso Captcha = ' + str(anp.testaSucessoCaptcha()))
                if anp.testaSucessoCaptcha() == True:
                    anp.lerTabela()
                    rows = anp.getLinhas()
                    #print(rows) 
                    lines = 0           
                    for row in rows:
                        #cells = row.getColunas()                    
                        
                        lines = lines + 1
                        if lines >= 4:
                            linhas = linhas + 1  
                            #print(row)                  
                            colsCidades = anp.getColunas(row)
                            colunas = [ele.text.strip() for ele in colsCidades]
                            #print(colunas)
                            s.acquire()
                            try:
                                
                                precoRegiao['A'+str(linhas-1)] = combustivel
                                precoRegiao['B'+str(linhas-1)] = estados['B'+str(linha)].value
                                precoRegiao['C'+str(linhas-1)] = str(colunas[0])
                                precoRegiao['D'+str(linhas-1)] = int(colunas[1])
                                precoRegiao['E'+str(linhas-1)] = float(colunas[2])
                                precoRegiao['F'+str(linhas-1)] = float(colunas[3])
                                precoRegiao['G'+str(linhas-1)] = float(colunas[4])
                                precoRegiao['H'+str(linhas-1)] = float(colunas[5])
                                precoRegiao['I'+str(linhas-1)] = float(colunas[6])
                                precoRegiao['J'+str(linhas-1)] = float(colunas[7])
                                precoRegiao['K'+str(linhas-1)] = float(colunas[8])
                                precoRegiao['L'+str(linhas-1)] = float(colunas[9])
                                precoRegiao['M'+str(linhas-1)] = float(colunas[10])
                            except:
                                s.release()
                                pass
                            planilha.salvarPlanilha('anp_saida.xlsx')
                            s.release()
                    lines = 0        
                    for row in rows:
                        lines = lines + 1
                        if lines >=4:
                            colsPostos = anp.getColunas(row)
                            colunas = [ele.text.strip() for ele in colsPostos]
                            try:
                                anp.clickLink(str(colunas[0]))
                            except:
                                continue
                                pass
                            time.sleep(1)
                            anp.lerTabela()
                            try:
                                linhasTabelas = anp.getLinhas()
                            except:
                                continue
                                pass    
                            #print(linhasTabelas)
                            lines1=0
                            for linha2 in linhasTabelas:
                                lines1 = lines1 + 1
                                if lines1 >=2: 
                                    #print(linha2)
                                    cols1 = anp.getColunas(linha2)
                                    #print(cols1)
                                    colunas1 = [ele.text.strip() for ele in cols1]
                                    #lines1 = 0
                                    #print(colunas1)
                                    #for linha1 in linhasTabelas:
                                    
                                    linhas2 = linhas2 + 1
                                    s.acquire()
                                    try:       
                                        
                                        precoPostos['A'+str(linhas2-1)] = combustivel
                                        precoPostos['B'+str(linhas2-1)] = estados['B'+str(linha)].value
                                        precoPostos['C'+str(linhas2-1)] = str(colunas[0])
                                        precoPostos['D'+str(linhas2-1)] = str(colunas1[0])
                                        precoPostos['E'+str(linhas2-1)] = str(colunas1[1])
                                        precoPostos['F'+str(linhas2-1)] = str(colunas1[2])
                                        precoPostos['G'+str(linhas2-1)] = str(colunas1[3])
                                        precoPostos['H'+str(linhas2-1)] = float(colunas1[4])
                                        precoPostos['I'+str(linhas2-1)] = float(colunas1[5])
                                        precoPostos['J'+str(linhas2-1)] = str(colunas1[6])
                                        precoPostos['K'+str(linhas2-1)] = str(colunas1[7])
                                        precoPostos['L'+str(linhas2-1)] = str(colunas1[8])
                                    except:
                                        s.release()
                                        pass
                                    #precoPostos['L'+str(linhas2-1)] = str(colunas1[9])
                                    #precoPostos['M'+str(lines-1)] = str(colunas1[9])
                                    planilha.salvarPlanilha('anp_saida.xlsx')
                                    s.release()
                            anp.voltar()
                anp.voltar()            
def lerDados():

    linhas = 3
    linhas2 = 2
    for combustivel in combustiveis:
        for linha in range(2,estados.max_row + 1):        
            try:
                navegarPagina(anp,estados['B'+str(linha)].value,combustivel,tentativas) 
            except:
                continue
            #print('Teste sucesso Captcha = ' + str(anp.testaSucessoCaptcha()))
            if anp.testaSucessoCaptcha() == True:
                anp.lerTabela()
                rows = anp.getLinhas()
                #print(rows) 
                lines = 0           
                for row in rows:
                    #cells = row.getColunas()                    
                    
                    lines = lines + 1
                    if lines >= 4:
                        linhas = linhas + 1  
                        #print(row)                  
                        colsCidades = anp.getColunas(row)
                        colunas = [ele.text.strip() for ele in colsCidades]
                        #print(colunas)
                        try:
                            precoRegiao['A'+str(linhas-1)] = combustivel
                            precoRegiao['B'+str(linhas-1)] = estados['B'+str(linha)].value
                            precoRegiao['C'+str(linhas-1)] = str(colunas[0])
                            precoRegiao['D'+str(linhas-1)] = int(colunas[1])
                            precoRegiao['E'+str(linhas-1)] = float(colunas[2])
                            precoRegiao['F'+str(linhas-1)] = float(colunas[3])
                            precoRegiao['G'+str(linhas-1)] = float(colunas[4])
                            precoRegiao['H'+str(linhas-1)] = float(colunas[5])
                            precoRegiao['I'+str(linhas-1)] = float(colunas[6])
                            precoRegiao['J'+str(linhas-1)] = float(colunas[7])
                            precoRegiao['K'+str(linhas-1)] = float(colunas[8])
                            precoRegiao['L'+str(linhas-1)] = float(colunas[9])
                            precoRegiao['M'+str(linhas-1)] = float(colunas[10])
                        except:
                            pass
                        planilha.salvarPlanilha('anp_saida.xlsx')
                lines = 0        
                for row in rows:
                    lines = lines + 1
                    if lines >=4:
                        colsPostos = anp.getColunas(row)
                        colunas = [ele.text.strip() for ele in colsPostos]
                        try:
                            anp.clickLink(str(colunas[0]))
                        except:
                            continue
                            pass
                        time.sleep(1)
                        anp.lerTabela()
                        try:
                            linhasTabelas = anp.getLinhas()
                        except:
                            continue
                            pass    
                        #print(linhasTabelas)
                        lines1=0
                        for linha2 in linhasTabelas:
                            lines1 = lines1 + 1
                            if lines1 >=2: 
                                #print(linha2)
                                cols1 = anp.getColunas(linha2)
                                #print(cols1)
                                colunas1 = [ele.text.strip() for ele in cols1]
                                #lines1 = 0
                                #print(colunas1)
                                #for linha1 in linhasTabelas:
                                
                                linhas2 = linhas2 + 1
                                try:       
                                    precoPostos['A'+str(linhas2-1)] = combustivel
                                    precoPostos['B'+str(linhas2-1)] = estados['B'+str(linha)].value
                                    precoPostos['C'+str(linhas2-1)] = str(colunas[0])
                                    precoPostos['D'+str(linhas2-1)] = str(colunas1[0])
                                    precoPostos['E'+str(linhas2-1)] = str(colunas1[1])
                                    precoPostos['F'+str(linhas2-1)] = str(colunas1[2])
                                    precoPostos['G'+str(linhas2-1)] = str(colunas1[3])
                                    precoPostos['H'+str(linhas2-1)] = float(colunas1[4])
                                    precoPostos['I'+str(linhas2-1)] = float(colunas1[5])
                                    precoPostos['J'+str(linhas2-1)] = str(colunas1[6])
                                    precoPostos['K'+str(linhas2-1)] = str(colunas1[7])
                                    precoPostos['L'+str(linhas2-1)] = str(colunas1[8])
                                except:
                                    pass
                                #precoPostos['L'+str(linhas2-1)] = str(colunas1[9])
                                #precoPostos['M'+str(lines-1)] = str(colunas1[9])                                
                                planilha.salvarPlanilha('anp_saida.xlsx')
                        anp.voltar()
                anp.voltar()        

@pysnooper.snoop()
def lerDados2(combustivel):

    global tentativas
    global MAX_TENTATIVAS
    
    anp = inicializar()    
    MAX_TENTATIVAS = 10
    #anp = inicializar()
    tentativas = 0
    planilha = Excel('anp.xlsx')
    estados = planilha.lerAba('Estados')
    precoPostos = planilha.lerAba('Postos')
    precoRegiao = planilha.lerAba('Regiao')
    linhas = 3
    linhas2 = 2
    #for combustivel in combustiveis:
    for linha in range(2,estados.max_row + 1):        
        try:
            navegarPagina(anp,estados['B'+str(linha)].value,combustivel,tentativas) 
        except:
            continue
        #print('Teste sucesso Captcha = ' + str(anp.testaSucessoCaptcha()))
        if anp.testaSucessoCaptcha() == True:
            anp.lerTabela()
            rows = anp.getLinhas()
            #print(rows) 
            lines = 0           
            for row in rows:
                #cells = row.getColunas()                    
                
                lines = lines + 1
                if lines >= 4:
                    linhas = linhas + 1  
                    #print(row)                  
                    colsCidades = anp.getColunas(row)
                    colunas = [ele.text.strip() for ele in colsCidades]
                    #print(colunas)
                    try:
                        precoRegiao['A'+str(linhas-1)] = combustivel
                        precoRegiao['B'+str(linhas-1)] = estados['B'+str(linha)].value
                        precoRegiao['C'+str(linhas-1)] = str(colunas[0])
                        precoRegiao['D'+str(linhas-1)] = int(colunas[1])
                        precoRegiao['E'+str(linhas-1)] = float(colunas[2])
                        precoRegiao['F'+str(linhas-1)] = float(colunas[3])
                        precoRegiao['G'+str(linhas-1)] = float(colunas[4])
                        precoRegiao['H'+str(linhas-1)] = float(colunas[5])
                        precoRegiao['I'+str(linhas-1)] = float(colunas[6])
                        precoRegiao['J'+str(linhas-1)] = float(colunas[7])
                        precoRegiao['K'+str(linhas-1)] = float(colunas[8])
                        precoRegiao['L'+str(linhas-1)] = float(colunas[9])
                        precoRegiao['M'+str(linhas-1)] = float(colunas[10])
                    except:
                        pass
                    planilha.salvarPlanilha('anp_saida.xlsx')
            lines = 0        
            for row in rows:
                lines = lines + 1
                if lines >=4:
                    colsPostos = anp.getColunas(row)
                    colunas = [ele.text.strip() for ele in colsPostos]
                    try:
                        anp.clickLink(str(colunas[0]))
                    except:
                        continue
                        pass
                    time.sleep(1)
                    anp.lerTabela()
                    try:
                        linhasTabelas = anp.getLinhas()
                    except:
                        continue
                        pass    
                    #print(linhasTabelas)
                    lines1=0
                    for linha2 in linhasTabelas:
                        lines1 = lines1 + 1
                        if lines1 >=2: 
                            #print(linha2)
                            cols1 = anp.getColunas(linha2)
                            #print(cols1)
                            colunas1 = [ele.text.strip() for ele in cols1]
                            #lines1 = 0
                            #print(colunas1)
                            #for linha1 in linhasTabelas:
                            
                            linhas2 = linhas2 + 1
                            try:       
                                precoPostos['A'+str(linhas2-1)] = combustivel
                                precoPostos['B'+str(linhas2-1)] = estados['B'+str(linha)].value
                                precoPostos['C'+str(linhas2-1)] = str(colunas[0])
                                precoPostos['D'+str(linhas2-1)] = str(colunas1[0])
                                precoPostos['E'+str(linhas2-1)] = str(colunas1[1])
                                precoPostos['F'+str(linhas2-1)] = str(colunas1[2])
                                precoPostos['G'+str(linhas2-1)] = str(colunas1[3])
                                precoPostos['H'+str(linhas2-1)] = float(colunas1[4])
                                precoPostos['I'+str(linhas2-1)] = float(colunas1[5])
                                precoPostos['J'+str(linhas2-1)] = str(colunas1[6])
                                precoPostos['K'+str(linhas2-1)] = str(colunas1[7])
                                precoPostos['L'+str(linhas2-1)] = str(colunas1[8])
                            except:
                                pass
                            #precoPostos['L'+str(linhas2-1)] = str(colunas1[9])
                            #precoPostos['M'+str(lines-1)] = str(colunas1[9])                                
                            planilha.salvarPlanilha('anp_saida_'+combustivel+'.xlsx')
                    anp.voltar()
            anp.voltar()        
def main():
    #global estado
    #global combustivel
    #global anp
    #global tentativas
    #global planilha
    #global MAX_TENTATIVAS
    #global combustiveis
    #global estados
    #global precoPostos
    #global precoRegiao
    #global s

    #MAX_TENTATIVAS = 10
    #anp = inicializar()
    #tentativas = 0
    #planilha = Excel('anp.xlsx')
    #estados = planilha.lerAba('Estados')
    #precoPostos = planilha.lerAba('Postos')
    #precoRegiao = planilha.lerAba('Regiao')
    #combustiveis = ['Gasolina','Etanol','Diesel','DieselS10']
    
    s = thread.allocate_lock()
    thread.start_new_thread(lerDados2, ('Gasolina',))
    thread.start_new_thread(lerDados2, ('Etanol',))
    #while 1:pass
    #lerDados2(inicializar(),'Gasolina')
    #lerDados2('Etanol')

#teste()
main()