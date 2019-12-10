import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.webdriver.support.ui as ui
from selenium.webdriver.support.ui import Select
import ctypes
from PIL import Image, ImageFilter,ImageEnhance, ImageChops
from pytesseract import *
from io import BytesIO
import os
import cv2 as cv
import numpy as np
from matplotlib import pyplot as plt
import openpyxl
import sys
import re

class captchaSolver:
    def __init__(self):
        pass
    def trim(self,img):
        self.img = img
        im = Image.open(self.img)
        bg = Image.new(im.mode, im.size, im.getpixel((0,0)))
        diff = ImageChops.difference(im, bg)
        diff = ImageChops.add(diff, diff, 2.0, -100)
        bbox = diff.getbbox()
        if bbox:
            im.crop(bbox)
            im.save("screenshot-temp3.png") 
    def resize(self,file):
        self.file = file 
        img = cv.imread(self.file, cv.IMREAD_UNCHANGED) 
        
        scale_percent = 300 # percent of original size
        width = int(img.shape[1] * scale_percent / 100)
        height = int(img.shape[0] * scale_percent / 100)
        dim = (width, height)
        # resize image
        resized = cv.resize(img, dim, interpolation = cv.INTER_AREA)

        cv.imwrite("screenshot-temp2.png",resized) 

    def read_ocr3(self):
        img = cv.imread('screenshot.png',0)
        # global thresholding
        ret1,th1 = cv.threshold(img,127,255,cv.THRESH_BINARY)
        # Otsu's thresholding
        ret2,th2 = cv.threshold(img,0,255,cv.THRESH_BINARY+cv.THRESH_OTSU)
        # Otsu's thresholding after Gaussian filtering
        blur = cv.GaussianBlur(img,(5,5),0)
        ret3,th3 = cv.threshold(blur,0,255,cv.THRESH_BINARY+cv.THRESH_OTSU)
        # plot all the images and their histograms
        images = [img, 0, th1,
                  img, 0, th2,
                  blur, 0, th3]
        titles = ['Original Noisy Image','Histogram','Global Thresholding (v=127)',
                  'Original Noisy Image','Histogram',"Otsu's Thresholding",
                  'Gaussian filtered Image','Histogram',"Otsu's Thresholding"]
        for i in range(3):
            plt.subplot(3,3,i*3+1),plt.imshow(images[i*3],'gray')
            plt.title(titles[i*3]), plt.xticks([]), plt.yticks([])
            plt.subplot(3,3,i*3+2),plt.hist(images[i*3].ravel(),256)
            plt.title(titles[i*3+1]), plt.xticks([]), plt.yticks([])
            plt.subplot(3,3,i*3+3),plt.imshow(images[i*3+2],'gray')
            plt.title(titles[i*3+2]), plt.xticks([]), plt.yticks([])
        plt.show()
    def read_ocr2(self,file):
        # transforma ela de colorida (RGB) para tons de cinza
        self.file=file
        img = cv.imread(self.file,0)

        th3 = cv.adaptiveThreshold(img,255,cv.ADAPTIVE_THRESH_GAUSSIAN_C,\
            cv.THRESH_BINARY,11,2)
        cv.imwrite("screenshot-temp.png",th3)
    
    def test_ocr(self):
        im = Image.open('screenshot.png')
        #im = im.filter(ImageFilter.MedianFilter()) # blur the image, the stripes will be erased
        im = ImageEnhance.Contrast(im).enhance(2)  # increase the contrast (to make image clear?)
        im = im.convert('1')                       # convert to black-white image
        im.show()
    def read_ocr(self):

        pytesseract.tesseract_cmd = r"C:\\Users\\Y4E9\\AppData\\Local\\Tesseract-OCR5\\tesseract.exe"
        text = pytesseract.image_to_string(Image.open('screenshot-temp.png'))       
        return text    
    def solveCaptcha2(self,file):
        
        self.file=file
        imagem = Image.open(self.file).convert('RGB')

        # convertendo em um array editável de numpy[x, y, CANALS]
        npimagem = np.asarray(imagem).astype(np.uint8)  

        # diminuição dos ruidos antes da binarização
        npimagem[:, :, 0] = 0 # zerando o canal R (RED)
        npimagem[:, :, 2] = 0 # zerando o canal B (BLUE)

        # atribuição em escala de cinza
        im = cv.cvtColor(npimagem, cv.COLOR_RGB2GRAY) 

        # aplicação da truncagem binária para a intensidade
        # pixels de intensidade de cor abaixo de 127 serão convertidos para 0 (PRETO)
        # pixels de intensidade de cor acima de 127 serão convertidos para 255 (BRANCO)
        # A atrubição do THRESH_OTSU incrementa uma análise inteligente dos nivels de truncagem
        ret, thresh = cv.threshold(im, 127, 255, cv.THRESH_BINARY | cv.THRESH_OTSU) 

        # reconvertendo o retorno do threshold em um objeto do tipo PIL.Image
        
        binimagem = Image.fromarray(thresh) 
        #binimagem.show()

        pytesseract.tesseract_cmd = r"C:\\Users\\Y4E9\\AppData\\Local\\Tesseract-OCR5\\tesseract.exe"
        text = pytesseract.image_to_string(binimagem)
        return text
    def solveCaptcha(self,file):
        
        self.file=file
        pytesseract.tesseract_cmd = r"C:\\Users\\Y4E9\\AppData\\Local\\Tesseract-OCR5\\tesseract.exe"

        text = pytesseract.image_to_string(Image.open(self.file))

        return text
                                 
class Excel:

    def __init__(self,file):
        self.file = file    
        try:
            wb = openpyxl.load_workbook(self.file)
            wb.get_sheet_names()
            self.planilha = wb
            #return self.planilha
        except:
            print('Não foi possivel abrir a planilha %s ', self.file)
            sys.exit()
        
    def salvarPlanilha(self,arquivo):
        self.arquivo = arquivo
        self.planilha.save(arquivo)
    def max_row(self):
        return self.max_row
    def lerAba(self,aba):
        self.aba = aba        
        try:
            abaPlanilha = self.planilha.get_sheet_by_name(self.aba)
            return abaPlanilha
        except:
            print('Não foi possivel ler na planilha %s a aba %s',self.planilha,self.aba)
            sys.exit()        
class ANPWebScrap:

    def __init__(self,url):
        self.url=url        
        browser = webdriver.Firefox()
        self.browser = browser
        self.browser.get(self.url)
        #self.browser.maximize_window()        
    def voltar(self):
        self.browser.back()
    def exportar(self):
        try:
            processarObj = WebDriverWait(self.browser, 5).until(
                EC.presence_of_element_located((By.NAME,'btnSalvar')))
            processarObj.click()    
        except:
            pass  
    def clickLink(self,link):
        self.link=link
        linkObj = WebDriverWait(self.browser, 5).until(
            EC.presence_of_element_located((By.PARTIAL_LINK_TEXT,self.link)))
        linkObj.click()
    def selecionarEstado(self,estado):
        self.estado = estado
        try:
            select = Select(self.browser.find_element_by_name('selEstado'))
            select.select_by_visible_text(self.estado)  
        except:
            pass      
    def inserirMunicipio(self,municipio):
        self.municipio = municipio
        municipioObj = WebDriverWait(self.browser, 5).until(
            EC.presence_of_element_located((By.NAME, 'txtMunicipio')))
        municipioObj.send_keys(self.municipio)
    def processar(self):
        try:
            processarObj = WebDriverWait(self.browser, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[text()[contains(.,"Processar")]]')))
            processarObj.click()    
        except:
            pass    
    def selecionarCombustivel(self,combustivel):
        self.combustivel = combustivel
        select = Select(self.browser.find_element_by_id('selCombustivel'))
        select.select_by_visible_text(self.combustivel)
    def escreveCaptcha(self,captcha):
        self.captcha=captcha
        #escrever no captcha o texto convertido
        try:
            captcha = WebDriverWait(self.browser, 5).until(
                EC.presence_of_element_located((By.NAME, 'txtValor')))
            captcha.send_keys(self.captcha)
            captcha.submit()    
        except:
            print('Não foi possivel escrever e/ou submeter o captcha')        
    def testaSucessoCaptcha(self):
        try:
            select = Select(self.browser.find_element_by_id('selCombustivel'))
            return False
        except:
            return True    
    def getColunas3(self,linha):
        self.linha=linha
        coluna = self.linha.find_all('tr')
        self.coluna = coluna
        return coluna  
    def getColunas2(self,linha):
        self.linha=linha
        coluna = self.linha.find_all('th')
        self.coluna = coluna
        return coluna  
    def getLinhas(self):
        rows = self.table.find_all('tr')
        #self.linha = rows
        return rows
    def getColunas(self,linha):
        self.linha=linha
        coluna = self.linha.find_all('td')
        self.coluna = coluna
        return coluna
    def lerTabela(self):
        html = self.browser.page_source
        soup = BeautifulSoup(html, 'html5lib')
        self.table = soup.find('tbody')

    def getCaptcha(self,file):
        self.file = file        
        #Imagem a ser quebrada, neste ponto você poderia usar urlib, httplib ou curl para carregar esta imagem.
        #self.browser.save_screenshot()
        png = self.browser.get_screenshot_as_png()
        element = WebDriverWait(self.browser, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="frmAberto"]/table/tbody/tr[2]/td[1]/img'))) 
        location = element.location
        size = element.size
        im = Image.open(BytesIO(png)) # uses PIL library to open image in memory

        left = location['x']
        top = location['y']
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']

        im = im.crop((left, top, right, bottom)) # defines crop points
        im.save(self.file) # saves new cropped image    

        
        
    