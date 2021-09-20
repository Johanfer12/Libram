import pyautogui, time, pyperclip, openpyxl, webbrowser
from playsound import playsound


path = ".\\Libros.xlsx"
wb = openpyxl.load_workbook(path)
sheet = wb['Librero']
max_row = sheet.max_row
contar = 0
erro = 0
location = 0
numblog = sheet['N2'].value

entrada = "D:\\crearentrada.png"
imagen = "D:\\imagen.png"
crearent = "D:\\titulo.png"
elegirarchivo = "D:\\elegirarchivo.png"
subirimg = "D:\\subirimg.png"
redactar = "D:\\redactar.png"
amper = "D:\\amper.png"
agreimg = "D:\\agreimg.png"
etiqueta = "D:\\etiqueta.png"
finalizado = "D:\\finalizado.png"
publicar = "D:\\publicar.png"
htm = "D:\\htm.png"


webbrowser.open('https://www.blogger.com/blogger.g?blogID="yourblogidnumberhere"#allposts/src=sidebar')
time.sleep(10)

while contar < 25 and location != None and erro == 0:
    
    if sheet.cell(row=numblog, column=6).value == "Si img" and sheet.cell(row=numblog, column=9).value != "No descripción":
        print ("Disponible libro número " + str(numblog))

# Crear Entrada    
          
        location = pyautogui.locateOnScreen(entrada, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado crear entrada")
            pyautogui.moveTo(location)
            pyautogui.click(duration=1)
            time.sleep(2.5)
            erro = 0
        else:
            print ("Crear entrada no encontrada")
            erro = 1
                                
#htm
        
        location = pyautogui.locateOnScreen(htm, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado Htm")
            pyautogui.moveTo(location)
            pyautogui.click(duration=1)
            time.sleep(1)
            pyperclip.copy(str(sheet.cell(row=1, column=15).value)+str(sheet.cell(row=numblog, column=1).value)+str(sheet.cell(row=3, column=15).value)+str(sheet.cell(row=4, column=15).value)+str(sheet.cell(row=5, column=15).value)+str(sheet.cell(row=6, column=15).value)+str(sheet.cell(row=7, column=15).value)+str(sheet.cell(row=numblog, column=9).value)+str(sheet.cell(row=9, column=15).value)+str(sheet.cell(row=10, column=15).value)+str(sheet.cell(row=11, column=15).value)+str(sheet.cell(row=12, column=15).value)+str(sheet.cell(row=13, column=15).value)+str(sheet.cell(row=14, column=15).value)+str(sheet.cell(row=15, column=15).value)+str(sheet.cell(row=16, column=15).value)+str(sheet.cell(row=17, column=15).value)+str(sheet.cell(row=numblog, column=8).value)+str(sheet.cell(row=19, column=15).value))
            pyautogui.hotkey('ctrl', 'v', interval = 0.15)
            erro = 0
        else:
            print ("Htm no encontrado")
            erro = 1
            
#Titulo
        
        location = pyautogui.locateOnScreen(crearent, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado Titulo")
            pyautogui.moveTo(location)
            pyautogui.click(duration=1)
            pyperclip.copy(str(sheet.cell(row=numblog, column=1).value) + " by " + str(sheet.cell(row=numblog, column=2).value))
            pyautogui.hotkey('ctrl', 'v', interval = 0.15)
            erro = 0
        else:
            print ("Titulo no encontrado")
            erro = 1
#Redactar
        
        location = pyautogui.locateOnScreen(redactar, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado Redactar")
            pyautogui.moveTo(location)
            pyautogui.click(duration=1)
            time.sleep(1)
            erro = 0
        else:
            print ("Redactar no encontrado")
            erro = 1

#posicion imagen
            
        location = pyautogui.locateOnScreen(amper, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado posicion imagen")
            pyautogui.moveTo(location)
            pyautogui.click(clicks=3)
            pyautogui.press('delete', interval=0.1)
            pyautogui.press(['up', 'up'])
            time.sleep(1)
            erro = 0
        else:
            print ("Posicion imagen no encontrado")
            erro = 1

#Subir imagen
        
        location = pyautogui.locateOnScreen(imagen, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado img")
            pyautogui.moveTo(location)
            pyautogui.click(duration=1)
            time.sleep(3)
            erro = 0
        else:
            print ("Img no encontrado")
            erro = 1
        
#elegir imagen
        
        location = pyautogui.locateOnScreen(elegirarchivo, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado elegir img")
            pyautogui.moveTo(location)
            pyautogui.click(duration=1)
            time.sleep(2)
            erro = 0
        else:
            print ("Elegir img no encontrado")
            erro = 1

#Subir imagen
        
        location = pyautogui.locateOnScreen(subirimg, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado Subir imagen")
            pyautogui.moveTo(location)
            pyautogui.click(duration=1)
            pyperclip.copy(str(sheet.cell(row=numblog, column=7).value))
            pyautogui.hotkey('ctrl', 'v', interval = 0.15)
            time.sleep(1)
            pyautogui.press('enter', interval=0.1)
            time.sleep(4)
            erro = 0
        else:
            print ("Subir imagen no encontrado")
            erro = 1

#agre imagen
        
        location = pyautogui.locateOnScreen(agreimg, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado agreg. img")
            pyautogui.moveTo(location)
            time.sleep(2)
            pyautogui.click(duration=1)
            time.sleep(2)
            erro = 0
        else:
            print ("Agreg. img no encontrado")
            erro = 1

#etiqueta
       
        location = pyautogui.locateOnScreen(etiqueta, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado etiqueta")
            pyautogui.moveTo(location)
            pyautogui.click(duration=1)
            time.sleep(2)
            pyperclip.copy(str(sheet.cell(row=numblog, column=10).value) + ", " + str(sheet.cell(row=numblog, column=2).value))
            pyautogui.hotkey('ctrl', 'v', interval = 0.15)
            erro = 0
        else:
            print ("Etiqueta no encontrado")
            erro = 1
            
        location = pyautogui.locateOnScreen(finalizado, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado Finalizado")
            pyautogui.moveTo(location)            
            pyautogui.click(duration=1)
            time.sleep(3)
            erro = 0
        else:
            print ("Elegir img. no encontrado")
            erro = 1

#Publicar
    
        location = pyautogui.locateOnScreen(publicar, grayscale=True, confidence=.9)
        
        if location != None:
            print ("Ubicado Publicar \n")
            pyautogui.moveTo(location)
            time.sleep(1)
            pyautogui.click(duration=1)
            time.sleep(3)
            contar += 1
            erro = 0
        else:
            print ("Publicar no encontrado")
            erro = 1
            
    else:
        print ("No disponible libro número " + str(numblog) + "\n")
        #contar += 1

    if erro == 0:
        numblog += 1
        wb.save(path)
    else:
        print('¡Falló algo!')
  
sheet['N2'] = numblog
wb.save(path)
print('¡Finalizado!, quedamos en: ' + str(numblog))
playsound('cat.wav')


