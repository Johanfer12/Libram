import pyautogui, time, pyperclip, openpyxl, webbrowser

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
contar = 0
location = 0

while contar < 1 and location != None:
            
       
    location = pyautogui.locateOnScreen(entrada, grayscale=True, confidence=.95)
    print ("ubicado crear entrada")
    print (location)
    pyautogui.moveTo(location)
    #pyautogui.click(duration=1)
    contar += 1
    print (contar)
    time.sleep(2.5)
