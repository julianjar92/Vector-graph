##HACER QUE EL PROGRAMA IMPROMA LOS GRAFICOS PARA CADA TIPO DE MODELO, Y GENERAR TABLAS CON SUS VALORES DE AZIMUTH Y MAGNITUD


import os                                                                                                               ## Importar modulo de comandos del sistema windows CMD
import math                                                                                                             ## Importar modulo matematico de python
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import load_workbook                                                                                      ##  Importa modulo de lectura de archivos (load_workbook ) -->>
                                                                                                                        ##  -->>de la libreria openpyxl
#Funciones Matematicas#
#//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////#
def magnitude(x1,y1,x2,y2):                                                                                             # X = Coordebadas de longitud Y = Coordenadas de latitud 
    raiz = round(math.sqrt((x2-x1)**2+(y2-y1)**2),2)                                                                    #calcula la magnitud de un vector y luego la redondea a dos decimales
    print ("Magnitud: ", str(raiz))    
    return raiz                                                                                                         ##devuelve el valor resultante de la operancion
    
def azimuth(y, x):

    rads = math.atan2(x, y)                                                                                             #Calcula la magnitud de un vector, se invierten xy, ya que los angulos->       
    angulo = round(math.degrees(rads),2)                                                                                #-> Se grafican en la cartografia, se indican en el sentido horario 
    if angulo < 0:
        angulo = angulo + 360
        print("Azimuth: " + str(angulo))       
        return angulo                         
    else: 
        print("Azimuth: " + str(angulo))       
        return angulo                                                                                                   ##devuelve el valor resultante de la operancion

#recorrido angular de los circulos de nivel
an = np.linspace(0, 2 * np.pi, 100)
#//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////#
     
###VARIABLES DE INDENTIFICACION###
estacion = "ABCC"
latd = 16.9
long = -1.07

#Colores en codigo hexadecimal y entre comillas 
gris = '#BBBBBB'
verde = '#008000'                                                                                   
negro = '#000000'    
rojo = '#C70039' 
amarillo = '#FFD93F'
colores = [verde,negro,rojo,amarillo] 
  
#Texto de los puntos cardinales
direccion = ['N','S','E','O']
                                                                                             
##Valores indicativos de las celdas de los modelos de movmiento de placas tectonicas

#Modelos de tipo geodesico
##LISTA           Long    Lat  Evel Nvel   Modelo de movimiento Color
ITRF2008       = ['B12','C12','D11','E11',      "ITRF2008",  '#8F0724'] #Modelo Geodesico       
ITRF2000AS     = ['B24','C24','D23','E23',    "ITRF2000AS",  '#FF0000'] #Modelo Geodesico
ITRF2000DA     = ['B30','C30','D29','E29',    "ITRF2000DA",  '#FF8000'] #Modelo Geodesico
APKIM2005_DGFI = ['B14','C14','D13','E13',"APKIM2005_DGFI",  '#86D200'] #Modelo Geodesico
APKIM2005_IGN  = ['B16','C16','D15','E15', "APKIM2005_IGN",  '#29D200'] #Modelo Geodesico
APKIM2000      = ['B28','C28','D27','E27',     "APKIM2000",  '#008C2B'] #Modelo Geodesico
CGPS2004       = ['B20','C20','D19','E19',      "CGPS2004",  '#00FFD4'] #Modelo Geodesico
REVEL2000      = ['B22','C22','D21','E21',     "REVEL2000",  '#00A7E5'] #Modelo Geodesico
GEODVEL2010    = ['B8','C8','D7','E7',       "GEODVEL2010",  '#0061E5'] #Modelo Geodesico

#Modelo Matricial para seleccion de modelo de movmiento de placa  y seleccion de sus celdas correspodientes en excel
MotionModel_Geodesic = [ITRF2008,ITRF2008,ITRF2000AS,ITRF2000DA,APKIM2005_DGFI,APKIM2005_IGN,APKIM2000,CGPS2004,REVEL2000,GEODVEL2010]

#Modelos de tipo Geofisico
##LISTA           Long  Lat  Evel Nvel  Modelo de movimiento Color
NNR_MORVEL     = ['B6','C6','D5','E5',        "NNR_MORVEL",   "#1AD3C2"] #Modelo Geofisico
HS3_NUVEL1A    = ['B26','C26','D25','E25',   "HS3_NUVEL1A",   "#40008D"] #Modelo Geofisico
HS2_NUVEL1A    = ['B32','C32','D31','E31',   "HS2_NUVEL1A",   "#20008D"] #Modelo Geofisico
NUVEL1A        = ['B34','C34','D33','E33',       "NUVEL1A",   "#C50C66"] #Modelo Geofisico
NUVEL1         = ['B36','C36','D35','E35',        "NUVEL1",   "#FC1EFF"] #Modelo Geofisico

#Modelo Matricial para seleccion de modelo de movmiento de placa  y seleccion de sus celdas correspodientes en excel
MotionModel_Geodephysic = [NNR_MORVEL,HS3_NUVEL1A,HS2_NUVEL1A, NUVEL1A, NUVEL1]

#Modelos de movimiento de tipo combinado
##LISTA            Long  Lat Evel Nvel  Modelo de movimiento Color
GSMR2_1        = ['B4','C4','D3','E3',           "GSMR2_1", "#565656"] #Modelo Combinado
GSMR1_2        = ['B18','C18','D17','E17',       "GSMR1_2", "#999999"] #Modelo Combinado
MORVEL2010     = ['B10','C10','D9','E9',      "MORVEL2010", "#7EBBA0"] #Modelo Combinado

#Modelo Matricial para seleccion de modelo de movmiento de placa  y seleccion de sus celdas correspodientes en excel
MotionModel_Combinated =[GSMR2_1,GSMR1_2,MORVEL2010]


###UBUCACION DE DIRECTORIO DE TRABAJO DEL PROGRAMA
#//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////#
os.chdir("D:\GNSS Project Files\MODEL MOTION PLATE EXCEL\ESTACIONES SA(NNR)\MAGNAECO")                                  ##Asignacion de la ruta donde se encuentran los archivos xlsx
ruta = os.getcwd()                                                                                                      #Obtiene de ubicacion de ruta actual en el programa
print ("DIRECCION: " + ruta)                                                                                            ##imprime la direccion actual del programa mediante el metodo getcwd()
#//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////#

wb = load_workbook(estacion+'_SA(NNR).xlsx')                                                                            ##Comando para cargar  el archivo.xlsx
sheetname = str(wb.get_sheet_names())                                                                                   ##Comando para obtener el nombre de las hojas de calculo del archivo y convertirlo en string
sheetname = sheetname[2:-2]                                                                                             ##Se ajusta el nombre del sheetname ya que viene con estos caracteres de mas ['sheetname']
print ('El archuvo ' + sheetname +'.xlsx ha sido abierto')  
sheet_ranges = wb[sheetname]                                                                                            ##De esta forma se guarda la hoja de calculo como un objeto y sus celdas como atributos del objeto
print(sheet_ranges)

##Configuracion de la ventana y tipo de grafico a usar
plt.figure(figsize=(10, 10), dpi=80)                                                                                    #Configura el tamaño y la resolucion de la ventana del grafico
plt.xlim(-21.3, 21.3)                                                                                                   #Estabkece los limites del eje X
plt.ylim(-21.3, 21.3)                                                                                                   #Estabkece los limites del eje Y


#Configuracion del ubcacion de los ejes XY                                                                                                                         
ax = plt.axes()                                                                                                         #Crea el objeto ax de la clase axes con las caracteristicas del eje
ax.spines['right'].set_color('none')                                                                                    #Le quita el color a la linea del borde derecho del marco del eje
ax.spines['top'].set_color('none')                                                                                      #Le quita el color a la linea del borde superior del marco del eje
ax.xaxis.set_ticks_position('bottom')                                                                                   #Selecciona el eje inferior x para que su posiciohn sea ajustada
ax.spines['bottom'].set_position(('data',0))                                                                            # Se Ajusta la interseccion del eje x al punto 0
ax.yaxis.set_ticks_position('left')                                                                                     #Selecciona el eje inferior y para que su posicion sea ajustada                                                                                 #Selecciona el eje inferior x para ser ajustado
ax.spines['left'].set_position(('data',0))                                                                              # Se Ajusta la interseccion del eje y al punto 0
    
                                                                 
magnitud = magnitude(0, 0, long, latd)
grados = azimuth(latd, long)

##Graficar circulos de nivel
ax.plot(21.3 * np.cos(an), 21.3 * np.sin(an), color=gris, linestyle="-.")
for i in range(0,21,5): ax.plot(i * np.cos(an), i * np.sin(an), color=gris, linestyle="-.")

#Graficacion del vector en Coordenadas polares                                                                          #Se le asignan los atributos de coordenadas y stilo al objeto flecha.
plt.arrow(0, 0, long, latd, head_width=0.4, head_length=0.8, width = 0.15,  fc=negro, ec=negro)                         #(X, ,Y, unidades a recorrer en x,unidades a recorrer en Y)  
                                                                                                                        #fc = color relleno de la flecha ec = color del borde de la flecha.   
#LEYENDA PPP
plt.arrow(-18.5, 21.1, 1, 0, head_width=0.30, head_length=0.60, width = 0.10,  fc=negro, ec=negro)                     #flecha de leyenda 
plt.text(-22,21, "PPP", family='serif', style='italic', ha='center', wrap=True, size=10)                                 #Texto de leyenda
                            
counter = 0
for model in MotionModel_Combinated:
    print(model[4])
    Evel = sheet_ranges[model[2]].value                                                                                 # Obtencion de los datos  de Evel de las celdas dentro de excel
    Nvel = sheet_ranges[model[3]].value                                                                                 # Obtencion de los datos  de Nvel de las celdas dentro de excel
    
    #Calculo de magnitud y azimuth 
    print(Evel, Nvel)
    magnitud = magnitude(0, 0, Evel, Nvel)
    grados = azimuth(Nvel, Evel)

    #Graficacion del vector 
    plt.arrow(0, 0, Evel, Nvel, head_width=0.40, head_length=0.80, width = 0.15,  fc=model[5], ec=negro)    

    #Graficacion de la leyenda del vector
    plt.arrow(-18.5, 20.25-counter, 1, 0, head_width=0.30, head_length=0.60, width = 0.10,  fc=model[5], ec=negro)                  #flecha de leyenda 
    plt.text(-22.3,20-counter, model[4], family='serif', style='italic', ha='center', wrap=True, size=10)                              #Texto de leyenda
                                                                                       
    counter = counter + 1

#Escritura de los Puntos cardinales 
plt.text(1,20.3, direccion[0], family='serif', style='italic', ha='right', wrap=True, size=15) 
plt.text(1,-21, direccion[1], family='serif', style='italic', ha='right', wrap=True, size=15) 
plt.text(21, 0.3, direccion[2], family='serif', style='italic', ha='right', wrap=True, size=15) 
plt.text(-20.3, 0.3, direccion[3], family='serif', style='italic', ha='right', wrap=True, size=15) 

ax.set_title("Vector de velocidad Estacion " + estacion, va='bottom',family='serif', style='italic', size=18)           # Titulo del grafico

plt.grid(False)                                                                                                         #Se omite la graficacion de la grilla
plt.show()                                                                                                              #Se crea muestra el grafico


#savefig("Vectores_Estacion:" + estacion + ".png", dpi=80)                                                              # Guardar la figura usando 72 puntos por pulgada