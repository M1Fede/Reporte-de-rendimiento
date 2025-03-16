# -----------------------------------------------------------------------------
#                           Â¿Que hace este codigo? 
# Genera el reporte trimestral que se envia al cliente. El codigo toma imagenes y 
# texto, y genera tablas y graficas que son colocados en un PDF (reporte)  

# CONSIDERACIONES ANTES DEL USO DEL SCRIPT
#   1) El parametro 'trimestre' sirve para la creacion/identificacion de la 
#      carpeta donde se guardaran los reportes
#   2) El año del parametro 'fecha_cierre' sirve, entre otras cosas, para la 
#      creacion/identificacion de la carpeta donde se guardaran los reportes

# CONCLUSION: Prestar atencion a la pre existencia de la carpeta

# -----------------------------------------------------------------------------

# Parametros
alyc = 'Bull' # Tres alycs Bull, IEB, y Balanz

usuario = 4

numero_interno = 61

fecha_cierre = '2024-12-31'

fecha_inicial = '2022-12-31'

trimestre = 4

# ----------------------
# -------------------------------------------------------
# -----------------------------------------------------------------------------
if usuario == 1: 
    sub_directorio = 'YYY'
    auxiliar = '--'
elif usuario == 2:
    sub_directorio = 'YYY_YYY'
    auxiliar = '--'
elif usuario == 3:
    sub_directorio = 'YYY_YYY_YYY'
    auxiliar = ''
elif usuario == 4:
    sub_directorio = 'YY'
    auxiliar = ''
elif usuario == 5:
    sub_directorio = 'YY_YY'
    auxiliar = ''
elif usuario == 6:
    sub_directorio = 'YY_YY_YY'
    auxiliar = ''

directorio_funciones=f'C:/Users\{sub_directorio}\Dropbox{auxiliar}\carpeta con cosas\carpeta con mas cosas\libreria_py_c' 

import sys
sys.path.append(f'{directorio_funciones}')
import pandas as pd
import numpy as np
import dp_funciones_c as fc
from unidecode import unidecode
from datetime import datetime as dt
from datetime import timedelta
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import PageBreak
from io import BytesIO
from reportlab.lib.pagesizes import A4
from docx import Document




# -----------------------------------------------------------------------------
# Corregimos la redaccion del nombre de la alyc
alyc = alyc.title()
alyc = unidecode(alyc)

# Ahora buscamos el excel con los rendimientos del cliente
datos_cliente = fc.cliente(alyc = alyc, 
                            numero_interno = numero_interno, 
                            usuario = usuario)

try:
    nombre_cliente = datos_cliente.loc['nombre cliente','Datos del cliente']
    numero_cliente = datos_cliente.loc['numero cliente','Datos del cliente']  
    fecha_movimientos = datos_cliente.loc['fecha movimientos','Datos del cliente']
    dia_corte = datos_cliente.loc['Dia de corte','Datos del cliente']

except:
    nombre_cliente = ''
    numero_cliente = 0
    fecha_movimientos = ''


# -----------------------------------------------------------------------------
# Ubicacion del archivo excel con la serie del dolar mep y tasa plazo fijo
nombre_excel1 = 'serie dolar y plazo fijo (para codigo python)'
direccion_excel1 = f'C:/Users\{sub_directorio}\Dropbox{auxiliar}\Informes Clientes\{nombre_excel1}'

excel_usd = pd.read_excel(f'{direccion_excel1}.xlsx', sheet_name = 'MEP')
excel_usd.set_index('Fecha',inplace = True)

excel_pf = pd.read_excel(f'{direccion_excel1}.xlsx', sheet_name = 'Plazo fijo')
excel_pf.set_index('Fecha',inplace = True)

serie_rend = 'Series'

# Ubicacion del archivo excel con la base de datos de clientes
nombre_bclientes = 'Base de Datos de Clientes'
base_clientes = f'C:/Users\{sub_directorio}\Dropbox{auxiliar}\caperta con cosas\{nombre_bclientes}' 

excel_bc =  pd.read_excel(f'{base_clientes}.xlsx', sheet_name = 'Clientes')
excel_bc.set_index('comitente',inplace = True)


# -----------------------------------------------------------------------------
# Definimos ubicacion de los informes y el nombre de los nuevos archivos
anio = f'20{fecha_cierre[2:4]}'
direccion_informes = f'C:/Users\{sub_directorio}\Dropbox{auxiliar}\Informes Clientes'

if alyc == 'Bull':
    cuentas = 'Cuentas de Bull'
    
elif alyc == 'IEB':
    cuentas = 'Cuentas de IEB'
    
elif alyc == 'Balanz':
    cuentas = 'Cuentas de Balanz'

ubicacion_archivo = f'{direccion_informes}/{cuentas}/{nombre_cliente} ({numero_cliente})'

destino_pdf = f'C:/Users\{sub_directorio}\Dropbox{auxiliar}\Informes Clientes\Reportes /{anio} {trimestre}T'



# --------------------------- PARTE: PRELUDIO ---------------------------------
# -----------------------------------------------------------------------------
# ------------------ BLOQUE 1 de 2: Definiendo el periodo ---------------------
# A partir de la fecha incial y la fecha de cierre se crea una serie de momentos
# que identifican al periodo de interes (trimestre, año, etc)
fecha_cierre = dt.strptime(fecha_cierre, "%Y-%m-%d")
fecha_inicial = dt.strptime(fecha_inicial , "%Y-%m-%d")


# Identificamos si el periodo esta comprendido en un mismo anio o en varios y
# creamos la lista 'fechas' para que contenga la serie de momentos del periodo 
# de interes
fechas = []
anios = fecha_cierre.year - fecha_inicial.year


if anios == 0: # Periodo dentro del mismo anio
    
    # Identificamos los meses comprendidos en el periodo de interes
    for i in range(fecha_cierre.month - fecha_inicial.month + 1):
        mes = fecha_cierre.month - i
        
        # Definimos el dia que corresponde a cada mes.
        if (mes == 1) or (mes == 3) or (mes == 5) or (mes == 7) or (mes == 8) or (mes == 10) or (mes == 12):
            dias = 31
        
        elif (mes == 4) or (mes == 6) or (mes == 9) or (mes == 11):
            dias = 30
            
        elif mes == 2:
            dias = 28
           
        else:
            dias = 'El mes no se especifico correctamente'
           
        # Definimos un string con formato de fecha y le damos formato datetime
        # para incorporarlo a la lista 'fechas'
        fecha = f'{fecha_cierre.year}-{mes}-{dias}'
        fechas.append(dt.strptime(fecha, "%Y-%m-%d"))
    

elif anios > 0: # Periodo dentro de diferentes anios
    
    # Armamos la lista por partes, primero para el anio de la fecha de cierre
    # y luego para el resto de los anios, finalmente se une todo
    for i in range(anios + 1):
        anioo = fecha_cierre.year - i
        
        if anioo == fecha_cierre.year:
            fecha_inicial_artificial = dt.strptime(f'{anioo}-01-31', "%Y-%m-%d")
            
            # Identificamos los meses comprendidos en el periodo de interes
            for j in range(fecha_cierre.month - fecha_inicial_artificial.month + 1):
                mes = fecha_cierre.month - j
                
                # Definimos el dia que corresponde a cada mes.
                if (mes == 1) or (mes == 3) or (mes == 5) or (mes == 7) or (mes == 8) or (mes == 10) or (mes == 12):
                    dias = 31
                
                elif (mes == 4) or (mes == 6) or (mes == 9) or (mes == 11):
                    dias = 30
                    
                elif mes == 2:
                    try:
                        dias = 29
                        fecha_prueba = dt.strptime(f'{anioo}-02-29', "%Y-%m-%d")
                        
                    except:
                        dias = 28
                   
                else:
                    dias = 'El mes no se especifico correctamente'
                   
                # Definimos un string con formato de fecha y le damos formato datetime
                # para incorporarlo a la lista 'fechas'
                fecha = f'{fecha_cierre.year}-{mes}-{dias}'
                fechas.append(dt.strptime(fecha, "%Y-%m-%d"))
                
        if anioo < fecha_cierre.year:
            fecha_cierre_artificial = dt.strptime(f'{anioo}-12-31', "%Y-%m-%d")
            
            # Creamos una fecha inicial artifical si es que el anio no coincide
            # con el anio de la fecha inicial
            if anioo > fecha_inicial.year:
                fecha_inicial_artificial = dt.strptime(f'{anioo}-01-31', "%Y-%m-%d")
            
                # Identificamos los meses comprendidos en el periodo de interes
                for h in range(fecha_cierre_artificial.month - fecha_inicial_artificial.month + 1):
                    mes = fecha_cierre_artificial.month - h
                    
                    # Definimos el dia que corresponde a cada mes.
                    if (mes == 1) or (mes == 3) or (mes == 5) or (mes == 7) or (mes == 8) or (mes == 10) or (mes == 12):
                        dias = 31
                    
                    elif (mes == 4) or (mes == 6) or (mes == 9) or (mes == 11):
                        dias = 30
                        
                    elif mes == 2:
                        try:
                            dias = 29
                            fecha_prueba = dt.strptime(f'{anioo}-02-29', "%Y-%m-%d")
                            
                        except:
                            dias = 28
                       
                    else:
                        dias = 'El mes no se especifico correctamente'
                       
                    # Definimos un string con formato de fecha y le damos formato datetime
                    # para incorporarlo a la lista 'fechas'
                    fecha = f'{fecha_cierre_artificial.year}-{mes}-{dias}'
                    fechas.append(dt.strptime(fecha, "%Y-%m-%d"))
            
            if anioo == fecha_inicial.year:
                # Identificamos los meses comprendidos en el periodo de interes
                for p in range(fecha_cierre_artificial.month - fecha_inicial.month + 1):
                    mes = fecha_cierre_artificial.month - p
                    
                    # Definimos el dia que corresponde a cada mes.
                    if (mes == 1) or (mes == 3) or (mes == 5) or (mes == 7) or (mes == 8) or (mes == 10) or (mes == 12):
                        dias = 31
                    
                    elif (mes == 4) or (mes == 6) or (mes == 9) or (mes == 11):
                        dias = 30
                        
                    elif mes == 2:
                        try:
                            dias = 29
                            fecha_prueba = dt.strptime(f'{anioo}-02-29', "%Y-%m-%d")
                            
                        except:
                            dias = 28
                       
                    else:
                        dias = 'El mes no se especifico correctamente'
                       
                    # Definimos un string con formato de fecha y le damos formato datetime
                    # para incorporarlo a la lista 'fechas'
                    fecha = f'{fecha_cierre_artificial.year}-{mes}-{dias}'
                    fechas.append(dt.strptime(fecha, "%Y-%m-%d"))
    
  
# Creamos al Dataframe
fechas = pd.DataFrame(fechas)
fechas = fechas.rename(columns={0: 'momento'})


# Agregamos las fechas de cierre si es que no esta. La inicial no porque en la 
# lista tienen que aparecer los momentos finales de cada subperiodo
if len(fechas.loc[fechas.momento == fecha_cierre]) == 0:
    fechas.loc[len(fechas),'momento'] = fecha_cierre

# Ordenamos de mayor a menor
fechas.sort_values('momento', ascending = False, inplace = True)


# --------------- BLOQUE 2 de 2: Definiendo el periodo del cliente ------------
# Del Dataframe 'fechas' nos quedamos con la mascara que respeta las fechas
# de alta y de baja del cliente. 

# Identificamos el momento de alta y de baja, luego tomamos la mascara.
try:
    alta = excel_bc.loc[numero_cliente, 'Inicio de gestion']
    
except:
    alta = 'La fecha de inicio de gestion no existe en la base de datos de clientes'

try:
    baja = excel_bc.loc[numero_cliente, 'Baja de cliente']

    if type(baja) != type(alta): # hay veces que la celda esta vacia y python la 
        baja = 'Cliente activo'  # lee como un string cuyo contenido es N/A 
    
except:
    baja = 'Cliente activo'


# Agregamos las fechas de alta y de baja y luego tomamos la mascara
if alta != 'La fecha de inicio de gestion no existe en la base de datos de clientes':
    
    # Agregamos la fecha de alta y la fecha de baja
    if len(fechas.loc[fechas.momento == alta]) == 0:
        fechas.loc[len(fechas),'momento'] = alta

    if baja != 'Cliente activo':
        fechas.loc[len(fechas),'momento'] = baja
    
    # Tomamos la mascara
    # Caso donde el cliente no baja su cuenta.
    if baja == 'Cliente activo':
        # mascara fecha_cierre
        fechas = fechas.loc[fechas.momento <= fecha_cierre].copy()
        
        # mascara fecha inicial o de alta
        if fecha_inicial >= alta:
            fechas = fechas.loc[fechas.momento >= fecha_inicial ].copy()     
        
        elif fecha_inicial < alta:
            fechas = fechas.loc[fechas.momento >= alta].copy()  
            
    # Caso donde el cliente baja su cuenta 
    else:
        # mascara fecha inicial o alta
        if fecha_inicial >= alta:
            fechas = fechas.loc[fechas.momento >= fecha_inicial ].copy()     
        
        elif fecha_inicial < alta:
            fechas = fechas.loc[fechas.momento >= alta].copy() 
        
        # mascara fecha de cierre o baja
        if fecha_cierre <= baja:
            fechas = fechas.loc[fechas.momento <= fecha_cierre].copy()
            
        elif fecha_cierre > baja: 
            fechas = fechas.loc[fechas.momento <= baja].copy()
   
else:
    print(alta) # Asi se indica el error que impide tomar la mascara.
    
    
# Reseteamos el indice para que se mantenga la secuencia de numeros consecutivos 
fechas.set_index('momento', inplace = True)
fechas.reset_index(inplace = True)




# ----------------------------  PRIMERA PARTE ---------------------------------
# Se identifican los depositos, retiros, y transferencias que ha hecho el 
# cliente durante el periodo de interes. Para la tabla de movimientos
# -----------------------------------------------------------------------------
# Se defininen las fechas de los momentos inicial y final del periodo de interes
if fecha_inicial > alta: 
    fecha_inicial_movimientos = fecha_inicial
    fecha_inicial_movimientos = fecha_inicial_movimientos.strftime("%Y-%m-%d")
    
elif fecha_inicial <= alta:
    fecha_inicial_movimientos = alta
    fecha_inicial_movimientos = fecha_inicial_movimientos.strftime("%Y-%m-%d")

if type(baja) != type('Cliente activo'):
    if fecha_cierre < baja: 
        fecha_cierre_movimientos = fecha_cierre
        fecha_cierre_movimientos = fecha_cierre_movimientos.strftime("%Y-%m-%d")
    
    elif fecha_cierre >= baja:
        fecha_cierre_movimientos = baja
        fecha_cierre_movimientos = fecha_cierre_movimientos.strftime("%Y-%m-%d")
    
elif baja == 'Cliente activo':
    fecha_cierre_movimientos = fecha_cierre
    fecha_cierre_movimientos = fecha_cierre_movimientos.strftime("%Y-%m-%d")
    

# Obtenemos los movimientos del periodo de interes
if alyc == 'Bull':
    movimientos = fc.depositos_retiros_bull(fecha_cierre = fecha_cierre_movimientos,
                                            fecha_inicial = fecha_inicial_movimientos,
                                            usuario = usuario,
                                            numero_interno = numero_interno,
                                            ctte_adm = '')

elif alyc == 'Balanz':
    movimientos = fc.depositos_retiros_balanz(fecha_cierre = fecha_cierre_movimientos,
                                              fecha_inicial = fecha_inicial_movimientos,
                                              usuario = usuario,
                                              numero_interno = numero_interno,
                                              ctte_adm = '')

elif alyc == 'Ieb':
    movimientos = fc.depositos_retiros_ieb(fecha_cierre = fecha_cierre_movimientos,
                                           fecha_inicial = fecha_inicial_movimientos,
                                           usuario = usuario,
                                           numero_interno = numero_interno,
                                           ctte_adm = '')

retiros_totales = movimientos.retiros.sum()
depositos_totales = movimientos.depositos.sum()


# Aplicamos mayusculas a los nombres de las columnas
movimientos.columns = movimientos.columns.str.capitalize()

 
# Transformamos las fechas de cierre e inicial en string para usos posteriores.
fecha_cierre = fecha_cierre.strftime("%Y-%m-%d")
fecha_inicial = fecha_inicial.strftime("%Y-%m-%d")




# # ----------------------------  SEGUNDA PARTE ---------------------------------
# Se calculan los rendimientos y el valor inicial, final, y la ganancia absoluta 
# del periodo
# -----------------------------------------------------------------------------
# Creamos la lista que contiene los valores de las carteras y dos floats que 
# contienen el valor inicial y final de la cartera
valores_cartera = []
valor_inicial = float(0)
valor_final = float(0)


# Calculamos el valor inicial de la cartera.
# (Con los honorarios calculamos el activo, por lo que debemos sumar la)
#    (palanca, pero con el rendimiento calculamos el patrimonio neto).
if alyc == 'Bull':
    cartera_inicio = fc.composicion_cartera_bull(fecha_cierre = fecha_inicial_movimientos,
                                                  alyc = alyc, 
                                                  numero_interno = numero_interno,
                                                  usuario = usuario)

    cartera_inicio['monto'] = cartera_inicio.Cantidad * cartera_inicio.iloc[:,1]

    cartera_inicio = cartera_inicio.loc[cartera_inicio.index != 'MEP'].copy()
   
elif alyc == 'Ieb':
    cartera_inicio = fc.composicion_cartera_ieb(fecha_cierre = fecha_inicial_movimientos,
                                                alyc = alyc, 
                                                numero_interno = numero_interno,
                                                usuario = usuario)

    cartera_inicio['monto'] = cartera_inicio.Cantidad * cartera_inicio.iloc[:,1]

    cartera_inicio = cartera_inicio.loc[cartera_inicio.index != 'MEP'].copy()
    
elif alyc == 'Balanz':
    cartera_inicio = fc.composicion_cartera_bal(fecha_cierre = fecha_inicial_movimientos,
                                                alyc = alyc, 
                                                numero_interno = numero_interno,
                                                usuario = usuario)

    cartera_inicio['monto'] = cartera_inicio.Cantidad * cartera_inicio.iloc[:,1]

    cartera_inicio = cartera_inicio.loc[cartera_inicio.index != 'MEP'].copy()

valor_inicial = cartera_inicio.monto.sum()


# Creamos un vector fecha y el dataframe de rendimientos
vector_fecha = []

for i in fechas.momento:
    vector_fecha.append(i.strftime("%Y-%m-%d"))

tabla_rendimientos = pd.DataFrame()
tabla_rendimientos['Fechas'] = str()
tabla_rendimientos['Bruto'] = float()
tabla_rendimientos['Neto'] =float()
tabla_rendimientos['Honorarios*'] =float()

for i in range(len(vector_fecha)):
    tabla_rendimientos.loc[i,'Fechas'] = vector_fecha[i]
    tabla_rendimientos.loc[i,'Bruto'] = float(0)
    tabla_rendimientos.loc[i,'Neto'] = float(0)
    tabla_rendimientos.loc[i,'Honorarios*'] = float(0)

# Hay que ordenar la tabla de rendimientos de acuerdo a la columna fechas. Para
# esto convertimos las fechas de str a datetime y luego de ordenar la regresamos a str.  
for i in range(len(tabla_rendimientos)):
    tabla_rendimientos.iloc[i,0] = dt.strptime(tabla_rendimientos.iloc[i,0], "%Y-%m-%d")

tabla_rendimientos.sort_values(by = 'Fechas', ascending = True, inplace = True)

for i in range(len(tabla_rendimientos)):
    tabla_rendimientos.iloc[i,0] = tabla_rendimientos.iloc[i,0].strftime("%Y-%m-%d")

tabla_rendimientos.set_index('Fechas', inplace = True)
tabla_rendimientos.reset_index(inplace = True)

# Reordenamos los elementos de la lista 'vector_fecha'
for i in range(len(vector_fecha)):
    vector_fecha[i] = tabla_rendimientos.iloc[i,0]

# Eliminamos la primera fila de la 'tabla de rendimientos' 
tabla_rendimientos = tabla_rendimientos.iloc[1:].copy()

# Reordenando la lista 'vector_fecha'
vector_fecha_bis = vector_fecha.copy()

for i in range(len(vector_fecha)):
    posicion = len(vector_fecha) - i - 1
    vector_fecha[i] = vector_fecha_bis[posicion]


# Rellenamos el dataframe con los rendimientos y fechas, y calculamos los valores 
# de cartera al final de cada mes junto con los honorarios
for i in range(len(vector_fecha)-1):
    fecha1 = dt.strptime(vector_fecha[i], '%Y-%m-%d')
    fecha2 = dt.strptime(vector_fecha[i+1], '%Y-%m-%d')

    dias = (fecha1 - fecha2).days
    
    # Transformamos las fechas nuevamente a formato string
    fecha1 = dt.strftime(fecha1, '%Y-%m-%d')
    
    # Calculamos el rendimiento
    if alyc == 'Bull':
        rendimiento = fc.rendimientos_bruto_neto(alyc = alyc,
                                                  usuario = usuario,
                                                  numero_interno = numero_interno,
                                                  fecha_cierre = fecha1,
                                                  dias = dias)
        
        valores_cartera.append(rendimiento.iloc[-2,0])

    elif alyc == 'Ieb':
        rendimiento = fc.rendimientos_bruto_neto_ieb(alyc = alyc,
                                                    usuario = usuario,
                                                    numero_interno = numero_interno,
                                                    fecha_cierre = fecha1,
                                                    dias = dias)
        
        valores_cartera.append(rendimiento.iloc[-2,0])

    elif alyc == 'Balanz':
        rendimiento = fc.rendimientos_bruto_neto_bal(alyc = alyc,
                                                      usuario = usuario,
                                                      numero_interno = numero_interno,
                                                      fecha_cierre = fecha1,
                                                      dias = dias)
        
        valores_cartera.append(rendimiento.iloc[-2,0])

    indice = len(vector_fecha) - i - 2 
    
    tabla_rendimientos.iloc[indice, 1] = rendimiento.iloc[0,0]
    tabla_rendimientos.iloc[indice, 2] = rendimiento.iloc[0,1]
    tabla_rendimientos.iloc[indice, 3] = round(float(rendimiento.iloc[-1,0]),2)


# Obtenemos el valor final de la cartera
valor_final = round(float(valores_cartera[0]),2)


# Creamos lista de honorarios para guardar los honorarios y poder calcular el 
# rendimiento neto del trimestre
lista_honorarios = []

for i in tabla_rendimientos['Honorarios*']:
    lista_honorarios.append(i)


# Funcion para formatear los valores, pasandolos a tipo string con simbolo en pesos 
# y separador de miles.
def formatear_pesos(valor):
    return f"$ {valor:,.1f}"


# Aplicar la funco³n anterior
tabla_rendimientos['Honorarios*'] = tabla_rendimientos['Honorarios*'].apply(formatear_pesos)


# Creamos los plazos de cada monto de honorarios para calcular el rendimiento neto 
# del trimestre
fecha_honorarios = []

for i in range(len(vector_fecha)-1):
    mes_bis = vector_fecha[i][5:7]
    anio_bis = vector_fecha[i][:4]
    
    # Se controlan los dias de corte por el mes de febrero.
    if mes_bis == '02':  
        # Este bloque anidado de try except es para que el mes de febrero 
        # permita dias de corte con 28 y 29 dias, siempre que corresponda.
        try: # Se intenta colocar febrero con dias igual a 'dia_corte'
            fecha_corte = f'{anio_bis}-{mes_bis}-{dia_corte}'
            fecha_corte = dt.strptime(fecha_corte, '%Y-%m-%d')
        
        except:
            try: # Se intenta colocar febrero de 29 dias
                fecha_corte = f'{anio_bis}-{mes_bis}-29'
                fecha_corte = dt.strptime(fecha_corte, '%Y-%m-%d')
                
            except: # Se coloca febrero de 28 dias como ultima opcion
                fecha_corte = f'{anio_bis}-{mes_bis}-28'
                fecha_corte = dt.strptime(fecha_corte, '%Y-%m-%d')
            
    # Se controlan los dias de corte para meses con 30 dias
    elif (dia_corte == 31 ) and ((mes_bis == '04') or (mes_bis == '06') or (mes_bis == '09'
                                                                    ) or (mes_bis == '11')):
        fecha_corte = f'{anio_bis}-{mes_bis}-30'
        fecha_corte = dt.strptime(fecha_corte, '%Y-%m-%d')
        
    # Caso donde no se requiere ningun control
    else:
        fecha_corte = f'{anio_bis}-{mes_bis}-{dia_corte}'
        fecha_corte = dt.strptime(fecha_corte, '%Y-%m-%d')
    
    fecha_honorarios.append(fecha_corte)
    

# Convertimos las fechas que estan en string en tipo datetime.
for i in range(len(tabla_rendimientos)):
    j = tabla_rendimientos.iloc[i,0]
    
    j = dt.strptime(j, '%Y-%m-%d')
    
    tabla_rendimientos.iloc[i,0] = j

tabla_rendimientos.set_index('Fechas',inplace = True)    


# Calculamos los plazos que corresponden a cada honorario
lista_plazo_honorarios = []

for i in fecha_honorarios:
    if tabla_rendimientos.index[-1] < i:
        plazo_honorario = 0
    
    else:
        plazo_honorario = (tabla_rendimientos.index[-1] - i).days  
        
    lista_plazo_honorarios.append(plazo_honorario)
    
# lista_plazo_honorarios = lista_plazo_honorarios[:-1].copy() # Recortamos la lista
lista_plazo_honorarios = lista_plazo_honorarios[::-1].copy() # Invertimos la lista


# Creamos la tabla con el valor final, inicial, y la ganancia absoluta.
# Creamos la tabla que contendra estos datos
tabla_valor = pd.DataFrame()

fecha_cierre = dt.strptime(fecha_cierre, "%Y-%m-%d")
fecha_inicial = dt.strptime(fecha_inicial , "%Y-%m-%d")

if fecha_inicial > alta:
    fecha_inicial = dt.strftime(fecha_inicial, "%Y-%m-%d")
    tabla_valor[fecha_inicial] = str(0)
    fecha_inicial = dt.strptime(fecha_inicial , "%Y-%m-%d")

elif fecha_inicial <= alta:
    alta = dt.strftime(alta, "%Y-%m-%d")
    tabla_valor[alta] = str(0)
    alta = dt.strptime(alta, "%Y-%m-%d")
    
tabla_valor['Aportes netos'] = str(0)


if type(baja) != type('Cliente activo'):  
    if fecha_cierre < baja:
        fecha_cierre = dt.strftime(fecha_cierre, "%Y-%m-%d")
        tabla_valor[fecha_cierre] = str(0)
        fecha_cierre = dt.strptime(fecha_cierre, "%Y-%m-%d")
        
    elif fecha_cierre >= baja:
        baja = dt.strftime(baja, "%Y-%m-%d")
        tabla_valor[baja] = str(0)
        baja = dt.strptime(baja, "%Y-%m-%d")
    
elif baja == 'Cliente activo':
    fecha_cierre = dt.strftime(fecha_cierre, "%Y-%m-%d")
    tabla_valor[fecha_cierre] = str(0)
    fecha_cierre = dt.strptime(fecha_cierre, "%Y-%m-%d")
    
tabla_valor[' '] = str(0)

tabla_valor.loc[0] = str(0)
tabla_valor.loc[1] = str(0)

tabla_valor.iloc[0,0] = 'Valor inicial'
tabla_valor.iloc[0,1] = ''
tabla_valor.iloc[0,2] = 'Valor Final'
tabla_valor.iloc[0,3] = 'Ganancia Neta*'

tabla_valor.iloc[1,0] = round(valor_inicial,2)
tabla_valor.iloc[1,1] = round(-1*(retiros_totales - depositos_totales),2)
tabla_valor.iloc[1,2] = round(valor_final,2)
tabla_valor.iloc[1,3] = round(valor_final - valor_inicial - depositos_totales + retiros_totales,2)

tabla_valor.iloc[1] = tabla_valor.iloc[1].apply(lambda x: f"$ {x:,}")

fecha_cierre = dt.strftime(fecha_cierre, "%Y-%m-%d")
fecha_inicial = dt.strftime(fecha_inicial , "%Y-%m-%d")



  
# ----------------------------  TERCERA PARTE ---------------------------------
# Actualizamos la serie 
# -----------------------------------------------------------------------------
# Importamos el excel con el precio del dolar mep y la tasa del plazo fijo
# Transformamos este archivo de acuerdo con las fechas de la 'tabla de rendimientos'



# -------------------- BLOQUE 1 de 2 PRECIO DOLAR MEP -------------------------
# Transformamos momentaneamente en datetime los string del 'vector_fecha_bis'
for i in range(len(vector_fecha_bis)):
    vector_fecha_bis[i] = dt.strptime(vector_fecha_bis[i], '%Y-%m-%d')


# Creamos una lista con las fechas que si existen en la serie de dolar mep, y que 
# son o estan cercanas a las fechas de la tabla de rendimientos
lista_fecha_mep = []

for j in vector_fecha_bis:
    fecha_precio_mep = j
    
    for i in range(60):
        
        if len(excel_usd.loc[excel_usd.index == (fecha_precio_mep - timedelta(days = i))]) == 0:
            fecha_precio_mep2 = fecha_precio_mep - timedelta(days = i)
            
        else:
            fecha_precio_mep2 = fecha_precio_mep - timedelta(days = i)
        
        if len(excel_usd.loc[excel_usd.index == fecha_precio_mep2]) == 1:
            break   
    
    lista_fecha_mep.append(fecha_precio_mep2)

# Buscamos el precio del mep para cada momento de la lista 'lista_fecha_mep'
lista_precio_mep = []

for i in lista_fecha_mep:
    precio_mep = excel_usd.loc[i,'MEP']
    lista_precio_mep.append(precio_mep)
    


# -------------------- BLOQUE 2 de 2 TASAS PLAZO FIJO -------------------------
# Ahora realizamos el mismo tipo de busqueda pero con las tasas del plazo fijo
# Creamos una lista con las fechas que si existen en la serie de plazo fijo, y 
# que son o estan cercanas a las fechas de la tabla de rendimientos
lista_fecha_pf = []

for j in vector_fecha_bis:
    fecha_pf = j
    
    for i in range(60):
        
        if len(excel_pf.loc[excel_pf.index == (fecha_pf - timedelta(days = i))]) == 0:
            fecha_pf2 = fecha_pf - timedelta(days = i)
            
        else:
            fecha_pf2 = fecha_pf - timedelta(days = i)
        
        if len(excel_pf.loc[excel_pf.index == fecha_pf2]) == 1:
            break   
    
    lista_fecha_pf.append(fecha_pf2)

# Buscamos la tasa del plazo fijo para cada momento de la lista 'lista_fecha_pf'
lista_tasa_pf = []

for i in lista_fecha_pf:
    tasa_pf = excel_pf.loc[i,'TNA']
    lista_tasa_pf.append(tasa_pf * 30 / 365 / 100)


# -----------------------------------------------------------------------------
# Buscamos el archivo que contiene las series de rendimiento de la cartera, de 
# la variacion del dolar mep, y del rendimiento del plazo fijo.
try: # Si el excel existe, lo importamos
    serie = pd.read_excel(f'{ubicacion_archivo}/{serie_rend}.xlsx')
    serie.set_index('Mes',inplace = True)
    
    # Colocamos la variacion del dolar mep y del plazo fijo
    tabla_mep_pf = pd.DataFrame()
    tabla_mep_pf['Mes'] = float(0)
    tabla_mep_pf['Mes'] = pd.to_datetime(tabla_mep_pf['Mes'])
    tabla_mep_pf['MEP'] = float(0)
    tabla_mep_pf['Plazo fijo'] = float(0)
    
    # Colocamos la tasa del plazo fijo, corrigiendolas por la cantidad de dias
    # de cada subperiodo (recordar que existen altas y bajas que provocan 
    # periodos menores a 30 dias)
    plazo_tasa_inicial = (vector_fecha_bis[1] - vector_fecha_bis[0]).days
    plazo_tasa_final = (vector_fecha_bis[-1] - vector_fecha_bis[-2]).days
        
    tasa_inicial = (1 + lista_tasa_pf[0]) ** (plazo_tasa_inicial/30) - 1
    tasa_final = (1+ lista_tasa_pf[-2]) ** (plazo_tasa_final/30) -1
    
    lista_tasa_pf[0] = tasa_inicial
    lista_tasa_pf[-2] = tasa_final
    
    lista_tasa_pf = lista_tasa_pf[:-1].copy()

    # Armamos la tabla de precio mep y tasa pf vinculada a las fechas
    for i in range(len(lista_precio_mep)-1):
        precio_mep2 = lista_precio_mep[i]
        precio_mep3 = lista_precio_mep[i+1]
        variacion = precio_mep3 /precio_mep2 - 1 
        
        tabla_mep_pf.loc[i,'Mes'] = vector_fecha_bis[i+1]
        tabla_mep_pf.loc[i,'MEP'] = variacion
        tabla_mep_pf.loc[i,'Plazo fijo'] = lista_tasa_pf[i]    
             
    tabla_mep_pf.set_index('Mes', inplace = True)

    # Colocamos todo en la tabla 'serie'
    for i in tabla_rendimientos.index:
        serie.loc[i,'Cartera'] = tabla_rendimientos.loc[i,'Bruto']
        serie.loc[i,'Dolar MEP'] = tabla_mep_pf.loc[i,'MEP']
        serie.loc[i,'Plazo fijo'] = tabla_mep_pf.loc[i,'Plazo fijo']

    # Tomamos la mascara para controlar que no se repitan fechas finales
    # Caso donde el cliente no baja su cuenta.
    if baja == 'Cliente activo':
        serie = serie.loc[serie.index <= fecha_cierre].copy()
        
    # Caso donde el cliente baja su cuenta 
    else:      
        if fecha_cierre <= baja:
            serie = serie.loc[serie.index <= fecha_cierre].copy()
            
        elif fecha_cierre > baja: 
            fechas = serie.loc[serie.index <= baja].copy()


except: # Si no existe el archivo excel, creamos el dataframe
    serie = pd.DataFrame()
    serie['Mes'] = float(0)
    serie['Mes'] = pd.to_datetime(serie['Mes'])
    serie['Cartera'] = float(0)
    serie['Dolar MEP'] = float(0)
    serie['Plazo fijo'] = float(0)
    
    # Colocamos las fechas y el rendimiento de la cartera
    for i in range(len(tabla_rendimientos)):
        serie.loc[i,'Mes'] = tabla_rendimientos.index[i]
        serie.loc[i,'Cartera'] = tabla_rendimientos.iloc[i,0]

    # Colocamos la variacion del dolar mep
    for i in range(len(lista_precio_mep)-1):
        precio_mep2 = lista_precio_mep[i]
        precio_mep3 = lista_precio_mep[i+1]
        variacion = precio_mep3 /precio_mep2 - 1
        
        serie.loc[i,'Dolar MEP'] = variacion
    
    # Colocamos la tasa del plazo fijo, corrigiendolas por la cantidad de dias
    # de cada subperiodo (recordar que existen altas y bajas que provocan 
    # periodos menores a 30 dias)
    plazo_tasa_inicial = (vector_fecha_bis[1] - vector_fecha_bis[0]).days
    plazo_tasa_final = (vector_fecha_bis[-1] - vector_fecha_bis[-2]).days
        
    tasa_inicial = (1 + lista_tasa_pf[0]) ** (plazo_tasa_inicial/30) - 1
    tasa_final = (1+ lista_tasa_pf[-2]) ** (plazo_tasa_final/30) -1
    
    lista_tasa_pf[0] = tasa_inicial
    lista_tasa_pf[-2] = tasa_final
    
    lista_tasa_pf = lista_tasa_pf[:-1].copy()
    
    for i in range(len(lista_tasa_pf)):
        serie.loc[i,'Plazo fijo'] = lista_tasa_pf[i]
    
    serie.set_index('Mes', inplace = True)
    
    # Tomamos la mascara para controlar que no se repitan fechas finales
    # Caso donde el cliente no baja su cuenta.
    fecha_cierre = dt.strptime(fecha_cierre, '%Y-%m-%d')
    
    if baja == 'Cliente activo':
        serie = serie.loc[serie.index <= fecha_cierre].copy()
        
    # Caso donde el cliente baja su cuenta 
    else:      
        if fecha_cierre <= baja:
            serie = serie.loc[serie.index <= fecha_cierre].copy()
            
        elif fecha_cierre > baja: 
            fechas = serie.loc[serie.index <= baja].copy()
            
    fecha_cierre = dt.strftime(fecha_cierre, '%Y-%m-%d')


# Transformamos nuevamente en string los datetime del 'vector_fecha_bis'. Hacemos
# lo mismo con la fecha inicial y la fecha de cierre
for i in range(len(vector_fecha_bis)):
    vector_fecha_bis[i] = dt.strftime(vector_fecha_bis[i], '%Y-%m-%d')

 
# Guardamos la tabla serie con los datos nuevos en la carpeta del cliente
serie.to_excel(f'{ubicacion_archivo}/Series.xlsx', 
                index = True, engine = 'openpyxl') 


# # Tomamos la serie hasta la fecha de cierre
# fecha_cierre_bis = dt.strptime(fecha_cierre, "%Y-%m-%d") # ESTO NO VA.

# serie = serie.loc[serie.index <= fecha_cierre_bis].copy() # ESTO NO VA.




# ----------------------------  CUARTA PARTE ----------------------------------
# Se construyen las graficas: 1) de linea, 2) de torta, y 3) de barra
# -----------------------------------------------------------------------------
# 1) Antes de graficar colocamos el momento inicial en el dataframe 'serie', que 
# representa el momento donde el cliente comenzo a invertir con nosotros. 
# 2) Luego construimos las columnas y los datos que pretendemos graficar.
# Armamos la tabla basica directamente util para la grafica de lineas
fecha_comienzo = alta 

serie.Cartera = serie.Cartera + 1 
serie['Dolar MEP'] = serie['Dolar MEP'] + 1 
serie['Plazo fijo'] = serie['Plazo fijo'] + 1 

serie['Cartera en pesos'] = serie.Cartera.cumprod()
serie['Cartera en usd'] = serie.Cartera / serie['Dolar MEP'] 
serie['var_usd'] = serie['Cartera en usd'] - 1
serie['Cartera en usd'] = serie['Cartera en usd'].cumprod()

serie.loc[fecha_comienzo] = 0
serie.sort_index(inplace = True)

serie.loc[fecha_comienzo, 'Cartera en pesos'] = 1
serie.loc[fecha_comienzo, 'Cartera en usd'] = 1

serie['Cartera en pesos'] = serie['Cartera en pesos'] * 1_000
serie['Cartera en usd'] = serie['Cartera en usd'] * 1_000 

serie['var_pesos'] = serie['Cartera'] - 1
serie.iloc[0,-1] = 0


# Armamos la tabla para la grafica de barras 
serieb = serie.iloc[1:].copy()




# -----------------------------------------------------------------------------
# Hacemos la grafica de TORTA
def grafica_torta():
    # Calculamos el rendimiento
    if alyc == 'Bull':
        torta = fc.grafica_composicion_bull(fecha_cierre = fecha_cierre, alyc = alyc, 
                                            numero_interno = numero_interno, 
                                            usuario = usuario)
        
    elif alyc == 'Ieb':
        torta = fc.grafica_composicion_ieb(fecha_cierre  = fecha_cierre, alyc = alyc, 
                                            numero_interno = numero_interno, 
                                            usuario = usuario)
    
    elif alyc == 'Balanz':
        torta = fc.grafica_composicion_balanz(fecha_cierre = fecha_cierre, alyc = alyc, 
                                              numero_interno = numero_interno, 
                                              usuario = usuario)
    img_stream = BytesIO()
    plt.savefig(img_stream, format='png')
    plt.close()
    img_stream.seek(0)
    
    return img_stream

torta_img_stream = grafica_torta()


# -----------------------------------------------------------------------------
# Hacemos la grafica de BARRA
def grafica_barra():
      # Hacemos la grafica de BARRAS
    fig, ax = plt.subplots(figsize=(14, 7))
    
    # Establecer el ancho de las barras
    bar_width = 0.35
    
    # Establecer la posiciÃ³n de las barras
    indices = np.arange(len(serieb))
    
    # Graficar las barras
    ax.bar(indices, serieb['var_pesos'], width=bar_width, label='En pesos', color='red')
    ax.bar(indices + bar_width, serieb['var_usd'], width=bar_width, label='En dolares', color='black')
    
    # Configurar el tÃ­tulo y las etiquetas
    ax.set_title(f'RENDIMIENTOS MENSUALES\n  Desde {dt.strftime(serieb.index[0], "%Y-%m-%d")}', 
                                      fontsize=29, fontweight='bold', color='red')
    ax.set_ylabel('Rendimientos', fontsize=20)
  
    # Ajustar ticks del eje X
    ax.set_xticks(indices + bar_width / 2)  
    ax.set_xticklabels(serieb.index.strftime('%b %y'), rotation=90, fontsize=21)
    
    # FunciÃ³n para formatear los valores del eje Y en porcentaje
    def percent_formatter(x, pos):
        return f'{round(x * 100,1)}%'  
    
    # Aplicar el formateador al eje Y izquierdo
    ax.yaxis.set_major_formatter(FuncFormatter(percent_formatter))
    
    # Aumentar el tamaÃ±o de los ticks del eje Y
    ax.tick_params(axis='y', labelsize=17)  # Cambia el tamaÃ±o de la etiqueta del eje Y
    
    # AÃ±adir leyenda
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.03), fontsize=17, ncol=2, 
              frameon = False)
    
    # Activar solo la cuadrÃ­cula horizontal
    ax.yaxis.grid(True)
    ax.xaxis.grid(False)
    
    # Ajustar el diseÃ±o
    plt.tight_layout()
    
    img_stream = BytesIO()
    plt.savefig(img_stream, format='png')
    plt.close()
    img_stream.seek(0)
    
    return img_stream

barra_img_stream = grafica_barra()


# -----------------------------------------------------------------------------
# Hacemos grafica de LINEA
def grafica_linea():
    # Crear la grÃ¡fica
    fig, ax1 = plt.subplots(figsize=(14, 7))
    
    
    # Graficar la primera lÃ­nea con el eje y izquierdo
    ax1.plot(serie.index, serie['Cartera en pesos'], color='red', label='En pesos') 
    ax1.set_ylabel('Pesos', color='Black', fontsize=20)
    ax1.tick_params(axis='y', labelcolor='Black', labelsize=17)
    
    
    # FunciÃ³n para formatear los valores del eje Y
    def currency_formatter(x, _):
        return f'${int(x):,}'.replace(',', '.')  # Formato con separador de miles y sÃ­mbolo
    
    # Aplicar el formateador al eje Y izquierdo
    ax1.yaxis.set_major_formatter(FuncFormatter(currency_formatter))
    
    
    # Crear el segundo eje y[]
    ax2 = ax1.twinx()  
    
    
    # Graficar la segunda lÃ­nea con el eje y derecho
    ax2.plot(serie.index, serie['Cartera en usd'], color='black', label='En dolares')
    ax2.set_ylabel('Dolares', color='black', fontsize=20)
    ax2.tick_params(axis='y', labelcolor='black', labelsize=17)
    
    
    # FunciÃ³n para formatear los valores del eje Y derecho
    def currency_formatter_usd(x, _):
        return f'USD{int(x):,}'  # Formato USD con dos decimales
    
    # Aplicar el formateador al eje Y derecho
    ax2.yaxis.set_major_formatter(FuncFormatter(currency_formatter_usd))
    
    
    # Establecer tÃ­tulo y etiquetas
    plt.title(f'EVOLUCION DEL VALOR DE LA CARTERA\n      Desde {dt.strftime(serieb.index[0], "%Y-%m-%d")}', 
                                                  color='red', fontweight='bold', 
                                                  fontname='Cambria', fontsize=29)
    plt.xticks(rotation=45)
    ax1.grid()

 
    # Formato de fechas en el eje X
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%b-%y'))  # Formato mmm-yy
    plt.xticks(rotation=45)
    ax1.grid()
    
    
    # Establecer todos los ticks del eje X y el tamaÃ±o de las abcisas
    ax1.set_xticks(serie.index)  
    ax1.tick_params(axis='x', labelsize=17)
    
    for label in ax1.get_xticklabels():
        label.set_rotation(90)
    
    
    # Anotar el valor inicial y final para la primera lÃ­nea
    ax1.text(serie.index[0], serie.loc[fecha_comienzo, 'Cartera en usd'] + 25, 
              f'${serie.iloc[0, 3]:,}'.replace(',', '.'), color='red', 
              verticalalignment='bottom', fontsize=13, ha='center')
    
    ax1.text(serie.index[-1], serie.iloc[-1, 3] + 30, 
              f'${round(serie.iloc[-1, 3],0):,}'.replace(',', '.'), color='red', 
              verticalalignment='bottom', fontsize=13, ha='center')
    
    
    # Anotar el valor inicial y final para la segunda lÃ­nea
    ax2.text(serie.index[0], serie.loc[fecha_comienzo, 'Cartera en usd'] - 10, 
              f'USD{serie.iloc[0, 4]:,}', color='Black', verticalalignment='bottom', 
              fontsize=13, ha='center')
    
    ax2.text(serie.index[-1], serie.iloc[-1, 4] - 10, 
              f'USD{round(serie.iloc[-1, 4],0):,}', color='Black', verticalalignment='bottom', 
              fontsize=13, ha='center')
    
    
    # Eliminar la lÃ­nea horizontal superior
    ax1.spines['top'].set_visible(False)
    
    
    # Agregar leyendas debajo del tÃ­tulo
    ax1.legend(loc='center', bbox_to_anchor=(0.6, 0.97), fancybox=True, ncol=2, 
                frameon=False, prop={'size': 17})
    ax2.legend(loc='center', bbox_to_anchor=(0.4, 0.97), fancybox=True, ncol=2, 
                frameon=False, prop={'size': 17})

    # Ajustar el espacio de la figura
    plt.subplots_adjust(right=0.85, bottom=0.2)  # Aumentar el espacio derecho y el horizontal

    img_stream = BytesIO()
    plt.savefig(img_stream, format='png')
    plt.close()
    img_stream.seek(0)
    
    return img_stream

linea_img_stream = grafica_linea()
     



# ----------------------------  QUINTA PARTE ----------------------------------
# Exportamos todo al reporte del cliente (archivo PDF)
# -----------------------------------------------------------------------------
# Crear un archivo PDF
pdf_filename = f'{destino_pdf}/Reporte {trimestre}T {anio} - {nombre_cliente} ({numero_cliente}).pdf'
document = SimpleDocTemplate(pdf_filename, pagesize=A4)


# -----------------------------------------------------------------------------
# COLOCANDO LA TABLA DE VALORES ABSOLUTOS EN EL PDF
# Convertir DataFrame a lista de listas
data_list = [tabla_valor.columns.tolist()] + tabla_valor.values.tolist()

# Cambiar el contenido de la celda que se va a fusionar
data_list[0][-1] = "Ganancia Neta*"  # Este serÃ¡ el contenido de la celda fusionada

# Crear un estilo especÃ­fico para el texto de la cuarta fila
custom_style = ParagraphStyle(
    'CustomStyle',
    parent=getSampleStyleSheet()['Normal'],
    fontName='Times-Roman',
    fontSize=8,
    alignment=0,  # AlineaciÃ³n centrada
)

# Agregar una cuarta fila con ajuste de texto
text = "* Neta de depositos y retiros de dinero (Aportes netos)."
data_list.append([Paragraph(text, custom_style), "", ""])  # Nueva fila con Paragraph

# Especificar el ancho de las columnas
col_widths = [90, 90, 110]  # Ajusta los valores segÃºn sea necesario

# Crear una tabla con anchura de columnas especificada
table1 = Table(data_list, colWidths=col_widths)

# Definir el estilo de la tabla
style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.darkgrey),  # Color de fondo del encabezado
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Color del texto del encabezado
    ('FONTNAME', (0, 0), (-1, 0), 'Times-Bold'),         # Negrita en el encabezado
    ('FONTSIZE', (0, 0), (-1, 0), 11),                   # TamaÃ±o de fuente del encabezado
    ('BACKGROUND', (0, 1), (-1, 1), colors.darkgrey),   # Color de fondo de la primera fila
    ('TEXTCOLOR', (0, 1), (-1, 1), colors.whitesmoke),  # Color del texto de la primera fila
    ('FONTNAME', (0, 1), (-1, 1), 'Times-Bold'),         # Negrita en la primera fila
    ('FONTSIZE', (0, 1), (-1, 1), 11),                   # TamaÃ±o de fuente de la primera fila
    ('BACKGROUND', (0, 2), (-1, 2), colors.white),       # Color de fondo de la segunda fila
    ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),        # Color del texto de la segunda fila
    ('FONTNAME', (0, 2), (-1, 2), 'Times-Bold'),         # Negrita en la segunda fila
    ('FONTSIZE', (0, 2), (-1, 2), 11),                   # TamaÃ±o de fuente de la segunda fila
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),               # Alinear al centro
    ('ALIGN', (-1, 0), (-1, 1), 'CENTER'),               # Centrar horizontalmente el contenido de 'Ganancia neta'
    ('VALIGN', (-1, 0), (-1, 1), 'MIDDLE'),              # Centrar verticalmente el contenido de 'Ganancia neta'
    ('ALIGN', (1, 0), (1, 1), 'CENTER'),                 # Centrar horizontalmente el contenido de 'Aporte neto'
    ('VALIGN', (1, 0), (1, 1), 'MIDDLE'),                # Centrar verticalmente el contenido de 'Aporte neto'
    ('BOTTOMPADDING', (0, 0), (-1, 0), 3),               # Espaciado inferior del encabezado
    ('BOTTOMPADDING', (0, 1), (-1, 1), 3),               # Espaciado inferior de la primera fila
    ('BOTTOMPADDING', (0, 2), (-1, 2), 3),               # Espaciado inferior de la segunda fila
    ('GRID', (0, 0), (-1, 2), 0.5, colors.black),        # AÃ±adir una cuadrÃ­cula mÃ¡s fina
])

# Aplicar el estilo a la tabla
table1.setStyle(style)

# Fusionar celdas usando setStyle
table1.setStyle(TableStyle([
    ('SPAN', (0, 3), (2, 3)),  # Fusionar las celdas de la tercera fila
    ('SPAN', (1, 0), (1, 1)),   # Fusionar la celda de la primera fila y segunda columna
    ('SPAN', (3, 0), (3, 1)),  # Fusionar la celda de la primera fila y cuarta columna
]))


# -----------------------------------------------------------------------------
# COLOCANDO TABLA DE MOVIMIENTOS EN EL PDF
# Modificamos el dataframe 'movimientos' para poder exportarlo al PDF
# Convertimos los datos de la columna fecha a un tipo string tipo date
movimientos.Fecha = movimientos.Fecha.dt.date

# Formatear las columnas 'Monto1' y 'Monto2' en pesos con separador de miles
def format_currency(value):
    return f"$ {value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

movimientos['Depositos'] = movimientos['Depositos'].apply(format_currency)
movimientos['Retiros'] = movimientos['Retiros'].apply(format_currency)


# -----------------------------------------------------------------------------
# Convertir DataFrame a lista de listas para la segunda tabla
data_list2 = [movimientos.columns.tolist()] + movimientos.values.tolist()

# Crear un estilo para la fila que actuarÃ¡ como tÃ­tulo
header_style = ParagraphStyle(
    name='HeaderStyle',
    fontName='Times-Roman',
    fontSize=9,
    alignment=0,  # AlineaciÃ³n centrada
    textColor=colors.black,
    backColor=colors.white,
)

# Crear la fila de tÃ­tulo para la segunda tabla
title_text = "Movimientos en pesos y dolares pesificados al MEP"
title_row = [Paragraph(title_text, header_style), "", ""]

# Insertar la fila de tÃ­tulo al principio de la tabla
data_list2.insert(0, title_row)

# Especificar el ancho de las columnas de la segunda tabla
col_widths2 = [75, 75, 75]  # Ajusta los valores segÃºn sea necesario

# Crear la segunda tabla
table2 = Table(data_list2, colWidths=col_widths2)

# Definir el estilo de la segunda tabla
style2 = TableStyle([
    ('BACKGROUND', (0, 1), (-1, 1), colors.darkgrey),  # Color de fondo fila 1
    ('TEXTCOLOR', (0, 1), (-1, 1), colors.whitesmoke), # Color del texto fila 1
    ('FONTNAME', (0, 1), (-1, 1), 'Times-Bold'),      # Negrita fila 1
    ('FONTSIZE', (0, 1), (-1, 1), 11),                 # TamaÃ±o de fuente fila 1
    
    ('BACKGROUND', (0, 2), (-1, 2), colors.white),     # Color de fondo fila 2
    ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),      # Color del texto fila 2
    ('FONTNAME', (0, 2), (-1, 2), 'Times-Roman'),      # Negrita fila 2
    ('FONTSIZE', (0, 2), (-1, 2), 11),                 # TamaÃ±o de fuente fila 2

    ('GRID', (0, 1), (-1, -1), 0.5, colors.black),       # AÃ±adir una cuadrÃ­cula
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),               # Alinear al centro
])

# Aplicar el estilo a la segunda tabla
table2.setStyle(style2)

# Fusionar celdas usando setStyle
table2.setStyle(TableStyle([
    ('SPAN', (0, 0), (-1, 0)),  # Fusionar las celdas de la tercera fila
]))



# -----------------------------------------------------------------------------
# COLOCANDO TABLA DE RENDIMIENTOS EN EL PDF
# PRIMERO trabajamos la tabla de rendimientos, preparandola para exportarla al PDF 
# Primero: Generamos la columna con los meses y anios (mmm-yy)
tabla_rendimientos.reset_index(inplace = True)
tabla_rendimientos['Mes'] = str(0)
tabla_rendimientos.set_index('Mes',inplace = True)
tabla_rendimientos.reset_index(inplace = True)
tabla_rendimientos['Mes'] = tabla_rendimientos['Fechas'].dt.strftime('%b-%y') 
tabla_rendimientos.drop('Fechas', inplace = True, axis = 1)


# Segundo: traducimos los rendimientos a un porcentaje.
for i in range(len(tabla_rendimientos)):
    tabla_rendimientos.loc[tabla_rendimientos.index[i],'Bruto'] = round(tabla_rendimientos.loc[tabla_rendimientos.index[i],"Bruto"] * 100,1)  
    tabla_rendimientos.loc[tabla_rendimientos.index[i],'Neto'] = round(tabla_rendimientos.loc[tabla_rendimientos.index[i],"Neto"] * 100,1) 

tabla_rendimientos['Neto'] = tabla_rendimientos['Neto'].astype(str) 
tabla_rendimientos['Bruto'] = tabla_rendimientos['Bruto'].astype(str) 
for i in range(len(tabla_rendimientos)):
    tabla_rendimientos.loc[tabla_rendimientos.index[i], 'Neto'] = tabla_rendimientos.loc[tabla_rendimientos.index[i], 'Neto'] + ' %'
    tabla_rendimientos.loc[tabla_rendimientos.index[i], 'Bruto'] = tabla_rendimientos.loc[tabla_rendimientos.index[i], 'Bruto'] + ' %'


# Colocamos la ultima fila, el rendimiento trimestral.  
tabla_rendimientos.loc[len(tabla_rendimientos)] = float(0)
tabla_rendimientos.loc[len(tabla_rendimientos)-1,'Mes'] = 'Período'


# Calculamos el rendimiento para el trimestre
fecha_cierre_movimientos = dt.strptime(fecha_cierre_movimientos, "%Y-%m-%d")
fecha_inicial_movimientos = dt.strptime(fecha_inicial_movimientos , "%Y-%m-%d")

fecha1 = fecha_cierre_movimientos
fecha2 = fecha_inicial_movimientos

plazo = (fecha1 - fecha2).days

fecha1 = dt.strftime(fecha1, '%Y-%m-%d')

fecha_cierre_movimientos = dt.strftime(fecha_cierre_movimientos, "%Y-%m-%d")
fecha_inicial_movimientos = dt.strftime(fecha_inicial_movimientos , "%Y-%m-%d")


# Calculamos el rendimiento bruto y neto del trimestre para los clientes de bull
try:
    rendimiento_neto = fc.rendimiento_neto(usuario = usuario, 
                                            numero_interno = numero_interno, 
                                            fecha_cierre = fecha_cierre_movimientos, 
                                            fecha_inicial = fecha_inicial_movimientos, 
                                            dias = plazo, 
                                            lista_honorarios = lista_honorarios,
                                            lista_plazo_honorarios = lista_plazo_honorarios, 
                                            valor_final = valor_final, 
                                            valor_inicial = valor_inicial,                                 
                                            alyc = alyc)
except:
    rendimiento_neto = float(0)


if alyc == 'Bull':    
    rendimiento_bruto = fc.rendimientos_bruto_neto(alyc = alyc,
                                                    usuario = usuario,
                                                    numero_interno = numero_interno,
                                                    fecha_cierre = fecha1,
                                                    dias = plazo)
    
elif alyc == 'Ieb': 
    rendimiento_bruto = fc.rendimientos_bruto_neto_ieb(alyc = alyc,
                                                        usuario = usuario,
                                                        numero_interno = numero_interno,
                                                        fecha_cierre = fecha1,
                                                        dias = plazo)

elif alyc == 'Balanz':    
    rendimiento_bruto = fc.rendimientos_bruto_neto_bal(alyc = alyc,
                                                        usuario = usuario,
                                                        numero_interno = numero_interno,
                                                        fecha_cierre = fecha1,
                                                        dias = plazo)

tabla_rendimientos.iloc[-1, 1] = f'{round(rendimiento_bruto.iloc[0,0] * 100,1)} %' 
tabla_rendimientos.iloc[-1, 2] = f'{round(rendimiento_neto * 100,1)} %'




# ---------------------------------------------------------------------------
# AHORA SI PASAMOS A EXPORTAR LA TABLA AL PDF
# Convertir DataFrame a lista de listas para la tercera tabla
data_list3 = [tabla_rendimientos.columns.tolist()] + tabla_rendimientos.values.tolist()

# Crear la fila de tÃ­tulos para la tercera tabla
title_text3 = " "
title_row3 = [Paragraph(title_text3, ParagraphStyle('HeaderStyle', fontName='Times-Bold', 
                                                    fontSize=11, alignment=1, 
                                                    textColor=colors.whitesmoke, 
                                                    backColor=colors.darkgrey)), "", ""]

# Insertar la fila de tÃ­tulo al principio de la tabla
data_list3.insert(1, title_row3)

# Insertamos los textos 'Bruto', 'Neto', y 'Rendimiento'
data_list3[1][1] = Paragraph("Bruto", ParagraphStyle(
    name='RendimientoStyle',
    fontName='Times-Bold',
    fontSize=11,
    textColor=colors.whitesmoke,
    alignment=1  # 1 para centrar
))

data_list3[1][2] = Paragraph("Neto", ParagraphStyle(
    name='RendimientoStyle',
    fontName='Times-Bold',
    fontSize=11,
    textColor=colors.whitesmoke,
    alignment=1  # 1 para centrar
))

data_list3[0][1] = Paragraph("Rendimiento", ParagraphStyle(
    name='RendimientoStyle',
    fontName='Times-Bold',
    fontSize=11,
    textColor=colors.whitesmoke,
    alignment=1  # 1 para centrar
))

# Agregar una cuarta fila con ajuste de texto
text = "* Calculados sobre la base de https://PAGINA WEB al cierre de cada mes. Puede no coincidir con el calculo a la fecha de facturacion."
data_list3.append([Paragraph(text, custom_style), "", ""])  # Nueva fila con Paragraph

# Definir el estilo de la tercera tabla (la matriz se lee 'columna - fila')
style3 = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 1), colors.darkgrey),  # Color de fondo dos primeras filas
    ('TEXTCOLOR', (0, 0), (-1, 1), colors.whitesmoke), # Color del texto dos primeras filas
    ('FONTNAME', (0, 0), (-1, 1), 'Times-Bold'),       # Tipo de fuente dos primeras filas
    ('FONTSIZE', (0, 0), (-1, 1), 11),                 # TamaÃ±o de fuente dos primeras filas
    
    ('BACKGROUND', (0, 2), (-1, -1), colors.white),      # Color de fondo del resto
    ('TEXTCOLOR', (0, 2), (-1, -1), colors.black),       # Color del texto del resto
    ('FONTNAME', (0, 2), (-1, -1), 'Times-Roman'),       # Tipo de fuented del resto
    ('FONTSIZE', (0, 2), (-1, -1), 11),                  # TamaÃ±o de fuente del resto
    
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),      # AÃ±adir una cuadrÃ­cula
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),              # Alinear al centro
    ('GRID', (3, -2), (3, -2), 0.5, colors.white),        # AÃ±adir una cuadrÃ­cula
    ('GRID', (3, -3), (3, -3), 0.5, colors.black),       # AÃ±adir una cuadrÃ­cula
    
    ('GRID', (0, -1), (-1, -1), 0.5, colors.white),       # AÃ±adir una cuadrÃ­cula ultima fila
    ('GRID', (0, -2), (-2, -2), 0.5, colors.black),       # AÃ±adir una cuadrÃ­cula penultima fila
    
    ('TEXTCOLOR', (0, -2), (-1, -2), colors.black),        # Color del texto penultima fila
    ('FONTNAME', (0, -2), (-1, -2), 'Times-Bold'),       # Tipo de fuente penultima fila
    ('TEXTCOLOR', (3, -2), (3, -2), colors.white),        # Color del texto penultima celda
    
    ('FONTSIZE', (0, -1), (-1, -1), 8),                  # TamaÃ±o de fuente ultima fila
    
])
# ¿Cómo leer este codigo? Desde (x1, y1) hasta (x2, y2), donde x1 y x2 indican 
# las columnas inicial y final e y1 e y2 indican las filas inicial y final.

# Crear la tercera tabla y establecer su tamaÃ±o
# Especificar el ancho de las columnas de la segunda tabla
col_widths3 = [100, 75, 75, 100]  

table3 = Table(data_list3, colWidths=col_widths3)
table3.setStyle(style3)

# Fusionar celdas usando setStyle
table3.setStyle(TableStyle([
    ('SPAN', (0, 0), (0, 1)),  # Fusionar celdas 'Mes'
    ('SPAN', (-1, 0), (-1, 1)),  # Fusionar celdas 'Honorarios*'
    ('SPAN', (1, 0), (2, 0)),  # Fusionar celdas 'Rendimientos'
    ('SPAN', (0, -1), (-1, -1)),  # Fusionar celdas de la ultima fila
    
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Centrar horizontalmente
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centrar verticalmente
]))



# Crear un espaciador para aumentar el espacio entre las tablas
space_between_tables1 = Spacer(1, 20)  # 1 unidad de ancho, 20 puntos de alto
space_between_tables2 = Spacer(1, 30)  # 1 unidad de ancho, 20 puntos de alto
space_between_graficas = Spacer(1, 30) # Espacio entre graficas
 
# -----------------------------------------------------------------------------
# Sumamos las GRAFICAS
# Obtener dimensiones originales
pie_chart_image = Image(torta_img_stream)
pie_chart_image.drawHeight = 275 # Ajustar altura
pie_chart_image.drawWidth = 275   # Ajustar anchura

barra_chart_image = Image(barra_img_stream)
barra_chart_image.drawHeight = 275 # Ajustar altura
barra_chart_image.drawWidth = 450   # Ajustar anchura

linea_chart_image = Image(linea_img_stream)
linea_chart_image.drawHeight = 275 # Ajustar altura
linea_chart_image.drawWidth = 450   # Ajustar anchura


# -----------------------------------------------------------------------------
# Incorporamos el TEXTO E IMAGENES
styles = getSampleStyleSheet()

elementos = []

# -----------------------------------------------------------------------------
# Agregamos las IMAGENES
# IMAGEN PRIMERA PAGINA
image_path = f"{destino_pdf}/imagen_primera.png"  # Cambia 'tu_imagen.jpg' por la ruta de tu imagen
image = Image(image_path)

inch = 100

# Escalar la imagen manteniendo la relaciÃ³n de aspecto
max_width = 5 * inch  # Ancho mÃ¡ximo
max_height = 3 * inch  # Altura mÃ¡xima

# Obtener las dimensiones de la imagen
image.drawWidth, image.drawHeight = image.wrap(max_width, max_height)

# Comprobar si es necesario redimensionar
if image.drawWidth > max_width or image.drawHeight > max_height:
    ratio = min(max_width / image.drawWidth, max_height / image.drawHeight)
    image.drawWidth *= ratio
    image.drawHeight *= ratio

elementos.append(image)
elementos.append(Spacer(1, 2 * inch/10))  # Espaciador despuÃ©s de la imagen



# -----------------------------------------------------------------------------
# Incorporamos el TEXTO
# FunciÃ³n para leer texto desde un archivo de Word
def leer_texto_de_word(archivo_word):
    doc = Document(archivo_word)
    texto = []
    for parrafo in doc.paragraphs:
        texto.append(parrafo.text)
    return texto  # Devuelve una lista de pÃ¡rrafos


custom_style = ParagraphStyle(
    name='CustomStyle',
    fontName='Helvetica',
    fontSize=9,
    alignment=4,  # Justificado
    spaceAfter=12,  # Espacio despuÃ©s de cada pÃ¡rrafo
)

# Definir estilo para las Ãºltimas dos lÃ­neas
right_align_style = ParagraphStyle(
    name='RightAlignStyle',
    fontName='Helvetica',
    fontSize=9,
    alignment=2,  # Derecha
    spaceAfter=12,  # Espacio despuÃ©s de cada pÃ¡rrafo
)

# Importamos el archivo word
texto_word = leer_texto_de_word(f'{destino_pdf}/texto para inversor.docx')  

# AÃ±adir texto al PDF
for i, linea in enumerate(texto_word):
    if linea.strip():  # AsegÃºrate de que la lÃ­nea no estÃ© vacÃ­a
        # Comprobar si es la primera lÃ­nea o una de las Ãºltimas dos lÃ­neas
        if i == 0:
            # Poner en negrita
            elementos.append(Paragraph(f"<b>{linea}</b>", custom_style))
        elif i >= len(texto_word) - 2:
            # Ãltimas dos lÃ­neas en negrita y alineadas a la derecha
            elementos.append(Paragraph(f"<b>{linea}</b>", right_align_style))
        else:
            elementos.append(Paragraph(linea, custom_style))
        elementos.append(Spacer(1, 0.2 ))  # Espaciador para una lÃ­nea en blanco
        
   
# -----------------------------------------------------------------------------
# Agregar un salto de pÃ¡gina. Esto permite que al finalizar el texto, las tablas
# y gracias comiencen en una nueva pagina.
elementos.append(PageBreak())


# -----------------------------------------------------------------------------
# IMAGEN SEGUNDA
image_path = f"{destino_pdf}/imagen_segunda.png"  # Cambia 'tu_imagen.jpg' por la ruta de tu imagen
image = Image(image_path)

inch = 100

# Escalar la imagen manteniendo la relaciÃ³n de aspecto
max_width = 5 * inch  # Ancho mÃ¡ximo
max_height = 3 * inch  # Altura mÃ¡xima

# Obtener las dimensiones de la imagen
image.drawWidth, image.drawHeight = image.wrap(max_width, max_height)

# Comprobar si es necesario redimensionar
if image.drawWidth > max_width or image.drawHeight > max_height:
    ratio = min(max_width / image.drawWidth, max_height / image.drawHeight)
    image.drawWidth *= ratio
    image.drawHeight *= ratio

elementos.append(image)
elementos.append(Spacer(1, 2 * inch/200))  # Espaciador despuÃ©s de la imagen


# -----------------------------------------------------------------------------
# Unimos todos los elementos
# Incorporamos las tablas, las graficas, y los espacios entre las mismas.
elements = [space_between_tables2, 
            table1, space_between_tables2,           # TABLAS
            table3, space_between_tables1, 
            table2, space_between_graficas, 
            
            pie_chart_image, space_between_graficas, # GRAFICAS
            barra_chart_image, space_between_graficas, 
            linea_chart_image] 

for i in elements:
    elementos.append(i)


# -----------------------------------------------------------------------------
# Al PDF se le insertan todos los elementos construidos previamente.
document.build(elementos)



print('Valores absolutos')
print(tabla_valor)
print('')
print('Movimientos')
print(movimientos)
print('')
print('rendimientos')
print(tabla_rendimientos)










