gitimport pandas as pd
import numpy as np
from tkinter import *
from tkinter import messagebox
import glob, os


#archivo = numero -revisar variable



#______________________________________Aniones______________________________________

data = pd.DataFrame(columns=('Anión', 'Resultado', 'Peso Molecular', 'Equivalente 1', 'Equivalente 2'))
col_types = {"Resultado":float, "Peso Molecular":float, "Equivalente 1":float, "Equivalente 2":float}
data.loc[len(data)] = ['Cl', 0.0, 35.5, 35.5,0.0]
data.loc[len(data)] = ['NO2-', 0.0, 46, 46,0.0]
data.loc[len(data)] = ['NO3', 0.0, 62, 62,0.0]
data.loc[len(data)] = ['F-', 0.0, 19, 19,0.0]
data.loc[len(data)] = ['HCO3-', 0.0, 61, 61,0.0]
data.loc[len(data)] = ['CO3-2', 0.0, 60, 30,0.0]
data.loc[len(data)] = ['SO4-2', 0.0, 96.06, 48.05,0.0]
data.loc[len(data)] = ['H2PO4-', 0.0, 97, 97,0.0]

#_____________________________________Cationes_____________________________________

data2 = pd.DataFrame(columns=('Catión', 'Resultado', 'Peso Molecular', 'Equivalente 1', 'Equivalente 2'))
col_types2 = {"Resultado":float, "Peso Molecular":float, "Equivalente 1":float, "Equivalente 2":float}
data2.loc[len(data2)] = ['Na+', 0.0, 23, 23,0.0]
data2.loc[len(data2)] = ['k+', 0.0, 39.1, 39.1,0.0]
data2.loc[len(data2)] = ['N-Nh3', 0.0, 18, 18,0.0]
data2.loc[len(data2)] = ['Mg+2', 0.0, 24.3, 12.15,0.0]
data2.loc[len(data2)] = ['Ca+2', 0.0, 40.1, 20.05,0.0]
data2.loc[len(data2)] = ['Fe+2', 0.0, 55.8, 27.9,0.0]
data2.loc[len(data2)] = ['Mn+2', 0.0, 54.9, 27.45,0.0]
data2.loc[len(data2)] = ['Al+3', 0.0, 27, 9,0.0]




Name = len(glob.glob(r"C:\Users\sandrapdiaz\Documents\*.xlsx"))

datos=[]

os.chdir(r'C:\Users\sandrapdiaz\Documents')
archivos = []
for file in glob.glob("*.xlsx"):
    archivos.append(file)

Producto = pd.DataFrame(columns=('Muestra', 'SA - Cond. %', 'D SC - SA %', 'STD en ppm', 'DT', 'Volumen DT', 'DCa', 'Volumen DCa', 'DMg', 'Volumen DMg', 'RAS', 'PSI'))
colum = list(Producto)
dat = []
    
for k in range(len(archivos)):

    
    print(archivos[k])
    iter = 1
    for p in range(len(pd.read_excel(archivos[k], sheet_name=None))):#Cambia sobre las hojas
        hoja = "Hoja"+str(iter)
        try:
            df = pd.read_excel(archivos[k], sheet_name = hoja)
            #df = pd.concat(pd.read_excel('1.xlsx', sheet_name = None), ignore_index = True) #str(k)+'.xlsx' concatena las hojas
        except FileNotFoundError:
            master = Tk()
            master.withdraw()#Permite que no se abra tra caja de diálogo de tk
            messagebox.showwarning(message="Ese archivo no existe en el directorio", title="Nombre Incorrecto")

#df.columns.values
#df
#len(df.columns)

#________________________ Corrección de valores______________________________________
        #print(df['Unnamed: 2'])

        
        col = 6
    
        for i in range(6, len(df.columns)):

            for j in range(len(df)):
                if df.loc[j,'Unnamed: 2'] == 'Cloruro\n':
                    #data.loc[0,'Resultado'] = df.loc[i, 'SGI HIDROCARBUROS']
                    data.loc[0,'Resultado'] = df.iloc[j, col]
                if df.loc[j,'Unnamed: 2'] == 'Alcalinidad total\n':
                    data.loc[4,'Resultado'] = (df.iloc[j, col]*(data.loc[4, 'Peso Molecular'])/100.0869)
                if df.loc[j,'Unnamed: 2'] == 'Carbonatos\n':
                    if df.iloc[j, col] == '<10':
                        data.loc[5,'Resultado'] = 10*(data.loc[5, 'Peso Molecular']/100.0869)
                    else:
                        data.loc[5,'Resultado'] = (df.iloc[j, col]*(data.loc[5, 'Peso Molecular'])/100.0869)
                if df.loc[j,'Unnamed: 2'] == 'Sulfato\n':
                    if df.iloc[j, col] == '<11':
                        data.loc[6,'Resultado'] = 11
                    else:
                        data.loc[6,'Resultado'] = (df.iloc[j, col]*(data.loc[6, 'Peso Molecular'])/100.0869)

            data['Equivalente 2'] = data['Resultado'].div(data['Equivalente 1'])


            for l in range(len(df)):
                if df.loc[l,'Unnamed: 2'] == 'Sodio\n en agua':
                    #data.loc[0,'Resultado'] = df.loc[i, 'SGI HIDROCARBUROS']
                    data2.loc[0,'Resultado'] = df.iloc[l, col]
                if df.loc[l,'Unnamed: 2'] == 'Potasio\n en agua':
                    #data.loc[0,'Resultado'] = df.loc[i, 'SGI HIDROCARBUROS']
                    data2.loc[1,'Resultado'] = df.iloc[l, col]
        
                if df.loc[l,'Unnamed: 2'] == 'Amoniaco\n':
                    data2.loc[2,'Resultado'] = (df.iloc[l, col]*(data2.loc[2, 'Peso Molecular'])/14)
    
                if df.loc[l,'Unnamed: 2'] == 'Magnesio \n en agua':
                    #data.loc[0,'Resultado'] = df.loc[i, 'SGI HIDROCARBUROS']
                    data2.loc[3,'Resultado'] = df.iloc[l, col]
    
                if df.loc[l,'Unnamed: 2'] == 'Calcio\n en agua':
                    #data.loc[0,'Resultado'] = df.loc[i, 'SGI HIDROCARBUROS']
                    data2.loc[4,'Resultado'] = df.iloc[l, col]
    
                if df.loc[l,'Unnamed: 2'] == 'Hierro \n en agua':
                    #data.loc[0,'Resultado'] = df.loc[i, 'SGI HIDROCARBUROS']
                    data2.loc[5,'Resultado'] = df.iloc[l, col]
    
                if df.loc[l,'Unnamed: 2'] == 'Manganeso\n en agua':
                    #data.loc[0,'Resultado'] = df.loc[i, 'SGI HIDROCARBUROS']
                    data2.loc[6,'Resultado'] = df.iloc[l, col]
     
                if df.loc[l,'Unnamed: 2'] == 'Aluminio\n en agua':
                    #data.loc[0,'Resultado'] = df.loc[i, 'SGI HIDROCARBUROS']
                    data2.loc[7,'Resultado'] = df.iloc[l, col]
    
            #print(data2)        
    
            data2['Equivalente 2'] = data2['Resultado'].div(data2['Equivalente 1'])
           
            #print(data2)
         
        
            Suma = data['Equivalente 2'].sum()
            alt = data2['Equivalente 2'].sum()
            for m in range(len(df)):
                if df.loc[m,'Unnamed: 2'] == 'Conductividad Eléctrica a 25 ºC\n':
                    comp = df.iloc[m, col] / 100
            rae= abs((Suma - comp) / comp)*100
            doe = ((alt - Suma) /  (Suma + alt))
            std = data['Resultado'].sum() + data2['Resultado'].sum()
            dt = (2.497*(data2.loc[4,'Resultado'])) + (4.118*(data2.loc[3,'Resultado']))
            dc = (2.497*(data2.loc[4,'Resultado']))
            dm = (4.118*(data2.loc[3,'Resultado']))
            RAS = (data2.loc[0,'Resultado']/23) / np.sqrt((((data2.loc[3,'Resultado']/12.15) + (data2.loc[4,'Resultado']/20.05))/2)) 
            PSI = (1.475*RAS)/(1+(0.0127*RAS))
            
            
            
            
            
            values = [df.iloc[4,col], rae , doe, std, dt, ((dt*100)/1002), dc, ((dc*100)/1002), dm, ((dm*100)/1002), RAS, PSI]
            ziped = zip(colum, values)
            a_dictionary = dict(ziped)
            dat.append(a_dictionary)
            
            
            
            col+=1# Cambia sobre todas las muestras
        iter+=1 #Incrementa el número de la hoja a revisar

Producto = Producto.append(dat,True) 
            
Producto.to_excel(r'Calculos.xlsx')