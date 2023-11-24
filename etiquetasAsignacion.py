import pandas as pd



# Pedir al usuario que ingrese el nombre del archivo por consola
nombre_archivo = input("Ingrese el nombre del archivo Excel (sin la extensión): ").strip()

# Agregar la extensión al nombre del archivo
archivo_excel = f"{nombre_archivo}.xlsx"

# Intentar abrir el archivo Excel
try:
    excel_file = pd.ExcelFile(archivo_excel)
    nombres_de_hojas = excel_file.sheet_names
    print("Nombres de hojas en el archivo Excel:", nombres_de_hojas)

except FileNotFoundError:
    print(f"El archivo '{archivo_excel}' no se encontró. Por favor, verifique el nombre y la ruta del archivo.")
except Exception as e:
    print(f"Ocurrió un error al abrir el archivo Excel: {e}")


for hoja in nombres_de_hojas:
    df = pd.read_excel(io=archivo_excel, sheet_name=hoja)
    columnas=list(df)
    listaDlistas=[]
    for i in range(len(columnas)):
        filas=list(df[columnas[i]])
        new_filas = [x for x in filas if pd.isnull(x) == False]
        listaDlistas.append(new_filas)
    listatotal=[]
    for j in range(len(columnas)):
        for x in range(len(listaDlistas[j])):
            listatotal.append(listaDlistas[j][x])


    nueva_lista=[] #guardo los elementos ordenados 
    for i in range(len(listaDlistas)):
        sector= listaDlistas[i][0]
        piscina= listaDlistas[i][1]
        lista_aux=[]
        lista_iteracion=listaDlistas[i][2:]

        for j in range(len(lista_iteracion)):
            numero_motor= str(j+1)
            nuevo_elemento= f"{sector}-{piscina}-{numero_motor}"
            lista_aux.append(nuevo_elemento)
            lista_aux.append(lista_iteracion[j])
        lista_aux= [sector] + [piscina] + lista_aux
        nueva_lista.append(lista_aux)

    columnas_n= len(nueva_lista)

    datos_escribir= pd.DataFrame(nueva_lista).transpose()
    with pd.ExcelWriter(archivo_excel,
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                        datos_escribir.to_excel(writer, sheet_name=hoja, header=False, index=False, startcol= columnas_n + 3) 
    







