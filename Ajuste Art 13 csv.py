import os , pandas as pd , numpy as np , openpyxl
from tkinter import filedialog

def Ajuste_Art13_CSV():

    Objetivo = filedialog.askdirectory(title = "Seleccionar carpeta de archivos") + "/"

    if Objetivo != "/":

        Objetivo = filedialog.askdirectory(title = "Seleccionar carpeta de archivos") + "/"

        # Listar los archivos de la carpeta 'Anual'
        archivos = os.listdir ( Objetivo )

        # Eliminar los archivos que no sean .csv
        archivos = [i for i in archivos if i.endswith(".csv")]

        #crear un dataframe vacio
        Consolidado = pd.DataFrame()

        #consolidar los archivos Excel en un solo dataframe
        for i in archivos:
            data = pd.read_csv( (Objetivo + i) , decimal="," , sep=";" , header = 0 , low_memory=False)
            #Crear una columna con el nombre del archivo
            data['Archivo'] = i
            Consolidado = pd.concat([Consolidado , data])
        del data , i , archivos

        #Eliminar columnas 'id' , 'idcuit' , 'idtipo_operacion' , 'idprovinciaiibb' , 'provincia' , 'idcentrocosto' , 'idcomprobanteorigina' , 'idclipro' , 'idcliprodomicilio' , 'idmodelo' , 'modelo' , 'codactividad' , 'numerohasta' , 'CAE_obtener' , 'CAE_intentos' , 'CAE_iderrorlog' , 'CAE_fecha' , 'CAE_numero' , 'CAE_fechavencimiento' , 'CAE_importetotal' , 'CAE_testing' , 'CAE_fechahorapedido' , 'CAE_fechahoraobtenido' , 'CAE_idusuariomodi' , 'FE_url' , 'FE_fechahoramodi' , 'FE_idusuariomodi' , 'FE_fechahoraenvio' , 'FE_idusuarioenvio' , 'memo_factura' , 'memo_remito' , 'memo_cheque' , 'idcuenta' , 'fechaalta' , 'idusuarioalta' , 'usuarioalta' , 'fechamodi' , 'idusuariomodi' , 'usuariomodi' , 'fechabaja' , 'idusuariobaja' , 'usuariobaja' , 'tipoespecial' , 'idcomprobanteventa' , 'identificador' de consolidado
        Consolidado = Consolidado.drop(['id' , 'idcuit' , 'idtipo_operacion' , 'idprovinciaiibb' , 'provincia' , 'idcentrocosto' , 'idcomprobanteorigina' , 'idclipro' , 'idcliprodomicilio' , 'idmodelo' , 'modelo' , 'codactividad' , 'numerohasta' , 'CAE_obtener' , 'CAE_intentos' , 'CAE_iderrorlog' , 'CAE_fecha' , 'CAE_numero' , 'CAE_fechavencimiento' , 'CAE_importetotal' , 'CAE_testing' , 'CAE_fechahorapedido' , 'CAE_fechahoraobtenido' , 'CAE_idusuariomodi' , 'FE_url' , 'FE_fechahoramodi' , 'FE_idusuariomodi' , 'FE_fechahoraenvio' , 'FE_idusuarioenvio' , 'memo_factura' , 'memo_remito' , 'memo_cheque' , 'idcuenta' , 'fechaalta' , 'idusuarioalta' , 'usuarioalta' , 'fechamodi' , 'idusuariomodi' , 'usuariomodi' , 'fechabaja' , 'idusuariobaja' , 'usuariobaja' , 'tipoespecial' , 'idcomprobanteventa' , 'identificador'] , axis = 1)

        #Rellenar las columnas 'montoneto' , 'montonogravado' , 'montoexento' , 'montoiva' , 'montoperc' , 'montoperc_nac' , 'montoperc_iibb' , 'montoimpmunic' , 'montoimpint' , 'montootrostrib' , 'montototal' con 0
        Consolidado['montoneto'] = Consolidado['montoneto'].fillna(0)
        Consolidado['montonogravado'] = Consolidado['montonogravado'].fillna(0)
        Consolidado['montoexento'] = Consolidado['montoexento'].fillna(0)
        Consolidado['montoiva'] = Consolidado['montoiva'].fillna(0)
        Consolidado['montoperc'] = Consolidado['montoperc'].fillna(0)
        Consolidado['montoperc_nac'] = Consolidado['montoperc_nac'].fillna(0)
        Consolidado['montoperc_iibb'] = Consolidado['montoperc_iibb'].fillna(0)
        Consolidado['montoimpmunic'] = Consolidado['montoimpmunic'].fillna(0)
        Consolidado['montoimpint'] = Consolidado['montoimpint'].fillna(0)
        Consolidado['montootrostrib'] = Consolidado['montootrostrib'].fillna(0)
        Consolidado['montototal'] = Consolidado['montototal'].fillna(0)


        #Crear columna de 'CLIENTE Art 13' teniendo en cuenta los caracteres de la columna 'Archivo' desde la posición 15 hasta el final (sin tener el cuenta' - Anual 2022.xlsx')
        Consolidado['CLIENTE Art 13'] = Consolidado['Archivo'].str.slice(15, -18)

        #crear columna de periodo en base a la columna 'fechaiva' (primero pasarla a formato datetime DD/MM/AAAA) y convertirla a formato AAAAMM
        Consolidado['periodo'] = pd.to_datetime(Consolidado['fechaiva'], format='%d/%m/%Y').dt.strftime('%Y%m')

        #Crear columnas auxiliares
        Consolidado['NG'] = Consolidado['montonogravado'] + Consolidado['montoexento']
        Consolidado['Neto + NG'] = Consolidado['montoneto'] + Consolidado['montonogravado'] + Consolidado['montoexento']

        #Crear dataframe con 'Compras' y 'Ventas' en la columna 'tipooperacion'
        Compras = Consolidado[Consolidado['tipooperacion'] == 'Compras']
        Ventas = Consolidado[Consolidado['tipooperacion'] == 'Ventas']


        #Crear tablas pivote de 'Compras' y 'Ventas' por 'CLIENTE Art 13' y 'periodo'
        Compras_TD = Compras.pivot_table(index = ['CLIENTE Art 13' , 'periodo'] , values = ['montoneto' , 'NG' , 'Neto + NG' , 'montoiva' ] , aggfunc = 'sum')
        Ventas_TD = Ventas.pivot_table(index = ['CLIENTE Art 13' , 'periodo'] , values = ['montoneto' , 'NG' , 'Neto + NG'] , aggfunc = 'sum')
        Ventas_Computable = Ventas.pivot_table(index = ['CLIENTE Art 13'] , values = ['montoneto' , 'Neto + NG'] , aggfunc = 'sum')

        #Crear columna de '% Computable' en base a la columna 'montoneto' / 'Neto + NG' de la tabla 'Ventas' y 'Ventas_Computable'
        Ventas_TD['% Computado'] = Ventas_TD['montoneto'] / Ventas_TD['Neto + NG']
        Ventas_Computable['% Computable'] = Ventas_Computable['montoneto'] / Ventas_Computable['Neto + NG']

        #Separar el indice de las tablas pivote en columnas
        Compras_TD.reset_index(inplace=True)
        Ventas_TD.reset_index(inplace=True)
        Ventas_Computable.reset_index(inplace=True)

        #crear columas AUX en Compas_TD y Ventas_TD
        Compras_TD['Aux'] = Compras_TD['CLIENTE Art 13'] + ' - '+ Compras_TD['periodo'].astype(str)
        Ventas_TD['Aux'] = Ventas_TD['CLIENTE Art 13'] + ' - ' + Ventas_TD['periodo'].astype(str)

        #Merge de Compras_TD y Ventas_TD en base a la columna 'Aux' y traer solamente la columna '% Computado' de Ventas_TD
        Compras_TD = pd.merge(  Compras_TD,
                                Ventas_TD[['Aux', '% Computado']],
                                left_on='Aux' ,
                                right_on='Aux',
                                how='left')

        #merge de Compras_TD y Ventas_Computable en base a la columna 'CLIENTE Art 13' y traer solamente la columna '% Computable' de Ventas_Computable
        Compras_TD = pd.merge(  Compras_TD,
                                Ventas_Computable[['CLIENTE Art 13', '% Computable']],
                                left_on='CLIENTE Art 13' ,
                                right_on='CLIENTE Art 13',
                                how='left')

        #Crear columna de diferencia entre '% Computado' y '% Computable' y multiplicarlo por el 'montoiva'
        Compras_TD['Diferencia'] = (Compras_TD['% Computable'] - Compras_TD['% Computado']) * Compras_TD['montoiva']

        #Crear pivot table de 'Diferencia' por 'CLIENTE Art 13'
        Diferencia = Compras_TD.pivot_table(index = ['CLIENTE Art 13'] , values = ['Diferencia'] , aggfunc = 'sum')

        #exportar las tablas pivote a excel en solapas distintas
        with pd.ExcelWriter(f'{Objetivo}Consolidado.xlsx') as writer:
            Compras_TD.to_excel(writer, sheet_name='Compras', index=False)
            Ventas_TD.to_excel(writer, sheet_name='Ventas' , index=False)
            Ventas_Computable.to_excel(writer, sheet_name='Ventas_Computable' , index=False)
            Diferencia.to_excel(writer, sheet_name='Resumen' , index=True)

if __name__ == '__main__':
    Ajuste_Art13_CSV()