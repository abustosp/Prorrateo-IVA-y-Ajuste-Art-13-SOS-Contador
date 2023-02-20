import os , pandas as pd , numpy as np , openpyxl
from tkinter import filedialog

def Prorrateo_Mensual_XLS():

    Objetivo = filedialog.askdirectory(title = "Seleccionar carpeta de archivos") + "/"

    if Objetivo != "/":

        # Listar los archivos de la carpeta seleccionada
        archivos = os.listdir ( Objetivo )

        # Eliminar los archivos que no sean .xls
        archivos = [i for i in archivos if i.endswith(".xls")]

        #crear un dataframe vacio
        Consolidado = pd.DataFrame()

        #consolidar los archivos Excel en un solo dataframe
        for i in archivos:
            data = pd.read_html( (Objetivo + i) , decimal="," , thousands="" , header = 0 ,)[0]
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

        #Crear columna de 'CLIENTE' teniendo en cuenta los caracteres de la columna 'Archivo' desde la posición 15 hasta el final (sin tener el cuenta' - Anual 2022.xlsx')
        Consolidado['CLIENTE'] = Consolidado['Archivo'].str.slice(15, -18)

        #crear columna de periodo en base a la columna 'fechaiva' (primero pasarla a formato datetime DD/MM/AAAA) y convertirla a formato AAAAMM
        Consolidado['periodo'] = pd.to_datetime(Consolidado['fechaiva'], format='%d/%m/%Y').dt.strftime('%Y%m')

        #Crear columnas auxiliares
        Consolidado['NG'] = Consolidado['montonogravado'] + Consolidado['montoexento']
        Consolidado['Neto + NG'] = Consolidado['montoneto'] + Consolidado['montonogravado'] + Consolidado['montoexento']

        #Crear dataframe con 'Compras' y 'Ventas' en la columna 'tipooperacion'
        Compras = Consolidado[Consolidado['tipooperacion'] == 'Compras']
        Ventas = Consolidado[Consolidado['tipooperacion'] == 'Ventas']


        #Crear tablas pivote de 'Compras' y 'Ventas' por 'CLIENTE' y 'periodo'
        Compras_TD = Compras.pivot_table(index = ['CLIENTE' , 'periodo'] , values = ['montoneto' , 'NG' , 'Neto + NG' , 'montoiva' ] , aggfunc = 'sum')
        Ventas_TD = Ventas.pivot_table(index = ['CLIENTE' , 'periodo'] , values = ['montoneto' , 'NG' , 'Neto + NG'] , aggfunc = 'sum')

        #Crear columna de '% Computable' en base a la columna 'montoneto' / 'Neto + NG' de la tabla 'Ventas' y 'Ventas_Computable'
        Ventas_TD['% computable'] = Ventas_TD['montoneto'] / Ventas_TD['Neto + NG']

        #Separar el indice de las tablas pivote en columnas
        Compras_TD.reset_index(inplace=True)
        Ventas_TD.reset_index(inplace=True)

        #crear columas AUX en Compas_TD y Ventas_TD
        Compras_TD['Aux'] = Compras_TD['CLIENTE'] + ' - '+ Compras_TD['periodo'].astype(str)
        Ventas_TD['Aux'] = Ventas_TD['CLIENTE'] + ' - ' + Ventas_TD['periodo'].astype(str)

        #Merge de Compras_TD y Ventas_TD en base a la columna 'Aux' y traer solamente la columna '% computable' de Ventas_TD
        Compras_TD = pd.merge(  Compras_TD,
                                Ventas_TD[['Aux', '% computable']],
                                left_on='Aux' ,
                                right_on='Aux',
                                how='left')

        #Crear columna de diferencia entre '% computable' y '% Computable' y multiplicarlo por el 'montoiva'
        Compras_TD['CF computable'] = Compras_TD['% computable'] * Compras_TD['montoiva']

        #exportar las tablas pivote a excel en solapas distintas
        with pd.ExcelWriter( f'{Objetivo}Prorrateo mensual.xlsx') as writer:
            Compras_TD.to_excel(writer, sheet_name='Compras', index=False)
            Ventas_TD.to_excel(writer, sheet_name='Ventas' , index=False)

if __name__ == '__main__':
    Prorrateo_Mensual_XLS()