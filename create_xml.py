import math

import pandas as pd
from lxml import etree as et

raw_data = pd.read_excel("ATS SEPTIEMBRE.xlsx")

Column_heading_ = [None] * len(raw_data)
root = et.Element('iva')
Column_heading_[1] = et.SubElement(root, 'TipoIDInformante')
Column_heading_[2] = et.SubElement(root, 'IdInformante')
Column_heading_[3] = et.SubElement(root, 'razonSocial')
Column_heading_[4] = et.SubElement(root, 'Anio')
Column_heading_[5] = et.SubElement(root, 'Mes')
Column_heading_[6] = et.SubElement(root, 'numEstabRuc')
Column_heading_[7] = et.SubElement(root, 'totalVentas')
Column_heading_[8] = et.SubElement(root, 'codigoOperativo')
root_tag_compras = et.SubElement(root, 'compras')

raw_data['estabRetencion1'] = raw_data['estabRetencion1'].fillna('NA')
raw_data['estabRetencion1'] = raw_data['estabRetencion1'].astype(str)

raw_data['secRetencion1'] = raw_data['secRetencion1'].fillna('NA')
raw_data['secRetencion1'] = raw_data['secRetencion1'].astype(str)

raw_data['ptoEmiRetencion1'] = raw_data['ptoEmiRetencion1'].fillna('NA')
raw_data['ptoEmiRetencion1'] = raw_data['ptoEmiRetencion1'].astype(str)

raw_data['secRetencion1'] = raw_data['secRetencion1'].astype(str).str.replace('.0', '')
raw_data['formaPago'] = raw_data['formaPago'].astype(str).str.replace('.0', '')

raw_data['paisEfecPago'] = raw_data['paisEfecPago'].fillna('NA')
raw_data['aplicConvDobTrib'] = raw_data['aplicConvDobTrib'].fillna('NA')
raw_data['pagExtSujRetNorLeg'] = raw_data['pagExtSujRetNorLeg'].fillna('NA')

raw_data['autRetencion1'] = raw_data['autRetencion1'].fillna('NA')
raw_data['fechaEmiRet1'] = raw_data['fechaEmiRet1'].fillna('NA')
# raw_data['formaPago'] = raw_data['formaPago'].fillna('NA')


# Obtener los registros duplicados en todas las columnas excepto en [codRetAir, baseImpAir, porcentajeAir, valRetAir]

# Obtener los registros duplicados en todas las columnas excepto en [codRetAir, baseImpAir, porcentajeAir, valRetAir]
columnas_excluidas = ['codRetAir', 'baseImpAir', 'porcentajeAir', 'valRetAir']
columnas_a_verificar = raw_data.columns.difference(columnas_excluidas)
duplicados = raw_data[raw_data.duplicated(subset=columnas_a_verificar, keep=False)]

# duplicados['index'] = duplicados.index
duplicados.loc[:, 'index'] = duplicados.index
# duplicados_count = duplicados.groupby('idProv','codSustento','tipoComprobante','establecimiento','puntoEmision','secuencial','autorizacion').size().reset_index(name='count')
duplicados_count = duplicados.groupby(['idProv', 'codSustento', 'tipoComprobante', 'establecimiento', 'puntoEmision', 'secuencial', 'autorizacion']).size().reset_index(name='count')
# duplicados_result = duplicados_count.merge(duplicados['idProv', 'codSustento', 'tipoComprobante', 'establecimiento', 'puntoEmision', 'secuencial', 'autorizacion', 'index'], on='idProv')
duplicados_result = duplicados_count.merge(duplicados[['idProv', 'codSustento', 'tipoComprobante', 'establecimiento', 'puntoEmision', 'secuencial', 'autorizacion', 'index']], on=['idProv', 'codSustento', 'tipoComprobante', 'establecimiento', 'puntoEmision', 'secuencial', 'autorizacion'])
i = 0
new_index = None

for i, row in raw_data.iterrows():
    if new_index == i:
        continue
    # root_tag_iva = et.SubElement(root, 'iva')  # ==> Root Name
    # These are the tag names for each row
    new_index = None



    root_tags_detalle_compras = et.SubElement(root_tag_compras, 'detalleCompras')
    Column_heading_[9] = et.SubElement(root_tags_detalle_compras, 'codSustento')
    Column_heading_[10] = et.SubElement(root_tags_detalle_compras, 'tpIdProv')
    Column_heading_[11] = et.SubElement(root_tags_detalle_compras, 'idProv')
    Column_heading_[12] = et.SubElement(root_tags_detalle_compras, 'tipoComprobante')
    Column_heading_[13] = et.SubElement(root_tags_detalle_compras, 'parteRel')
    Column_heading_[14] = et.SubElement(root_tags_detalle_compras, 'fechaRegistro')
    Column_heading_[15] = et.SubElement(root_tags_detalle_compras, 'establecimiento')
    Column_heading_[16] = et.SubElement(root_tags_detalle_compras, 'puntoEmision')
    Column_heading_[17] = et.SubElement(root_tags_detalle_compras, 'secuencial')
    Column_heading_[18] = et.SubElement(root_tags_detalle_compras, 'fechaEmision')
    Column_heading_[19] = et.SubElement(root_tags_detalle_compras, 'autorizacion')
    Column_heading_[20] = et.SubElement(root_tags_detalle_compras, 'baseNoGraIva')
    Column_heading_[21] = et.SubElement(root_tags_detalle_compras, 'baseImponible')
    Column_heading_[22] = et.SubElement(root_tags_detalle_compras, 'baseImpGrav')
    Column_heading_[23] = et.SubElement(root_tags_detalle_compras, 'baseImpExe')
    Column_heading_[24] = et.SubElement(root_tags_detalle_compras, 'montoIce')
    Column_heading_[25] = et.SubElement(root_tags_detalle_compras, 'montoIva')
    Column_heading_[26] = et.SubElement(root_tags_detalle_compras, 'valRetBien10')
    Column_heading_[27] = et.SubElement(root_tags_detalle_compras, 'valRetServ20')
    Column_heading_[28] = et.SubElement(root_tags_detalle_compras, 'valorRetBienes')
    Column_heading_[29] = et.SubElement(root_tags_detalle_compras, 'valRetServ50')
    Column_heading_[30] = et.SubElement(root_tags_detalle_compras, 'valorRetServicios')
    Column_heading_[31] = et.SubElement(root_tags_detalle_compras, 'valRetServ100')
    Column_heading_[32] = et.SubElement(root_tags_detalle_compras, 'totbasesImpReemb')
    root_tag_pagoExterior = et.SubElement(root_tags_detalle_compras, 'pagoExterior')
    Column_heading_[33] = et.SubElement(root_tag_pagoExterior, 'pagoLocExt')
    Column_heading_[34] = et.SubElement(root_tag_pagoExterior, 'paisEfecPago')
    Column_heading_[35] = et.SubElement(root_tag_pagoExterior, 'aplicConvDobTrib')
    Column_heading_[36] = et.SubElement(root_tag_pagoExterior, 'pagExtSujRetNorLeg')
    valor = float(row['formaPago'])
    # Aqui pregunto si la columno formaPago tiene un valor ya que si lo tiene debo de comenzar a agregar un tag nuevo 
    # 1. Se obtiene el valor de la columna 'formaPago' y se convierte a un número decimal utilizando la función float(). 
    # 2. Se verifica si el valor es un NaN (Not a Number) utilizando la función math.isnan(). Si el valor es NaN, significa que no es un número y se ejecuta el bloque de código dentro del if. 
    
    if math.isnan(valor):
        root_tag_air = et.SubElement(root_tags_detalle_compras, 'air')

        if i in duplicados_result['index'].values:
            numero_veces = duplicados_result.loc[duplicados_result['index'] == i, 'count'].values[0]
            idProveedor = duplicados_result.loc[duplicados_result['index'] == i, 'idProv'].values[0]
            indice = duplicados_result.loc[duplicados_result['index'] == i, 'index'].values[0]
            numero = 36

            for cantidad in range(numero_veces):
                root_tag_detail_air = et.SubElement(root_tag_air, 'detalleAir')
                Column_heading_[numero + 1] = et.SubElement(root_tag_detail_air, 'codRetAir')
                Column_heading_[numero + 2] = et.SubElement(root_tag_detail_air, 'baseImpAir')
                Column_heading_[numero + 3] = et.SubElement(root_tag_detail_air, 'porcentajeAir')
                Column_heading_[numero + 4] = et.SubElement(root_tag_detail_air, 'valRetAir')
                numero = numero + 4

            Column_heading_[numero + 1] = et.SubElement(root_tags_detalle_compras, 'estabRetencion1')
            Column_heading_[numero + 2] = et.SubElement(root_tags_detalle_compras, 'ptoEmiRetencion1')
            Column_heading_[numero + 3] = et.SubElement(root_tags_detalle_compras, 'secRetencion1')
            Column_heading_[numero + 4] = et.SubElement(root_tags_detalle_compras, 'autRetencion1')
            Column_heading_[numero + 5] = et.SubElement(root_tags_detalle_compras, 'fechaEmiRet1')

        else:
            root_tag_detail_air = et.SubElement(root_tag_air, 'detalleAir')
            Column_heading_[37] = et.SubElement(root_tag_detail_air, 'codRetAir')
            Column_heading_[38] = et.SubElement(root_tag_detail_air, 'baseImpAir')
            Column_heading_[39] = et.SubElement(root_tag_detail_air, 'porcentajeAir')
            Column_heading_[40] = et.SubElement(root_tag_detail_air, 'valRetAir')

            Column_heading_[41] = et.SubElement(root_tags_detalle_compras, 'estabRetencion1')
            Column_heading_[42] = et.SubElement(root_tags_detalle_compras, 'ptoEmiRetencion1')
            Column_heading_[43] = et.SubElement(root_tags_detalle_compras, 'secRetencion1')
            Column_heading_[44] = et.SubElement(root_tags_detalle_compras, 'autRetencion1')
            Column_heading_[45] = et.SubElement(root_tags_detalle_compras, 'fechaEmiRet1')
    
    # 3. Si el valor no es NaN, significa que es un número y no se ejecuta el bloque de código dentro del if.
    # Es decir debo de agregar un tag llamado formasPago
    else:
        root_tag_formas_pago = et.SubElement(root_tags_detalle_compras, 'formasDePago')
        Column_heading_[37] = et.SubElement(root_tag_formas_pago, 'formaPago')
        root_tag_air = et.SubElement(root_tags_detalle_compras, 'air')

        if i in duplicados_result['index'].values:
            numero_veces = duplicados_result.loc[duplicados_result['index'] == i, 'count'].values[0]
            idProveedor = duplicados_result.loc[duplicados_result['index'] == i, 'idProv'].values[0]
            indice = duplicados_result.loc[duplicados_result['index'] == i, 'index'].values[0]
            numero = 37

            for cantidad in range(numero_veces):
                root_tag_detail_air = et.SubElement(root_tag_air, 'detalleAir')
                Column_heading_[numero + 1] = et.SubElement(root_tag_detail_air, 'codRetAir')
                Column_heading_[numero + 2] = et.SubElement(root_tag_detail_air, 'baseImpAir')
                Column_heading_[numero + 3] = et.SubElement(root_tag_detail_air, 'porcentajeAir')
                Column_heading_[numero + 4] = et.SubElement(root_tag_detail_air, 'valRetAir')
                numero = numero + 4

            Column_heading_[numero + 1] = et.SubElement(root_tags_detalle_compras, 'estabRetencion1')
            Column_heading_[numero + 2] = et.SubElement(root_tags_detalle_compras, 'ptoEmiRetencion1')
            Column_heading_[numero + 3] = et.SubElement(root_tags_detalle_compras, 'secRetencion1')
            Column_heading_[numero + 4] = et.SubElement(root_tags_detalle_compras, 'autRetencion1')
            Column_heading_[numero + 5] = et.SubElement(root_tags_detalle_compras, 'fechaEmiRet1')
        else:
            root_tag_detail_air = et.SubElement(root_tag_air, 'detalleAir')
            Column_heading_[38] = et.SubElement(root_tag_detail_air, 'codRetAir')
            Column_heading_[39] = et.SubElement(root_tag_detail_air, 'baseImpAir')
            Column_heading_[40] = et.SubElement(root_tag_detail_air, 'porcentajeAir')
            Column_heading_[41] = et.SubElement(root_tag_detail_air, 'valRetAir')

            Column_heading_[42] = et.SubElement(root_tags_detalle_compras, 'estabRetencion1')
            Column_heading_[43] = et.SubElement(root_tags_detalle_compras, 'ptoEmiRetencion1')
            Column_heading_[44] = et.SubElement(root_tags_detalle_compras, 'secRetencion1')
            Column_heading_[45] = et.SubElement(root_tags_detalle_compras, 'autRetencion1')
            Column_heading_[46] = et.SubElement(root_tags_detalle_compras, 'fechaEmiRet1')


    # Aqui ya comienzo a ingresar informacion en el esqueleto del xml
    Column_heading_[1].text = str(row['TipoIDInformante'])
    Column_heading_[2].text = '0' + str(row['IdInformante']) if len(str(row['IdInformante'])) <= 12 else str(
        row['IdInformante'])
    Column_heading_[3].text = str(row['razonSocial'])
    Column_heading_[4].text = str(row['Anio'])
    Column_heading_[5].text = str(row['Mes']).zfill(2)
    Column_heading_[6].text = str(row['numEstabRuc']).zfill(3)
    Column_heading_[7].text = str(row['totalVentas'])
    Column_heading_[8].text = str(row['codigoOperativo'])
    Column_heading_[9].text = str(row['codSustento']).zfill(2)
    Column_heading_[10].text = str(row['tpIdProv']).zfill(2)
    Column_heading_[11].text = '0' + str(row['idProv']) if len(str(row['idProv'])) <= 12 else str(row['idProv'])
    Column_heading_[12].text = str(row['tipoComprobante']).zfill(2)
    Column_heading_[13].text = str(row['parteRel'])
    Column_heading_[14].text = str(row['fechaRegistro'])
    Column_heading_[15].text = str(row['establecimiento']).zfill(3)
    Column_heading_[16].text = str(row['puntoEmision']).zfill(3)
    Column_heading_[17].text = str(row['secuencial'])
    Column_heading_[18].text = str(row['fechaEmision'])
    Column_heading_[19].text = str(row['autorizacion'])
    Column_heading_[20].text = str(row['baseNoGraIva'])
    Column_heading_[21].text = "{:.2f}".format(float(str(row['baseImponible'])))
    Column_heading_[22].text = "{:.2f}".format(float(str(row['baseImpGrav'])))
    Column_heading_[23].text = "{:.2f}".format(float(str(row['baseImpExe'])))
    Column_heading_[24].text = "{:.2f}".format(float(str(row['montoIce'])))
    Column_heading_[25].text = "{:.2f}".format(float(str(row['montoIva'])))
    Column_heading_[26].text = "{:.2f}".format(float(str(row['valRetBien10'])))
    Column_heading_[27].text = "{:.2f}".format(float(str(row['valRetServ20'])))
    Column_heading_[28].text = "{:.2f}".format(float(str(row['valorRetBienes'])))
    Column_heading_[29].text = "{:.2f}".format(float(str(row['valRetServ50'])))
    Column_heading_[30].text = "{:.2f}".format(float(str(row['valorRetServicios'])))
    Column_heading_[31].text = "{:.2f}".format(float(str(row['valRetServ100'])))
    Column_heading_[32].text = "{:.2f}".format(float(str(row['totbasesImpReemb'])))
    Column_heading_[33].text = str(row['pagoLocExt']).zfill(2)
    Column_heading_[34].text = str(row['paisEfecPago'])
    Column_heading_[35].text = str(row['aplicConvDobTrib'])
    Column_heading_[36].text = str(row['pagExtSujRetNorLeg'])
    valor = float(row['formaPago'])
     # Aqui pregunto si la columno formaPago tiene un valor ya que si lo tiene debo de comenzar a agregar un tag nuevo 
    # 1. Se obtiene el valor de la columna 'formaPago' y se convierte a un número decimal utilizando la función float(). 
    # 2. Se verifica si el valor es un NaN (Not a Number) utilizando la función math.isnan(). Si el valor es NaN, significa que no es un número y se ejecuta el bloque de código dentro del if. 
    if math.isnan(valor):

        if i in duplicados_result['index'].values:
            numero_veces = duplicados_result.loc[duplicados_result['index'] == i, 'count'].values[0]
            idProveedor = duplicados_result.loc[duplicados_result['index'] == i, 'idProv'].values[0]
            indice = duplicados_result.loc[duplicados_result['index'] == i, 'index'].values[0]
            numero = 36


            filtered_data = duplicados_result[duplicados_result['idProv'] == idProveedor]

            for index in filtered_data['index']:
                if index >= i and abs(index - i) <= 1:
                    Column_heading_[numero + 1].text = str(raw_data.loc[index, 'codRetAir'])
                    Column_heading_[numero + 2].text = "{:.2f}".format(float(str(raw_data.loc[index, 'baseImpAir'])))
                    Column_heading_[numero + 3].text = "{:.2f}".format(float(str(raw_data.loc[index, 'porcentajeAir'])))
                    Column_heading_[numero + 4].text = "{:.2f}".format(float(str(raw_data.loc[index, 'valRetAir'])))
                    numero = numero + 4
                    new_index = index

            if row['estabRetencion1'] != 'NA':
                Column_heading_[numero + 1].text = str(row['estabRetencion1']).replace('.0', '').zfill(3)
            else:
                Column_heading_[numero + 1].text = str(row['estabRetencion1'])

            if row['ptoEmiRetencion1'] != 'NA':
                Column_heading_[numero + 2].text = str(row['ptoEmiRetencion1']).replace('.0', '').zfill(3)
            else:
                Column_heading_[numero + 2].text = str(row['ptoEmiRetencion1'])




            Column_heading_[numero + 3].text = str(row['secRetencion1'])

            Column_heading_[numero + 4].text = str(row['autRetencion1'])
            Column_heading_[numero + 5].text = str(row['fechaEmiRet1'])

        else:
            Column_heading_[37].text = str(row['codRetAir'])
            Column_heading_[38].text = "{:.2f}".format(float(str(row['baseImpAir'])))
            Column_heading_[39].text = "{:.2f}".format(float(str(row['porcentajeAir'])))
            Column_heading_[40].text = "{:.2f}".format(float(str(row['valRetAir'])))

            if row['estabRetencion1'] != 'NA':
                Column_heading_[41].text = str(row['estabRetencion1']).replace('.0', '').zfill(3)
            else:
                Column_heading_[41].text = str(row['estabRetencion1'])

            if row['ptoEmiRetencion1'] != 'NA':
                Column_heading_[42].text = str(row['ptoEmiRetencion1']).replace('.0', '').zfill(3)
            else:
                Column_heading_[42].text = str(row['ptoEmiRetencion1'])

            Column_heading_[43].text = str(row['secRetencion1'])
            Column_heading_[44].text = str(row['autRetencion1'])
            Column_heading_[45].text = str(row['fechaEmiRet1'])


    # 3. Si el valor no es NaN, significa que es un número y no se ejecuta el bloque de código dentro del if.
    # Es decir debo de agregar un tag llamado formasPago
    else:
        Column_heading_[37].text = str(row['formaPago'])
        if i in duplicados_result['index'].values:
            numero_veces = duplicados_result.loc[duplicados_result['index'] == i, 'count'].values[0]
            idProveedor = duplicados_result.loc[duplicados_result['index'] == i, 'idProv'].values[0]
            indice = duplicados_result.loc[duplicados_result['index'] == i, 'index'].values[0]
            numero = 37

            filtered_data = duplicados_result[duplicados_result['idProv'] == idProveedor]
            for index in filtered_data['index']:

                if index>=i and abs(index - i) <= 1:
                    Column_heading_[numero + 1].text = str(raw_data.loc[index, 'codRetAir'])
                    Column_heading_[numero + 2].text = "{:.2f}".format(float(str(raw_data.loc[index, 'baseImpAir'])))
                    Column_heading_[numero + 3].text = "{:.2f}".format(float(str(raw_data.loc[index, 'porcentajeAir'])))
                    Column_heading_[numero + 4].text = "{:.2f}".format(float(str(raw_data.loc[index, 'valRetAir'])))
                    numero = numero + 4
                    new_index = index


            if row['estabRetencion1'] != 'NA':
                Column_heading_[numero + 1].text = str(row['estabRetencion1']).replace('.0', '').zfill(3)
            else:
                Column_heading_[numero + 1].text = str(row['estabRetencion1'])

            if row['ptoEmiRetencion1'] != 'NA':
                Column_heading_[numero + 2].text = str(row['ptoEmiRetencion1']).replace('.0', '').zfill(3)
            else:
                Column_heading_[numero + 2].text = str(row['ptoEmiRetencion1'])

            Column_heading_[numero + 3].text = str(row['secRetencion1'])
            Column_heading_[numero + 4].text = str(row['autRetencion1'])
            Column_heading_[numero + 5].text = str(row['fechaEmiRet1'])


        else:
            Column_heading_[38].text = str(row['codRetAir'])
            Column_heading_[39].text = "{:.2f}".format(float(str(row['baseImpAir'])))
            Column_heading_[40].text = "{:.2f}".format(float(str(row['porcentajeAir'])))
            Column_heading_[41].text = "{:.2f}".format(float(str(row['valRetAir'])))

            if row['estabRetencion1'] != 'NA':
                Column_heading_[42].text = str(row['estabRetencion1']).replace('.0', '').zfill(3)

            else:
                Column_heading_[42].text = str(row['estabRetencion1'])

            if row['ptoEmiRetencion1'] != 'NA':
                Column_heading_[43].text = str(row['ptoEmiRetencion1']).replace('.0', '').zfill(3)
            else:
                Column_heading_[43].text = str(row['ptoEmiRetencion1'])

            Column_heading_[44].text = str(row['secRetencion1'])
            Column_heading_[45].text = str(row['autRetencion1'])
            Column_heading_[46].text = str(row['fechaEmiRet1'])

tree = et.ElementTree(root)
et.indent(tree, space="\t", level=0)
tree.write('ATS SEPTIEMBRE.xml', encoding="utf-8")
