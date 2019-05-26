#!/usr/bin/env python
# coding: utf-8

# In[105]:


import pandas as pd
import numpy as np
import os
import openpyxl
import datetime
import parser


# In[106]:


# Files Directions, Rutes and file names
filepathCostos = r'\\u-srvnas\Areas\Mantenimiento\00.1 Mantto Planeamiento\0000. Gestión 2019\1. REPORTES\1. BASES DE DATOS DIARIOS\Data SAP\01. Costos'
filepathOtAv = r'\\u-srvnas\Areas\Mantenimiento\00.1 Mantto Planeamiento\0000. Gestión 2019\1. REPORTES\1. BASES DE DATOS DIARIOS\Data SAP\04. OT y Av'

#File Names
CeCo = r'BD-CeCo1004.xlsx' #Estatico
ClaseCostos = r'BD-ClaseCoste1004.xlsx' #Estatico
Budget = r'Presupuesto1004-2019Qlik.xlsx' #Estatico

filename3611 = r'3611.XLSX' #Dinamico
filename2993 = r'2993.XLSX' #Dinamico

#OT
filenameOt = r'IW39.xlsx'


# In[107]:


#Charging DataSet
fullpath_CeCo = os.path.join(filepathCostos,CeCo)
fullpath_ClCo = os.path.join(filepathCostos,ClaseCostos)
fullpath_Budget = os.path.join(filepathCostos,Budget)

fullpath_3611 = os.path.join(filepathCostos,filename3611)
fullpath_2993 = os.path.join(filepathCostos,filename2993)
fullpath_Ot = os.path.join(filepathOtAv,filenameOt)


# In[108]:


dbceco = pd.read_excel(fullpath_CeCo, converters={'Centro de Costo':str})
dbclco = pd.read_excel(fullpath_ClCo, converters={'Clase de Coste':str})
dbbudget = pd.read_excel(fullpath_Budget, converters={'CeCo':str, 'Clase de Coste':str, 'Mes Numero':str})
dbbudget = dbbudget.rename(columns={'Importe US$':'Costo','Mes Numero':'MesNumero'})


# In[109]:


db3611 = pd.read_excel(fullpath_3611, converters={'Orden partner':str,
                                                  'Centro de coste':str,
                                                  'Clase de coste':str,
                                                  'Documento compras':str,
                                                  'Proveedor':str,
                                                  'Material':str})
db3611 = db3611.dropna(subset=['Centro de coste'])
db3611 = db3611.rename(columns={'Período':'Periodo',
                                'Fe.contabilización':'FechaContable',
                                'Valor/Moneda objeto':'Costo',
                                'Moneda del informe':'Moneda',
                                'Orden partner':'Orden',
                                'Denominación del objeto del interlocutor':'DenomOrden',
                                'Documento compras':'DocumentoCompras',
                                'Texto de pedido':'DenomDocCompras',
                                'Nombre':'DenomProveedor',
                                'Centro de coste':'CeCo',
                                'Denominación del objeto':'DenomCeCo',
                                'Clase de coste':'ClCo',
                                'Descrip.clases coste':'DenomClCo',
                                'Texto breve de material':'DenomMaterial',
                                'Cantidad total reg.':'CantMaterial',
                                'Clase de documento':'ClaseDocumento',
                                'Indic.cargo/abono':'IndCargoAbono'})
db3611['Escenario'] = 'Real'
db3611 = db3611.dropna(subset=['CeCo'])


# In[110]:


db2993 = pd.read_excel(fullpath_2993, converters={'Orden':str,
                                                  'Clase de coste':str})
db2993 = db2993.dropna(subset=['Orden'])

db2993 = db2993.rename(columns={'División':'Division',
                                'Período':'Periodo',
                                'Fe.contabilización':'FechaContable',
                                'Denominación del objeto':'DenomOrden',
                                'Valor/Moneda objeto':'Costo',
                                'Moneda del objeto':'Moneda',
                                'Texto breve de material':'DenomMaterial',
                                'Cantidad total reg.':'CantMaterial',
                                'Ud. cantidad contab.':'UnidadCantMaterial',
                                'Documento compras':'DocumentoCompras',
                                'Texto de pedido':'DenomDocCompras',
                                'Clase de coste':'ClCo',
                                'Descrip.clases coste':'DenomClCo',
                                'Nombre':'DenomProveedor',
                                'Clase de documento':'ClaseDocumento',
                                'Indic.cargo/abono':'IndCargoAbono'})

db2993['Escenario'] = 'Real'


# In[111]:


dbot = pd.read_excel(fullpath_Ot, head = 0,converters={'Orden':str,'Aviso':str,'Centro coste':str})
dbot = dbot.rename(columns={'Texto breve':'DenomOrden',
                                'Clase de orden':'ClaseOrden',
                                'Cl.actividad PM':'ClaseActividadOrden',
                                'Status sistema':'EstatusSistema',
                                'Status usuario':'EstatusUsuario',
                                'Indicador ABC':'CriticidadEq',
                                'Fe.inic.extrema':'FechaProgramada',
                                'Hora inicio':'HoraProgramada',
                                'Fe.inicio real':'FechaInicioEjecucion',
                                'Hora in.real':'HoraInicioEjecucion',
                                'Fecha fin real':'FechaFinEjecucion',
                                'Fin real (hora)':'HoraFinEjecucion',
                                'Liberación real':'FechaDeProgramacon',
                                'Denominación':'DenomEq',
                                'Ubicac.técnica':'UbicacionTecnica',
                                'Denominación.1':'DenmUbT',
                                'Plan mant.prev.':'PlanMtto',
                                'Revisión':'Revision',
                                'Centro coste':'CeCo',
                                'Grupo planif.':'GrupoPlanif',
                                'Pto.tbjo.resp.':'PuestoTrabajoResp',
                                'Puesto trabajo':'PuestoTrabajo'})


# In[113]:


#Reporte de Costos / 2993 <-> Clase de Coste
clco2993 = pd.merge(db2993, 
                    dbclco[['Clase de Coste', 'Grupo Naturaleza']], 
                    left_on='ClCo', right_on='Clase de Coste', 
                    how = 'left',
                    validate = 'm:1')
clco2993 = clco2993.drop(['Clase de Coste'], axis = 1)


# In[114]:


#Reporte de Costos / 2993,clase de coste<-> Orden de Mantenimiento
clco2993ot = pd.merge(clco2993, 
                    dbot[['Orden',
                          'ClaseOrden',
                          'CeCo',
                          'CriticidadEq',
                          'Equipo','DenomEq',
                          'UbicacionTecnica', 
                          'DenmUbT',
                          'PlanMtto',
                          'GrupoPlanif']], 
                    on = 'Orden', 
                    how = 'left',
                    validate = 'm:1')


# In[115]:


#Reporte de Costos / 2993,clase de coste,Orden de Mantenimiento<-> Centro de Coste
clco2993otceco = pd.merge(clco2993ot, 
                    dbceco[['Centro de Costo','Área Responsable','Proceso']],
                    left_on='CeCo', 
                    right_on='Centro de Costo',
                    how = 'left',
                    validate = 'm:1')
clco2993otceco = clco2993otceco.drop(['Centro de Costo'], axis = 1)
clco2993otceco["MesNumero"] = clco2993otceco["FechaContable"].apply(lambda x: x.strftime('%m'))


# In[116]:


dbbudgetceco = pd.merge(dbbudget,
                   dbceco[['Centro de Costo','Área Responsable', 'Proceso']],
                   left_on='CeCo', right_on='Centro de Costo',
                    how = 'left',)
dbbudgetceco = dbbudgetceco.drop(['Centro de Costo'], axis = 1)


# In[117]:


#Reporte de Costos / 3611 <-> Clase de Coste
clco3611 = pd.merge(db3611, 
                    dbclco[['Clase de Coste', 
                            'Grupo Naturaleza']], 
                    left_on='ClCo', 
                    right_on='Clase de Coste', 
                    how = 'left',
                    validate = 'm:1')
clco3611 = clco3611.drop(['Clase de Coste'], axis = 1)


# In[118]:


#Reporte de Costos / 2993,clase de coste<-> Orden de Mantenimiento
clco3611ot = pd.merge(clco3611, 
                    dbot[['Orden',
                          'ClaseOrden',
                          'CriticidadEq',
                          'Equipo','DenomEq',
                          'UbicacionTecnica', 
                          'DenmUbT',
                          'PlanMtto',
                          'GrupoPlanif']], 
                    on='Orden', 
                    how = 'left',
                    validate = 'm:1')


# In[119]:


#Reporte de Costos / 3611,clase de coste,Orden de Mantenimiento<-> Centro de Coste
clco3611otceco = pd.merge(clco3611ot, 
                    dbceco[['Centro de Costo','Área Responsable','Proceso']],
                    left_on='CeCo', 
                    right_on='Centro de Costo',
                    how = 'left',
                    validate = 'm:1')
clco3611otceco = clco3611otceco.drop(['Centro de Costo'], axis = 1)
clco3611otceco["MesNumero"] = clco3611otceco["FechaContable"].apply(lambda x: x.strftime('%m'))


# In[120]:


RC02 = pd.concat([clco2993otceco[['FechaContable', 
                                  'Orden',
                                  'DenomOrden', 
                                  'Costo', 
                                  'Moneda', 
                                  'ClaseDocumento', 
                                  'IndCargoAbono', 
                                  'Escenario', 
                                  'Grupo Naturaleza',
                                  'CeCo',
                                  'Área Responsable', 
                                  'Proceso', 
                                  'MesNumero']], 
                  dbbudgetceco[['CeCo', 
                            'MesNumero',
                            'Costo',
                            'Proceso',
                            'Área Responsable',
                            'Escenario',
                            'Grupo Naturaleza']]], 
                 ignore_index=True, 
                 sort=True)


# In[121]:


with pd.ExcelWriter('output.xlsx', mode='w') as writer:
    RC02.to_excel(writer, sheet_name='Sheetname2',index = False)


# In[ ]:




