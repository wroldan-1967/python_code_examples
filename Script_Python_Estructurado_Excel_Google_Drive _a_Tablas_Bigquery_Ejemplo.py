#***************************************************************************************
# Proyecto	:  
# Fecha		: 15 / Sep / 2021
# Autor		: Wilson Roldan  
# Descripción   : Script Estructurado en Python para convertir Hojas dentro de Archivos Excel 
#                 Ubicados en Google Drive, a Tablas en Bigquery Tomando una Tabla previamente   
#                 Creada en Esquema ( Estructura de Campos).  
#                 Usando una cuenta de Servicio como autenticación.  
#****************************************************************************************
#                       MODIFICACIONES POSTERIORES                  *
#===================================================================*
#=  Fecha   | Persona |           Modificación Realizada           =*
#*===================================================================*
#*           |         |                                            
#*********************************************************************/

import os, google.auth 
from google.cloud import storage
from google.cloud import bigquery
from google.oauth2 import service_account

# Establecer la conexión al proyecto y dataset a través de una cuenta de servicio
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'CAMBIA ACA POR EL PATH+NOMBRE ARCHIVO JSON DE LA CUENTA DE SERVICIO p.e. D:\archvios\project-id-name-440a84e35d78.json'
credentials, your_project_id = google.auth.default(
    scopes=["https://www.googleapis.com/auth/cloud-platform"]
)
client = bigquery.Client(credentials=credentials, project=your_project_id,)

# Configuración para la creación de las tablas externas tomando como fuente Hojas de Archivos Excel en Google Drive 
external_config = bigquery.ExternalConfig("GOOGLE_SHEETS")
external_config.options.skip_leading_rows = 1  
external_config.options.write_disposition="WRITE_TRUNCATE",
external_config.max_bad_records = 0     
external_config.ignore_unknown_values = True  

# Generalmente las hojas del archivo pueden ser una lista Iterable , cada hoja se convertirá en una tabla externa, vista desde Bigquery
# Dos Archivos como fuente, con formato diferente pero mismo nombre de Hojas 
lista_hojas=['febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
uri_1 = ['https://docs.google.com/spreadsheets/d/1tCjm8VbEDCCAhV1emCROjg_kQucX-TIz7nsmxsaAgeY/edit?usp=sharing']
uri_2 = ['https://docs.google.com/spreadsheets/d/1Xdnn24o7-cy9OKhhH-ErsJOAlpd6NRJTj_fzdd1FVXA/edit?usp=sharing']
project = 'AJUSTA ACA EL NOMBRE DE TU PROYECTO '
dataset_id = 'AJUSTA ACA EL NOMBRE DEL DATASET DESTINO'
# Nombres de Tablas Modelo previamente Creadas en Estructura 
table_id_1 = 'Nombre_Tabla_Modelo_Gruopo_1'
table_id_2 = 'Nombre_Tabla_Modelo_Gruopo_1'

# Funcion encargada de Extraer/Generar el esquema de la Tabla Externa Destino con base en el mismo Esquema de las Hojas de Excel 
def extract_schema(project, dataset_id, table_source): 
    dataset_ref = client.dataset(dataset_id, project=project)
    table_ref = dataset_ref.table(table_source)
    table = client.get_table(table_ref)  
    lista_esquema = []
    result = ['{3}"{0}"{2}"{1}"{4}{5}'.format(schema.name,schema.field_type,',','bigquery.SchemaField(',')',',') for schema in table.schema]
    for x in result:
        lista_esquema.append(x.rstrip("'"))  
    return lista_esquema  

# Crea las Tablas Destino en el proyecto y Dataset Determinado 
# Con la estrutura fuente y los atos fuente. 
def cargarexcel(nom_tabla, nom_hoja, esquema, uri ): 
    table_id = nom_tabla
    schema = esquema 
    external_config.source_uris = uri 
    table = bigquery.Table(dataset.table(table_id), schema=schema)
    client.delete_table(table, not_found_ok=True)
    external_config.options.range = (nom_hoja)
    table.external_data_configuration = external_config
    table = client.create_table(table)

# Obtener el Esquema de Campos para la creación de las Tablas Externas. 
schema_1 = extract_schema(project,dataset_id,table_id_1)
schema_2 = extract_schema(project,dataset_id,table_id_2)

# Iterar Sobre los nombres de las Hojas de cada uno de los archivos a cargar 
# Para generar una tabla Externa en Bigquery con la misma Estructura/Esquema 
# que la hoja origen dentro de los archivos de Excel. 
# Se adioionan los Prefijos y Sufijos de los nombres de las tablas destino 
# Tambien se brinda el rango dentro de las hojas de excel a ser tenido en cuenta 
# como base de los datos a cargar. 
for item in lista_hojas:
    nom_tabla_1 = 'prefijo_grupo_uno_'+item+'sufijo_ext'
    nom_hoja_1 = item+'!A1:K501'
    nom_tabla_2 = 'prefijo_grupo_dos'+item+'sufijo_ext'
    nom_hoja_2 = item
    print(f"Cargando tabla {nom_tabla_1} Hoja {nom_hoja_1} " )
    cargarexcel(nom_tabla_1, nom_hoja_1, schema_1, uri_1)
    print(f"Cargando tabla {nom_tabla_2} Hoja {nom_hoja_2} " )
    cargarexcel(nom_tabla_2, nom_hoja_2, schema_2, uri_2)
