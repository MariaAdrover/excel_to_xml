import pandas as pd
from lxml import etree
import argparse
import os
import zipfile
import shutil

# Configurar argparse para manejar los argumentos de la línea de comandos
parser = argparse.ArgumentParser(description='Convert Excel to XML grouped by a specified column.')
parser.add_argument('column', type=str, help='The column name to group by')
parser.add_argument('--max_records', type=int, help='The maximum number of records per XML file (optional)')
args = parser.parse_args()

# Obtener la ruta del directorio del script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Definir la ruta de la carpeta 'xml'
output_dir = os.path.join(script_dir, 'xml')

# Borrar la carpeta 'xml' si ya existe
if os.path.exists(output_dir):
    shutil.rmtree(output_dir)

# Crear la carpeta 'xml'
os.makedirs(output_dir, exist_ok=True)

# Leer el archivo Excel
excel_path = os.path.join(script_dir, 'datos.xlsx')
df = pd.read_excel(excel_path)

# Agrupar el DataFrame por la columna especificada
grouped = df.groupby(args.column)

# Crear un archivo XML para cada grupo
xml_files = []
for group_name, group in grouped:
    if args.max_records:
        # Dividir el grupo en subgrupos si supera el máximo de registros
        subgroups = [group[i:i + args.max_records] for i in range(0, len(group), args.max_records)]
    else:
        subgroups = [group]
    
    for idx, subgroup in enumerate(subgroups):
        # Crear el elemento raíz del XML
        root = etree.Element('Root')
        
        # Convertir cada fila del subgrupo en un elemento XML
        for _, row in subgroup.iterrows():
            item = etree.SubElement(root, 'Item')
            for col in subgroup.columns:
                child = etree.SubElement(item, col)
                child.text = str(row[col])
        
        # Convertir el árbol XML a una cadena
        xml_str = etree.tostring(root, pretty_print=True, xml_declaration=True, encoding='UTF-8')
        
        # Guardar el XML en un archivo con el nombre del grupo y un sufijo si es necesario
        suffix = f'_{idx}' if args.max_records and idx > 0 else ''
        filename = os.path.join(output_dir, f'{group_name}{suffix}.xml')
        with open(filename, 'wb') as f:
            f.write(xml_str)
        xml_files.append(filename)

        # Validar el XML contra el esquema XSD
        xsd_path = os.path.join(script_dir, 'schema.xsd')
        xsd_doc = etree.parse(xsd_path)
        xsd = etree.XMLSchema(xsd_doc)

        xml_doc = etree.fromstring(xml_str)
        if xsd.validate(xml_doc):
            print(f"El XML para {group_name}{suffix} es válido.")
        else:
            print(f"El XML para {group_name}{suffix} no es válido.")
            for error in xsd.error_log:
                print(error.message)

# Crear un archivo ZIP con todos los archivos XML
zip_filename = os.path.join(script_dir, 'xml_files.zip')
with zipfile.ZipFile(zip_filename, 'w') as zipf:
    for xml_file in xml_files:
        zipf.write(xml_file, os.path.basename(xml_file))

print(f"Todos los archivos XML se han guardado en {output_dir} y se han comprimido en {zip_filename}.")