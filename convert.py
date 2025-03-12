import pandas as pd
from lxml import etree
import argparse
import os
import zipfile
import shutil

def parse_arguments():
    parser = argparse.ArgumentParser(description='Convert Excel to XML grouped by a specified column.')
    parser.add_argument('column', type=str, nargs='?', help='The column name or number to group by (optional)')
    parser.add_argument('--max_records', type=int, help='The maximum number of records per XML file (optional)')
    parser.add_argument('--by_number', action='store_true', help='Indicate if the column is specified by number')
    return parser.parse_args()

def setup_directories(script_dir):
    output_dir = os.path.join(script_dir, 'xml')
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    return output_dir

def read_excel(script_dir):
    excel_path = os.path.join(script_dir, 'datos.xlsx')
    return pd.read_excel(excel_path), os.path.splitext(os.path.basename(excel_path))[0]

def get_column_name(df, column, by_number):
    if by_number:
        column_index = int(column) - 1  # Ajustar para que el índice empiece en 1
        return df.columns[column_index]
    return column

def get_xsd_columns(script_dir):
    xsd_path = os.path.join(script_dir, 'schema.xsd')
    xsd_doc = etree.parse(xsd_path)
    xsd_root = xsd_doc.getroot()
    ns = {'xs': 'http://www.w3.org/2001/XMLSchema'}
    elements = xsd_root.findall('.//xs:element[@name="Item"]/xs:complexType/xs:sequence/xs:element', ns)
    return [element.get('name') for element in elements]

def create_root_element():
    return etree.Element('Root')

def add_meta_to_xml(root, meta_data):
    meta = etree.SubElement(root, 'meta')
    for key, value in meta_data.items():
        child = etree.SubElement(meta, key)
        child.text = str(value)

def add_elements_to_xml(root, df):
    for _, row in df.iterrows():
        item = etree.SubElement(root, 'Item')
        for col in df.columns:
            child = etree.SubElement(item, col)
            child.text = str(row[col])

def save_xml_to_file(root, filename):
    xml_str = etree.tostring(root, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    with open(filename, 'wb') as f:
        f.write(xml_str)
    return xml_str

def create_and_save_xml(df, output_dir, base_name, meta_data, max_records=None):
    xml_files = []
    subgroups = [df[i:i + max_records] for i in range(0, len(df), max_records)] if max_records else [df]
    for idx, subgroup in enumerate(subgroups):
        root = create_root_element()
        add_meta_to_xml(root, meta_data)
        add_elements_to_xml(root, subgroup)
        suffix = f'_{idx}' if max_records and idx > 0 else ''
        filename = os.path.join(output_dir, f'{base_name}{suffix}.xml')
        xml_str = save_xml_to_file(root, filename)
        xml_files.append((xml_str, base_name, suffix, filename))
    return xml_files

def create_xml_files(df, column, max_records, output_dir, script_dir, base_name):
    xsd_columns = get_xsd_columns(script_dir)
    df = df[xsd_columns]  # Reordenar las columnas del DataFrame según el XSD
    meta_data = {
        'created': df['created'].iloc[0] if 'created' in df.columns else '',
        'author': df['author'].iloc[0] if 'author' in df.columns else '',
        'version': df['version'].iloc[0] if 'version' in df.columns else ''
    }
    if column:
        grouped = df.groupby(column)
        xml_files = []
        for group_name, group in grouped:
            xml_files.extend(create_and_save_xml(group, output_dir, group_name, meta_data, max_records))
    else:
        xml_files = create_and_save_xml(df, output_dir, base_name, meta_data, max_records)
    
    for xml_str, group_name, suffix, filename in xml_files:
        validate_xml(xml_str, group_name, suffix, script_dir)
    
    return [filename for _, _, _, filename in xml_files]

def validate_xml(xml_str, group_name, suffix, script_dir):
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

def create_zip_file(xml_files, script_dir):
    zip_filename = os.path.join(script_dir, 'xml_files.zip')
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for xml_file in xml_files:
            zipf.write(xml_file, os.path.basename(xml_file))
    print(f"Todos los archivos XML se han guardado en {os.path.dirname(xml_files[0])} y se han comprimido en {zip_filename}.")

def main():
    args = parse_arguments()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = setup_directories(script_dir)
    df, base_name = read_excel(script_dir)
    column_name = get_column_name(df, args.column, args.by_number) if args.column else None
    xml_files = create_xml_files(df, column_name, args.max_records, output_dir, script_dir, base_name)
    create_zip_file(xml_files, script_dir)

if __name__ == '__main__':
    main()