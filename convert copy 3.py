import pandas as pd
from lxml import etree
import argparse
import os
import zipfile
import shutil

def parse_arguments():
    parser = argparse.ArgumentParser(description='Convert Excel to XML grouped by a specified column.')
    parser.add_argument('column', type=str, help='The column name or number to group by')
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
    return pd.read_excel(excel_path)

def get_column_name(df, column, by_number):
    if by_number:
        column_index = int(column) - 1  # Ajustar para que el índice empiece en 1
        return df.columns[column_index]
    return column

def create_xml_files(df, column, max_records, output_dir, script_dir):
    grouped = df.groupby(column)
    xml_files = []
    for group_name, group in grouped:
        subgroups = [group[i:i + max_records] for i in range(0, len(group), max_records)] if max_records else [group]
        for idx, subgroup in enumerate(subgroups):
            root = etree.Element('Root')
            for _, row in subgroup.iterrows():
                item = etree.SubElement(root, 'Item')
                for col in subgroup.columns:
                    child = etree.SubElement(item, col)
                    child.text = str(row[col])
            xml_str = etree.tostring(root, pretty_print=True, xml_declaration=True, encoding='UTF-8')
            suffix = f'_{idx}' if max_records and idx > 0 else ''
            filename = os.path.join(output_dir, f'{group_name}{suffix}.xml')
            with open(filename, 'wb') as f:
                f.write(xml_str)
            xml_files.append(filename)
            validate_xml(xml_str, group_name, suffix, script_dir)
    return xml_files

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
    df = read_excel(script_dir)
    column_name = get_column_name(df, args.column, args.by_number)
    xml_files = create_xml_files(df, column_name, args.max_records, output_dir, script_dir)
    create_zip_file(xml_files, script_dir)

if __name__ == '__main__':
    main()