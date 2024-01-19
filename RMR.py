from flask import Flask, render_template, request
import pandas as pd
import win32com.client as win32
import pythoncom
import codecs


app = Flask(__name__)


def get_part_number(files):
    pythoncom.CoInitialize()
    file_paths = get_path_selected_file(files)
    inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
    part_number_df = []
    for path in file_paths.values():
        apprenticeDoc = inv.Open(path)
        oPropSets = apprenticeDoc.PropertySets
        PropertySet = oPropSets.Item("Design Tracking Properties")
        part_number_df.append((PropertySet.Item(2).Value, path))
    part_number_df = pd.DataFrame(part_number_df, columns=['Part Number', 'Caminho'])
    old_parts = get_old_part(part_number_df['Part Number'].tolist())
    part_number_df['old_part'] = part_number_df['Part Number'].map(old_parts)
    return part_number_df


def get_path_selected_file(file):
    caminhos = set(codecs.open(r"R:\Rubens\30_Prog\Caminhos.txt", "r", "utf-8").readlines())
    local_paths = {}
    for item in file:
        for line in caminhos:
            if ("\\" + (str(item))) in line:
                local_paths[item] = line.rstrip()
    return local_paths


def get_ref_idw(old_file):
    caminhos = set(codecs.open(r"R:\Rubens\30_Prog\Caminhos.txt", "r", "utf-8").readlines())
    idw_paths = {}
    for item in old_file:
        for line in caminhos:
            if ("\\" + (str(item)) + '.idw') in line:
                idw_paths[item] = line.rstrip()
    return idw_paths


def get_old_part(file):
    old_part = {}
    for val in file:
        value = ''.join(cd for cd in val if cd.isdigit())
        old_pt = value[-6:].replace('.', '').lstrip('0')
        old_part[val] = old_pt if 5 <= len(str(old_pt)) <= 6 else 'Não é um código padrão'
    return old_part


def execute_replace(list_to_replace):
    inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
    for key, value in list_to_replace.items():
        if key and value:
            idw = inv.Open(value)
            idw.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor.ReplaceReference(key)
            inv.FileSaveAs.AddFileToSave(idw, (key[:-3] + "idw"))
            inv.FileSaveAs.ExecuteSaveAs()




@app.route('/', methods=['GET', 'POST'])
def process_file():
    try:
        if 'file' not in request.files:
            return render_template('index.html', table_html="", error="")
        files = request.files.getlist('file')
        file_list = []
        for file in files:
            file_list.append(file.filename)
        df = get_part_number(file_list)
        table_html = df.to_html(classes='table table-striped', justify='left', escape=False)

        df_back = df.copy()
        df_back['ref_idw_paths'] = df_back['old_part'].map(get_ref_idw(df_back['old_part'].tolist()))

        to_idw_replace = dict(zip(df_back['Caminho'], df_back['ref_idw_paths']))
        execute_replace(to_idw_replace)
        return render_template('index.html', table_html=table_html, error="")

    except Exception as e:
        return render_template('index.html', table_html="", error=str(f'Permission Denied - {e}'))


if hasattr(pythoncom, '__file__'):
    print(pythoncom.__file__)

if __name__ == '__main__':
    app.run(host='10.40.3.48', port=8001, debug=True)
