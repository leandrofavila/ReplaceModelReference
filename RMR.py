from flask import Flask, render_template, request
import pandas as pd
import win32com.client as win32
import pythoncom
import codecs
import os
import shutil


app = Flask(__name__)


def get_part_number(files):
    pythoncom.CoInitialize()
    file_paths = get_path_selected_file(files)
    if not file_paths:
        return
    inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
    part_number_df = []
    for path in file_paths.values():
        print('esse', path)
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
                local_paths[item] = line.rstrip() if line else 'Arquivo não encontrado nos caminhos'
    #print('local_paths', local_paths)
    return local_paths


def get_ref_idw(old_file):
    caminhos = set(codecs.open(r"R:\Rubens\30_Prog\Caminhos.txt", "r", "utf-8").readlines())
    idw_paths = {}
    for item in old_file:
        for line in caminhos:
            if ("\\" + (str(item)) + '.idw') in line:
                idw_paths[item] = line.rstrip()
                #print(line.rstrip())
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
    status_dic = {}
    for key, value in list_to_replace.items():
        if isinstance(key, str) and isinstance(value, str):
            if os.path.isabs(key) and os.path.isabs(value):
                dasda = ''
                try:
                    name_key = key[:-3] + "idw"
                    dasda = shutil.copy(value, os.path.join(os.path.dirname(key), os.path.basename(name_key)))
                except IOError as io_err:
                    print('erro ao COPIAR', io_err)
                    status_dic[key] = 'Erro ao copiar arquivo.'
                idw = inv.Open(dasda)
                idw.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor.ReplaceReference(key)

                try:
                    inv.FileSaveAs.AddFileToSave(idw, (key[:-3] + "idw"))
                    inv.FileSaveAs.ExecuteSave()
                    status_dic[key] = 'Salvo'
                except Exception as err_:
                    #print(f'Erro ao salvar {err_}')
                    status_dic["status"] = f"Error: {err_}"
            else:
                #print('Arquivo não existe no diretorio - ', key)
                status_dic[key] = 'Arquivo não existe no diretorio'
                continue
        else:
            status_dic[key] = 'Não foi encontrado um IDW com dada referência.'
    return status_dic


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

        try:
            ref_idw = get_ref_idw(df['old_part'].tolist())
        except Exception as e:
            print(e)
            return render_template('index.html', table_html="", error=str('Arquvivo não encontrado.'))

        df['ref_idw_paths'] = df['old_part'].map(ref_idw)
        print(df.to_string())
        to_idw_replace = dict(zip(df['Caminho'], df['ref_idw_paths']))
        df['status'] = df['Caminho'].map(execute_replace(to_idw_replace))
        table_html = df.to_html(classes='table table-striped', justify='left', escape=False)
        return render_template('index.html', table_html=table_html, error="")

    except Exception as e:
        return render_template('index.html', table_html="", error=str(e))



if hasattr(pythoncom, '__file__'):
    print(pythoncom.__file__)



if __name__ == '__main__':
    try:
        app.run(host='0.0.0.0', port=8001, debug=True)
    except Exception as err:
        print(err)
