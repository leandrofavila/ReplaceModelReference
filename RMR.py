from flask import Flask, render_template, request
import pandas as pd
import win32com.client as win32
import pythoncom
import codecs
import os
import shutil


app = Flask(__name__)


def get_part_number(file_paths):
    pythoncom.CoInitialize()
    status_dic = {}
    inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
    part_number_df = []
    for path in file_paths.values():
        full_path = get_path_selected_file(path)
        if not full_path:
            status_dic[path] = 'ipt não encontrado'
            return status_dic
        apprenticeDoc = inv.Open(full_path)
        oPropSets = apprenticeDoc.PropertySets
        PropertySet = oPropSets.Item("Design Tracking Properties")
        part_number_df.append((PropertySet.Item(2).Value, get_path_selected_file(os.path.basename(path))))
        apprenticeDoc.Close()
    part_number_df = pd.DataFrame(part_number_df, columns=['Part Number', 'Caminho'])
    old_parts = get_old_part(part_number_df['Part Number'].tolist())
    part_number_df['old_part'] = part_number_df['Part Number'].map(old_parts)
    return part_number_df


def get_path_selected_file(file):
    caminhos = set(codecs.open(r"C:\Users\pcp03\Desktop\asd\kaminhos.txt", "r").readlines())
    for line in caminhos:
        if ("\\" + (str(file))) in line:
            return line.rstrip().removeprefix('file://') if line else 'Arquivo não encontrado nos caminhos'



def get_ref_idw(old_file):
    caminhos = set(codecs.open(r"C:\Users\pcp03\Desktop\asd\kaminhos.txt", "r").readlines())
    idw_paths = {}
    for item in old_file:
        for line in caminhos:
            if ("\\" + (str(item)) + '.idw') in line:
                idw_paths[item] = line.rstrip().removeprefix('file://')
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
                try:
                    name_key = key[:-3] + "idw"
                    if not os.path.exists(os.path.join(os.path.dirname(key), os.path.basename(name_key))):
                        n_path = shutil.copy(value, os.path.join(os.path.dirname(key), os.path.basename(name_key)))
                    else:
                        status_dic[key] = 'Arquivo ja existe no destino.'
                        return status_dic
                except IOError as io_err:
                    print('erro ao COPIAR', io_err)
                    status_dic[key] = 'Erro ao copiar arquivo.'
                    continue

                rmr_no_dir = os.path.join(os.path.dirname(key) + "\\RMR.ipj")
                #    os.remove(rmr_no_dir)
                if not os.path.exists(rmr_no_dir):
                    ProjectNovo = inv.DesignProjectManager.DesignProjects.Add(36353, "RMR",
                                                                              os.path.join(os.path.dirname(key)))
                    ProjectNovo.Activate()
                try:
                    idw = inv.Open(n_path)
                    idw.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor.ReplaceReference(key)
                except Exception as erro:
                    print(f"Erro no save {erro}")
                    status_dic[key] = 'Erro ao copiar arquivo.'
                    continue
                try:
                    if idw.NeedsMigrating:
                        return 'Precisará migrar o arquivo antes.'
                    inv.FileSaveAs.AddFileToSave(idw, (key[:-3] + "idw"))
                    inv.FileSaveAs.ExecuteSave()
                    status_dic[key] = 'Salvo'
                    idw.Close()
                except Exception as err_:
                    print(f'Erro ao salvar {err_}')
                    status_dic["status"] = "Erro ao salvar."
            else:
                status_dic[key] = 'Arquivo não existe no diretorio'
        else:
            status_dic[key] = 'Não foi encontrado um IDW com dada referência.'
    return status_dic


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload_and_process', methods=['POST'])
def upload_and_process():
    file_list = {}
    if 'file' not in request.files:
        return render_template('index.html', table_html="", error="")
    files = request.files.getlist('file')

    for file in files:
        file_list[file] = file.filename

    df = get_part_number(file_list)

    try:
        ref_idw = get_ref_idw(df['old_part'].tolist())

    except Exception as e:
        print(e)
        return render_template('index.html', table_html="", error=df)

    df['ref_idw_paths'] = df['old_part'].map(ref_idw)
    to_idw_replace = dict(zip(df['Caminho'], df['ref_idw_paths']))

    df['status'] = df['Caminho'].apply(lambda x: execute_replace({str(x): to_idw_replace[str(x)]})[str(x)])
    table_html = df.to_html(classes='table table-striped', justify='left', escape=False, render_links=True)
    return render_template('index.html', table_html=table_html, error="")



if hasattr(pythoncom, '__file__'):
    print(pythoncom.__file__)


if __name__ == '__main__':
    try:
        app.run(host='0.0.0.0', port=8001)
    except Exception as err:
        print(err)
