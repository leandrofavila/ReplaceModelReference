from flask import Flask, render_template, request
import pandas as pd
import win32com.client as win32
import pythoncom
import codecs
import os
import shutil


app = Flask(__name__)


def get_part_number(full_path):
    print('get_part_number')
    pythoncom.CoInitialize()
    # Recebe caminho do arquivo na rede vindo do txt
    inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
    try:
        apprenticeDoc = inv.Open(full_path)
        oPropSets = apprenticeDoc.PropertySets
        PropertySet = oPropSets.Item("Design Tracking Properties")
        # Pega o part number do arquvio
        part_number = PropertySet.Item(2).Value
        apprenticeDoc.Close()
    except Exception as error_abrir_apprentice:
        part_number = 'Inventor não abriu o arquivo'
        print(error_abrir_apprentice)
    return part_number


def get_path_selected_file(file):
    print('get_path_selected_file')
    caminhos = set(codecs.open(r"Z:\PCP\Leandro\kaminhos.txt", "r").readlines())
    for line in caminhos:
        if ("\\" + (str(file))) in line:
            return line.rstrip().removeprefix('file://') if line else 'Caminho do arquivo não encontrado.'


def get_ref_idw(old_file):
    print('get_ref_idw')
    caminhos = set(codecs.open(r"Z:\PCP\Leandro\kaminhos.txt", "r").readlines())
    for line in caminhos:
        if ("\\" + (str(old_file)) + '.idw') in line:
            idw_paths = line.rstrip().removeprefix('file://')
            return idw_paths
    else:
        return 'IDW de referência não encontrado.'



def get_old_part(file):
    print('get_old_part')
    value = ''.join(cd for cd in file if cd.isdigit())
    old_pt = value[-6:].replace('.', '').lstrip('0')
    old_part = old_pt if 5 <= len(str(old_pt)) <= 6 else 'Não é um código padrão'
    return old_part


def copy_to_new_dir(ref_idw, full_path):
    print('copy_to_new_dir')
    if os.path.isabs(full_path):
        try:
            name_key = full_path[:-3] + "idw"
            if not os.path.exists(os.path.join(os.path.dirname(full_path), os.path.basename(name_key))):
                n_path = shutil.copy(ref_idw, os.path.join(os.path.dirname(full_path), os.path.basename(name_key)))
                return n_path
            else:
                return 'Arquvio ja existe no diretório'
        except Exception as err_copy:
            print(err_copy)
            return 'Erro ao copiar'


def create_ipj(full_path):
    print('create_ipj')
    rmr_no_dir = os.path.join(os.path.dirname(full_path) + "\\RMR.ipj")
    if not os.path.exists(rmr_no_dir):
        inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
        ProjectNovo = inv.DesignProjectManager.DesignProjects.Add(36353, "RMR",
                                                                  os.path.join(os.path.dirname(full_path)))
        ProjectNovo.Activate()



def execute_replace(n_path, full_path):
    print('execute_replace')
    try:
        inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
        idw = inv.Open(n_path)
        idw.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor.ReplaceReference(full_path)
        inv.FileSaveAs.AddFileToSave(idw, n_path)
        inv.FileSaveAs.ExecuteSave()
        idw.Close()
        return 'Salvo.'
    except Exception as error_replace:
        print(f"Erro no save {error_replace}")
        return 'Erro ao aplicar o replace.'



@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload_and_process', methods=['POST'])
def upload_and_process():
    print('upload_and_process')
    if 'file' not in request.files:
        return render_template('index.html', table_html="", error="")
    files = request.files.getlist('file')

    df = pd.DataFrame(columns=['Part Number', 'Caminho', 'Old Part', 'IDW Referência', 'Status'])

    for file in files:
        df.loc[len(df)] = [None] * len(df.columns)
        full_path = get_path_selected_file(file.filename)
        df.loc[len(df)-1, 'Caminho'] = full_path
        if not full_path:
            continue

        part_number = get_part_number(full_path)
        df.loc[len(df)-1, 'Part Number'] = part_number
        if not part_number or part_number == 'Inventor não abriu o arquivo':
            continue

        old_parts = get_old_part(part_number)
        df.loc[len(df)-1, 'Old Part'] = old_parts
        if not old_parts or old_parts == 'Não é um código padrão':
            continue

        ref_idw = get_ref_idw(old_parts)
        df.loc[len(df)-1, 'IDW Referência'] = ref_idw
        if not ref_idw or ref_idw == 'IDW de referência não encontrado.':
            continue

        n_path = copy_to_new_dir(ref_idw, full_path)

        create_ipj(full_path)

        status = execute_replace(n_path, full_path)
        print('status', status)
        if status:
            df.loc[len(df)-1, 'Status'] = status
        else:
            df.loc[len(df)-1, 'Status'] = n_path


    df = df.sort_values(by=['Status']).reset_index(drop=True)
    df.index = df.index + 1
    table_html = df.to_html(classes='table table-striped', justify='left', escape=False, render_links=True)
    return render_template('index.html', table_html=table_html, error="")


if __name__ == '__main__':
    try:
        app.run(host='0.0.0.0', port=8001)
    except Exception as err:
        print(err)
