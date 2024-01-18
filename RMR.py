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


    part_number_df = pd.DataFrame(part_number_df, columns=['Descrição', 'Caminho'])
    get_old_part(part_number_df['Descrição'].tolist())
    df = part_number_df
    return part_number_df


def get_path_selected_file(file):
    caminhos = set(codecs.open(r"R:\Rubens\30_Prog\Caminhos.txt", "r", "utf-8").readlines())
    local_paths = {}
    for item in file:
        for line in caminhos:
            if ("\\" + (str(item))) in line:
                local_paths[item] = line.rstrip()
    return local_paths


def get_old_part(file):
    df = get_part_number(file)
    old_part = {}
    for idx, val in df.iterrows():
        old_pt = val['part_number'][-6:]. replace('.', '')
        print(old_pt)



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
        return render_template('index.html', table_html=table_html, error="")

    except Exception as e:
        return render_template('index.html', table_html="", error=str(e))


if __name__ == '__main__':
    app.run(host='10.40.3.48', port=8001, debug=True)