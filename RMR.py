from flask import Flask, render_template, request
import pandas as pd
import win32com.client as win32



def get_part_number():
    inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
    apprenticeDoc = inv.Open(r"T:\01_Projetos\Ped_1360_Cooxupe\14.Esteira_Transp_XXXXXX\1.LCC.029.436.ipt")
    oPropSets = apprenticeDoc.PropertySets
    PropertySet = oPropSets.Item("Design Tracking Properties")
    dts = {
        'path':["T:\01_Projetos\Ped_1360_Cooxupe\14.Esteira_Transp_XXXXXX\1.LCC.029.436.ipt"],
        'part_number':[PropertySet.Item(2).Value ]
    }
    return pd.DataFrame(dts)


def similar_cant():
    df = get_part_number()
    print(df)
    quit()
    #df['cod_item'] = df['cod_item'].astype(int)
    #df['qtde'] = df['qtde'].astype(int)
    #df = df.reset_index(drop=True)
    # converte a coluna cod_item pra um link para os pdfs. O problema disso é que o caminho em cada terminal e diferente ou depende do navegador
    #df['cod_item'] = df['cod_item'].apply(lambda x: f'<a href="http://localhost:8000/{x}.pdf" target="_blank"'f' download>{x}</a>')

    return df.to_html(classes='table table-striped', justify='left', escape=False)





app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html', table_html=similar_cant())


if __name__ == '__main__':
    app.run(debug=True)
