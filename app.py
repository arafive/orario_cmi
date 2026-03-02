
import re
import os
import xlrd
import locale
locale.setlocale(locale.LC_TIME, 'it_IT.UTF-8')

import pandas as pd

from datetime import datetime
from collections import defaultdict

from flask import Flask, Response, render_template_string, redirect, url_for

dict_mesi = {
    'gennaio': ['dicembre', 'febbraio'],
    'febbraio': ['gennaio', 'marzo'],
    'marzo': ['febbraio', 'aprile'],
    'aprile': ['marzo', 'maggio'],
    'maggio': ['aprile', 'giugno'],
    'giugno': ['maggio', 'luglio'],
    'luglio': ['giugno', 'agosto'],
    'agosto': ['luglio', 'settembre'],
    'settembre': ['agosto', 'ottobre'],
    'ottobre': ['settembre', 'novembre'],
    'novembre': ['ottobre', 'dicembre'],
    'dicembre': ['novembre', 'gennaio'],
}

dict_map_mesi = {
    'gennaio': '01',
    'febbraio': '02',
    'marzo': '03',
    'aprile': '04',
    'maggio': '05',
    'giugno': '06',
    'luglio': '07',
    'agosto': '08',
    'settembre': '09',
    'ottobre': '10',
    'novembre': '11',
    'dicembre': '12',
}

numeri_di_telefono = defaultdict(lambda: "???")
numeri_di_telefono.update({
    'Elisabetta Fiori': 502,
    'Stefano Murrau': 504,
    'Francesca Giannoni': 505,
    'Giuseppina Rossi': 507,
    'Veronica Bonati': 508,
    'Laura Pedemonte': 509,
    'Marco Tizzi': 510,
    'Barbara Turato': 511,
    'Alfredo Crosetti': 512,
    'Federico Pedemonte': 513,
    'Martina Raffellini': 514,
    'Luca Rusca': 515,
    'Alfredo Bertolone': 516,
    'Federico Cassola': 517,
    'Federica Martina': 518,
    'Federico Buscemi': 519,
    'Dario Hourngir': 520,
    'Luca Onorato': 521,
    'Paolo Gollo': 523,
    'Fabio Gardella': 526,
    'Giorgia Galvani Vezzi': 527,
    'Antonio Iengo': 528,
    'Massimiliano Coppolecchia': 529,
    'Andrea Cavallo': 530,
    'Angelo Forestieri': 531,
    'Mauro Damonte': 532,
    'Vito Danieli': 534,
    'Davide Sacchetti': 535,
    'Luca Napolitano': 533,
    'Edoardo Rocca': 537,
    'Paolo Oliveri': 538,
    'Daniele Carnevale': 539,
    'Mario Lecca': 540,
    'Daniele Luppi': 541,
    'Luca Repetto': 542,
    'Tiziana Poggio': 543,
    # 'Roberto Cresta': ,
})

app = Flask(__name__)

oggi = pd.Timestamp.now()
# oggi = pd.to_datetime('2025-07-10')
anno, mese, giorno = oggi.year, oggi.strftime('%B'), oggi.day
print(f'{oggi=}')

spazi = '&nbsp;' * 8
    
# %%
@app.route('/')
def index():
    return redirect(url_for('mostra_tabella', nome_file=f"{mese.capitalize()}_{anno}"))


@app.route('/link/<nome_file>')
def mostra_tabella(nome_file):
    oggi = pd.Timestamp.now()
    # oggi = pd.to_datetime('2025-07-10')
    anno, mese, giorno = oggi.year, oggi.strftime('%B'), oggi.day
    print(f'{nome_file=}')
    print(f'{oggi=}\n')

    try:
        mese_corrente, anno_corrente = nome_file.split('_')
        mese_corrente, anno_corrente = mese_corrente.lower(), int(anno_corrente)
    except:
        return f"Formato nome file non valido: {nome_file}", 400

    percorso = f'link/{mese_corrente.capitalize()} {anno_corrente}.xls'
    data_corrente = pd.to_datetime(f'{anno_corrente}-{dict_map_mesi[mese_corrente]}-01')
    if oggi < data_corrente:
        nome_foglio = 'prima emissione'
    else:
        nome_foglio = 'orari successivi'
    print(f'{percorso=}')
    print(f'{nome_foglio=}\n')

    ### Carica i dati con pandas (solo i valori)
    df = pd.read_excel(percorso, skiprows=1, engine='xlrd', sheet_name=nome_foglio)
    df = df.fillna("")  # mantiene le celle vuote
    df.columns = ["" for x in df.columns]
    df = df.iloc[0:df.index[(df == "").all(axis=1)][-1]]

    ### Leggo versione, data, eventuali impegni
    df_vd = pd.read_excel(percorso, engine='xlrd', sheet_name=nome_foglio)
    versione = df_vd.iloc[0].dropna().iloc[-2]
    data = pd.to_datetime(df_vd.iloc[0].dropna().iloc[-1]).round('1s')
    del (df_vd)
    
    titolo = percorso.split('/')[-1].split('.xls')[0]
    print(f'{versione=}')
    print(f'{data=}')
    print(f'{titolo=}\n')

    ### Calcola mese e anno precedente/successivo
    mese_prima, mese_dopo = dict_mesi[mese_corrente][0], dict_mesi[mese_corrente][1]

    anno_prima = anno_corrente - 1 if mese_corrente == 'gennaio' and mese_prima == 'dicembre' else anno_corrente
    anno_dopo = anno_corrente + 1 if mese_corrente == 'dicembre' and mese_dopo == 'gennaio' else anno_corrente

    nome_file_prima, nome_file_dopo = f'{mese_prima.capitalize()}_{anno_prima}', f'{mese_dopo.capitalize()}_{anno_dopo}'

    url_mese_prima, url_mese_dopo = url_for('mostra_tabella', nome_file=nome_file_prima), url_for('mostra_tabella', nome_file=nome_file_dopo)
    
    print(f'{url_mese_prima=}')
    print(f'{url_mese_dopo=}\n')

    if os.path.exists(f'./link/{mese_dopo.capitalize()} {anno_dopo}.xls'):
        html_titolo = f'''
        <h1 class="titolo" style="display: flex; align-items: center; justify-content: center; gap: 10px;">
          <a href="{url_mese_prima}" style="text-decoration:none; font-size: 2em;">&#8592;</a> {spazi}
          {titolo} {spazi}
          <a href="{url_mese_dopo}" style="text-decoration:none; font-size: 2em;">&#8594;</a>
        </h1>
        <br>
        '''
    else:
        html_titolo = f'''
        <h1 class="titolo" style="display: flex; align-items: center; justify-content: center; gap: 10px;">
          <a href="{url_mese_prima}" style="text-decoration:none; font-size: 2em;">&#8592;</a> {spazi}
          {titolo} {spazi}
        </h1>
        <br>
        '''
        
    print(f'{html_titolo=}\n')

    html_numeri_agenzia = '<a href="https://reportistica.arpal.org:8443/ords/r/arpal/ict-visualizzazione/elenco-del-telefono" target="_blank">Tutti i numeri d\'agenzia</a>'
    torna_ad_oggi_url = url_for('mostra_tabella', _anno=anno, nome_file=f"{mese.capitalize()}_{anno}")
    torna_ad_oggi = f'<a href="{torna_ad_oggi_url}">Torna ad oggi</a>'
    orari_OVG = f'<a href="https://cmi-servizi.arpal.liguria.it/orario/Turni_OVG/{anno_corrente}/{mese_corrente.capitalize()}_{anno_corrente}.pdf" target="_blank">Turni OVG</a>'
    html_info = f'<h2 class="versione-orario">Versione: {versione} {spazi} {data.strftime("%d/%m/%Y %H:%M:%S")} {spazi} {html_numeri_agenzia} {spazi} {orari_OVG} {spazi} {torna_ad_oggi}</h2>'

    print(f'{html_numeri_agenzia=}')
    print(f'{torna_ad_oggi=}')
    print(f'{orari_OVG=}\n')
    print(f'{html_info=}\n')

    html_tabella = html_titolo + '\n' + html_info + '\n'
    print(f'{html_tabella=}\n')

    ### Legge i colori con xlrd
    book = xlrd.open_workbook(percorso, formatting_info=True)
    sheet, palette = book.sheet_by_name(nome_foglio), book.colour_map
 
    skiprows_pandas, num_righe, background_colors = 1, df.shape[0], []

    for i in range(num_righe):
        # Sposto di una riga indietro rispetto a pandas per la lettura colori
        # !!! row_idx non può essere < 0
        row_idx = skiprows_pandas + i + 1
        if row_idx < 0: # Se siamo all’inizio, nessun colore o colore None
            print(f'{i=}, {row_idx=}, row_idx<0')
            row_colors = [None] * sheet.ncols
        else:
            print(f'{i=}, {row_idx=}, row_idx>=0')
            row_colors = []
            for col_idx in range(sheet.ncols):
                xf_index = sheet.cell_xf_index(row_idx, col_idx)
                xf = book.xf_list[xf_index]
                bg_index = xf.background.pattern_colour_index
                rgb = palette.get(bg_index)
                if rgb:
                    hex_color = "#{:02X}{:02X}{:02X}".format(*rgb)
                    if hex_color.upper() == "#C0C0C0":
                        hex_color = "#D3D3D3"
                    elif hex_color.upper() == "#FFFF00":
                        hex_color = "#ffea00"
                    # elif hex_color.upper() == "#FF00FF":
                    #     hex_color = "#FF00FF"
                    elif hex_color.upper() == "#FF0000":
                        hex_color = "#E50000"
                    elif hex_color.upper() == "#FF6600":
                        hex_color = "#FF7F03"
                    elif hex_color.upper() == "#00FFFF":
                        hex_color = "#38ffff"
                else:
                    hex_color = None
                row_colors.append(hex_color)
        background_colors.append(row_colors)

    print()
    
    ### Costruisce la tabella HTML
    html_tabella += '<table border="1" style="margin:auto; border-collapse:collapse;">\n'
    html_tabella += '  <thead><tr>' + ''.join(f'<th>{col}</th>' for col in df.columns) + '</tr></thead>\n'
    html_tabella += '  <tbody>\n'

    ################################################################################
    ################################################################################
    ################################################################################

    righe_principali, righe_con_C = [], []
    bool_riga_numeri_giorni, bool_riga_nomi_giorni = True, True

    colonna_gialla = set()
    for col_idx, val in enumerate(df.iloc[1]):  # seconda riga (indice 1)
        try:
            if int(str(val).strip()) == giorno and mese.capitalize() in titolo and anno == anno_corrente:
                # print(f'{col_idx=}, {val=}, trovata colonna gialla')
                col_riga_gialla = col_idx
                colonna_gialla.add(col_idx)
        except ValueError:
            continue  # ignora celle non numeriche

    ################################

    for r_idx, (_, row) in enumerate(df.iterrows()):
        row_values, num_colonne = [str(x).strip() for x in row], len(df.columns)

        if all(val == "" for val in row_values):
            html_riga = f'    <tr><td colspan="{num_colonne}" style="border:none; height: 20px;"></td></tr>\n'
            righe_principali.append(html_riga)
            continue

        contiene_C = any(val == 'C' for val in row_values)
        html_riga = '    <tr class="riga_da_evidenziare">'

        for c_idx, cell in enumerate(row):
            # print(f'{c_idx=}, {cell=}, {col_idx=}, {r_idx=}')
            colore, style_parts = background_colors[r_idx][c_idx], []

            if colore:
                style_parts.append(f'background-color: {colore};')

            ### Se la colonna è una di quelle da evidenziare e non ha colore preesistente
            if c_idx in colonna_gialla and not colore:
                style_parts.append('background-color: #ffe900;')

            cell_str = str(cell).strip()

            ### Tooltip + grassetto nome se in colonna_gialla c'è 'P' o 'CP'
            tooltip_attr = ""
            if c_idx == 0:
                numero = numeri_di_telefono.get(cell_str, "???")
                tooltip_attr = f' title="{cell_str}: {numero}" style="cursor: help;"'

            for col_idx in colonna_gialla:
                valore = df.iloc[r_idx, col_idx]
                if type(valore) == str:
                    if len(valore) == 1:
                        if valore in ['P', 'H', 'I']:
                            style_parts.append('color: #E50000; font-weight: bold;')
                            break
                        
                    elif len(valore) > 1:
                        if valore == 'CP':
                            style_parts.append('color: #FF00FF; font-weight: normal;')
                            break
                        
                        elif valore == 'SI':
                            style_parts.append('color: #E50000; font-weight: bold;')
                            break
                        
                        elif valore.endswith('H'):
                            if valore[-2:] == 'CH':
                                style_parts.append('color: #FF00FF; font-weight: normal;')
                                break
                            
                            else:
                                style_parts.append('color: #E50000; font-weight: bold;')
                                break

            ################################

            if cell_str == 'TC':
                cell_html = (
                    '<span style="color:#2CA02C; font-weight:bold;">T</span>'
                    '<span style="color:#000000; font-weight:bold;">C</span>'
                )
            elif cell_str == 'MH':
                cell_html = (
                    '<span style="color:#278C27; font-weight:bold;">M</span>'
                    '<span style="color:#E50000; font-weight:bold;">H</span>'
                )
            elif cell_str == 'MT':
                cell_html = (
                    '<span style="color:#278C27; font-weight:bold;">M</span>'
                    '<span style="color:#2CA02C; font-weight:bold;">T</span>'
                )
            elif cell_str == 'MCH':
                cell_html = (
                    '<span style="color:#278C27; font-weight:bold;">M</span>'
                    '<span style="color:#E50000; font-weight:bold;">CH</span>'
                )
            elif cell_str == 'TCH':
                cell_html = (
                    '<span style="color:#2CA02C; font-weight:bold;">T</span>'
                    '<span style="color:#E50000; font-weight:bold;">CH</span>'
                )
            elif cell_str == 'MSH':
                cell_html = (
                    '<span style="color:#278C27; font-weight:bold;">M</span>'
                    '<span style="color:#E50000; font-weight:bold;">SH</span>'
                )
            elif cell_str == 'M':
                style_parts.append('color: #278C27; font-weight: bold;')
                cell_html = cell_str
            elif cell_str == 'C':
                style_parts.append('color: #000000; font-weight: bold;')
                cell_html = cell_str
            elif cell_str == 'T':
                style_parts.append('color: #2CA02C; font-weight: bold;')
                cell_html = cell_str
            elif cell_str in ['F', 'RO', 'RP', 'RD', 'X']:
                style_parts.append('color: #1F77B4; font-weight: bold;')
                cell_html = cell_str
            elif cell_str in ['P', 'H', 'I', 'SI', 'SH']:
                style_parts.append('color: #E50000; font-weight: bold;')
                cell_html = cell_str
            elif cell_str in ['CP', 'CH']:
                style_parts.append('color: #FF00FF; font-weight: bold;')
                cell_html = cell_str
            elif cell_str == '.':
                cell_html = '·'
                style_parts.append('color: #1F77B4; font-weight: bold;')
            elif cell_str == 'Ì':
                cell_html = '✚'
                style_parts.append('color: #E50000; font-weight: bold;')
            else:
                cell_html = cell_str

            stile = f' style="{" ".join(style_parts)}"' if style_parts else ""
            html_riga += f'<td{stile}{tooltip_attr}>{cell_html}</td>'
                
        html_riga += '</tr>\n'
        # print(f'{html_riga=}')
       
        try:
            if col_idx == col_riga_gialla and r_idx == 1:
                print('Definita riga_numeri_giorni')
                riga_numeri_giorni = html_riga
                print(f'{riga_numeri_giorni=}')

            if col_idx == col_riga_gialla and r_idx == 2:
                print('Definita riga_nomi_giorni')
                riga_nomi_giorni = html_riga
                print(f'{riga_nomi_giorni=}')
        except UnboundLocalError:
            pass

        if contiene_C:
            righe_con_C.append(html_riga)
        else:
            righe_principali.append(html_riga)
            
    try:
        righe_principali.append(riga_nomi_giorni)
        righe_principali.append(riga_numeri_giorni)
    except UnboundLocalError:
        pass

    ### Inserisce righe con 'C' dopo la terza riga vuota
    corpo_tabella, count_spazi, inserito = [], 0, False
    riga_spazio = f'    <tr><td colspan="{num_colonne}" style="border:none; height: 20px;"></td></tr>\n'

    for riga in righe_principali:
        corpo_tabella.append(riga)
        if 'colspan' in riga and 'height' in riga:
            count_spazi += 1
        if count_spazi == 5 and not inserito:
            corpo_tabella.extend(righe_con_C)
            corpo_tabella.append(riga_spazio)
            inserito = True

    ### Se ci sono meno di 3 spazi, inseriscile alla fine
    if not inserito:
        corpo_tabella.extend(righe_con_C)
        corpo_tabella.append(riga_spazio)

    html_tabella += '  <tbody>\n' + ''.join(corpo_tabella) + '  </tbody>\n</table>'

    ### Rimuove righe vuote doppie consecutive
    html_tabella = re.sub(
        r'(<tr>\s*<td colspan="\d+" style="border:none; height: ?\d+px;"></td>\s*</tr>\s*){2,}',
        r'\1',
        html_tabella,
        flags=re.IGNORECASE
    )

    # print(f'{html_tabella=}')
    
    ### Inserisce la tabella nell'HTML base
    with open("index.html", encoding="utf-8") as f:
        html_base = f.read()

    link_css = '<link rel="stylesheet" type="text/css" href="/static/stile.css">'
    html_base = html_base.replace("</head>", f"  {link_css}\n</head>")
    html_completo = html_base.replace("<!-- La tabella verrà messa qui -->", html_tabella)
    # print(f'\n\n\n{html_completo}')

    return Response(html_completo, mimetype="text/html")


if __name__ == '__main__':
    app.run(debug=True)
