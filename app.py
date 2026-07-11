
from flask import Flask, Response, redirect, url_for
from collections import defaultdict
from datetime import datetime
import pandas as pd
import re
import os
import xlrd
import locale
locale.setlocale(locale.LC_TIME, 'it_IT.UTF-8')

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
    'Roberto Cresta': 209,
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
    data_corrente = pd.to_datetime(
        f'{anno_corrente}-{dict_map_mesi[mese_corrente]}-01')
    if oggi < data_corrente:
        nome_foglio = 'prima emissione'
    else:
        nome_foglio = 'orari successivi'
    print(f'{percorso=}')
    print(f'{nome_foglio=}\n')

    # Carica i dati con pandas (solo i valori)
    df = pd.read_excel(percorso, skiprows=1, engine='xlrd',
                       sheet_name=nome_foglio)
    df = df.fillna("")  # mantiene le celle vuote
    df.columns = ["" for x in df.columns]
    df = df.iloc[0:df.index[(df == "").all(axis=1)][-1]]

    # Leggo versione, data, eventuali impegni
    df_vd = pd.read_excel(percorso, engine='xlrd', sheet_name=nome_foglio)
    versione = df_vd.iloc[0].dropna().iloc[-2]
    data = pd.to_datetime(df_vd.iloc[0].dropna().iloc[-1]).round('1s')
    del (df_vd)

    titolo = percorso.split('/')[-1].split('.xls')[0]
    print(f'{versione=}')
    print(f'{data=}')
    print(f'{titolo=}\n')

    # Calcola mese e anno precedente/successivo
    mese_prima, mese_dopo = dict_mesi[mese_corrente][0], dict_mesi[mese_corrente][1]

    anno_prima = anno_corrente - \
        1 if mese_corrente == 'gennaio' and mese_prima == 'dicembre' else anno_corrente
    anno_dopo = anno_corrente + \
        1 if mese_corrente == 'dicembre' and mese_dopo == 'gennaio' else anno_corrente

    nome_file_prima, nome_file_dopo = f'{mese_prima.capitalize()}_{anno_prima}', f'{mese_dopo.capitalize()}_{anno_dopo}'

    url_mese_prima, url_mese_dopo = url_for('mostra_tabella', nome_file=nome_file_prima), url_for(
        'mostra_tabella', nome_file=nome_file_dopo)

    print(f'{url_mese_prima=}')
    print(f'{url_mese_dopo=}\n')

    if os.path.exists(f'./link/{mese_dopo.capitalize()} {anno_dopo}.xls'):
        html_titolo = f'''
        <h1 class="titolo" style="display: flex; align-items: center; justify-content: center; gap: 10px;">
          #8592;</a> {spazi}
          <a href="{url_mese_prima}" style="text-decoration:none; font-size: 2em;">&
          {titolo} {spazi}
          #8594;</a>
          <a href="{url_mese_dopo}" style="text-decoration:none; font-size: 2em;">&
        </h1>
        <br>
        '''
    else:
        html_titolo = f'''
        <h1 class="titolo" style="display: flex; align-items: center; justify-content: center; gap: 10px;">
          #8592;</a> {spazi}
          <a href="{url_mese_prima}" style="text-decoration:none; font-size: 2em;">&
          {titolo} {spazi}
        </h1>
        <br>
        '''

    print(f'{html_titolo=}\n')

    html_numeri_agenzia = '<a href="https://reportistica.arpal.org:8443/ords/r/arpal/ict-visualizzazione/elenco-del-telefono" target="_blank">Tutti i numeri d\'agenzia <img src="/static/telefono_fisso.svg" class="emoji-icon-gif"></a>'
    torna_ad_oggi_url = url_for(
        'mostra_tabella', _anno=anno, nome_file=f"{mese.capitalize()}_{anno}")
    torna_ad_oggi = f'<a href="{torna_ad_oggi_url}">Torna ad oggi <img src="/static/calendario.svg" class="emoji-icon-gif"></a>'
    scelta_emoji = '<a href="https://emojiterra.com/animated-emoji/" target="_blank">Scegli il tuo emoji <img src="/static/faccia_sottosopra.gif" class="emoji-icon-gif"></a>'
    orari_OVG = f'<a href="https://cmi-servizi.arpal.liguria.it/orario/Turni_OVG/{anno_corrente}/{mese_corrente.capitalize()}_{anno_corrente}.pdf" target="_blank">Vigilanti Gialli <img src="/static/occhiaie.gif" class="emoji-icon-gif"></a>'
    html_info = f'<h2 class="versione-orario">Versione: {versione} {spazi} {data.strftime("%d/%m/%Y %H:%M:%S")} {spazi} {html_numeri_agenzia} {spazi} {orari_OVG} {spazi} {torna_ad_oggi} {spazi} {scelta_emoji}</h2>'

    # print(f'{html_numeri_agenzia=}')
    # print(f'{torna_ad_oggi=}')
    # print(f'{orari_OVG=}\n')
    # print(f'{html_info=}\n')

    html_tabella = html_titolo + '\n' + html_info + '\n'
    # print(f'{html_tabella=}\n')

    # Legge i colori con xlrd
    book = xlrd.open_workbook(percorso, formatting_info=True)
    sheet, palette = book.sheet_by_name(nome_foglio), book.colour_map

    skiprows_pandas, num_righe, background_colors = 1, df.shape[0], []

    celle_R = set()
    celle_HR = set()
    celle_TR = set()
    celle_PuntoR = set()
    celle_CPR = set()
    celle_MR = set()
    celle_AP = set()
    celle_ACP = set()
    celle_Rep = set()
    celle_Mail_controllo = set()
    celle_Mamma = set()

    for i in range(num_righe):
        # Sposto di una riga indietro rispetto a pandas per la lettura colori
        # !!! row_idx non può essere < 0
        row_idx = skiprows_pandas + i + 1
        if row_idx < 0:  # Se siamo all’inizio, nessun colore o colore None
            # print(f'{i=}, {row_idx=}, row_idx<0')
            row_colors = [None] * sheet.ncols
        else:
            # print(f'{i=}, {row_idx=}, row_idx>=0')
            row_colors = []
            for col_idx in range(sheet.ncols):
                xf_index = sheet.cell_xf_index(row_idx, col_idx)
                xf = book.xf_list[xf_index]
                bg_index = xf.background.pattern_colour_index
                rgb = palette.get(bg_index)

                giorno_settimana = str(
                    sheet.cell_value(4, col_idx)).strip().lower()
                # print(giorno_settimana)
                if rgb:
                    hex_color = "#{:02X}{:02X}{:02X}".format(*rgb)
                    if hex_color.upper() == "#C0C0C0":
                        hex_color = "#D3D3D3"
                    elif hex_color.upper() == "#FFFF00":
                        hex_color = None
                        if sheet.cell_value(row_idx, col_idx) == '':
                            celle_R.add((i, col_idx))
                        elif sheet.cell_value(row_idx, col_idx) == 'H':
                            celle_HR.add((i, col_idx))
                        elif sheet.cell_value(row_idx, col_idx) == 'T':
                            celle_TR.add((i, col_idx))
                        elif sheet.cell_value(row_idx, col_idx) == '.':
                            celle_PuntoR.add((i, col_idx))
                        elif sheet.cell_value(row_idx, col_idx) == 'CP':
                            celle_CPR.add((i, col_idx))
                        elif sheet.cell_value(row_idx, col_idx) == 'M':
                            celle_MR.add((i, col_idx))
                        else:
                            hex_color = "#cfe619"

                    elif hex_color.upper() == "#FF0000":
                        if sheet.cell_value(row_idx, col_idx) == '':
                            hex_color = None
                            celle_AP.add((i, col_idx))
                        else:
                            hex_color = "#E50000"

                    elif hex_color.upper() == "#FF00FF":
                        hex_color = None
                        celle_ACP.add((i, col_idx))

                    elif hex_color.upper() == "#FF99CC":
                        celle_Mamma.add((i, col_idx))
                        if giorno_settimana == 'sab' or giorno_settimana == 'dom':
                            hex_color = '#D3D3D3'
                        else:
                            hex_color = None

                    elif hex_color.upper() == "#FF6600":
                        hex_color = "#FF7F03"

                    elif hex_color.upper() == "#00FFFF":
                        celle_Rep.add((i, col_idx))
                        if giorno_settimana == 'sab' or giorno_settimana == 'dom':
                            hex_color = '#D3D3D3'
                        else:
                            hex_color = None

                    elif hex_color.upper() == "#CCFFCC":
                        celle_Mail_controllo.add((i, col_idx))
                        if giorno_settimana == 'sab' or giorno_settimana == 'dom':
                            hex_color = '#D3D3D3'
                        else:
                            hex_color = None

                else:
                    hex_color = None
                row_colors.append(hex_color)
        background_colors.append(row_colors)

    print()

    # Costruisce la tabella HTML
    html_tabella += '<table border="1" style="margin:auto; border-collapse:collapse;">\n'
    html_tabella += '  <thead><tr>' + \
        ''.join(f'<th>{col}</th>' for col in df.columns) + '</tr></thead>\n'
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

            # Se la colonna è una di quelle da evidenziare e non ha colore preesistente
            if c_idx in colonna_gialla and not colore:
                # style_parts.append('background-color: #ffe900;')
                style_parts.append('background-color: #FFD700;')

            cell_str = str(cell).strip()

            # Tooltip + grassetto nome se in colonna_gialla c'è 'P' o 'CP'
            tooltip_attr = ""
            if c_idx == 0:
                numero = numeri_di_telefono.get(cell_str, "???")
                tooltip_attr = f' title="{cell_str}: {numero}" style="cursor: help;"'
                if cell_str == 'Daniele Carnevale':
                    cell_str = 'Daniele Carnevale&nbsp;<img src="/static/saturno.gif" class="emoji-icon">'
                elif cell_str == 'Martina Raffellini':
                    cell_str = 'Martina Raffellini&nbsp;<img src="/static/scorpione.gif" class="emoji-icon">'
                elif cell_str == 'Edoardo Rocca':
                    cell_str = 'Edoardo Rocca&nbsp;<img src="/static/ape.gif" class="emoji-icon">'
                elif cell_str == 'Angelo Forestieri':
                    cell_str = 'Angelo Forestieri&nbsp;<img src="/static/geco.gif" class="emoji-icon">'
                elif cell_str == 'Davide Sacchetti':
                    cell_str = 'Davide Sacchetti&nbsp;<img src="/static/scimmia.gif" class="emoji-icon">'
                elif cell_str == 'Luca Repetto':
                    cell_str = 'Luca Repetto&nbsp;<img src="/static/onda.gif" class="emoji-icon">'
                elif cell_str == 'Andrea Cavallo':
                    cell_str = 'Andrea Cavallo&nbsp;<img src="/static/cavallo.gif" class="emoji-icon">'
                elif cell_str == 'Paolo Gollo':
                    cell_str = 'Paolo Gollo&nbsp;<img src="/static/torta.gif" class="emoji-icon">'
                elif cell_str == 'Federico Cassola':
                    cell_str = 'Federico Cassola&nbsp;<img src="/static/pesca.gif" class="emoji-icon">'
                elif cell_str == 'Alfredo Crosetti':
                    cell_str = 'Alfredo Crosetti&nbsp;<img src="/static/soldi_che_volano.gif" class="emoji-icon">'
                elif cell_str == 'Dario Hourngir':
                    cell_str = 'Dario Hourngir&nbsp;<img src="/static/palla_tennis.gif" class="emoji-icon">'
                elif cell_str == 'Vito Danieli':
                    cell_str = 'Vito Danieli&nbsp;<img src="/static/mirror_ball.gif" class="emoji-icon">'
                elif cell_str == 'Mauro Damonte':
                    cell_str = 'Mauro Damonte&nbsp;<img src="/static/caffe.gif" class="emoji-icon">'
                elif cell_str == 'Marco Tizzi':
                    cell_str = 'Marco Tizzi&nbsp;<img src="/static/pugno_marco.gif" class="emoji-icon">'
                elif cell_str == 'Daniele Luppi':
                    cell_str = 'Daniele Luppi&nbsp;<img src="/static/drago.gif" class="emoji-icon">'
                elif cell_str == 'Veronica Bonati':
                    cell_str = 'Veronica Bonati&nbsp;<img src="/static/cocktail.gif" class="emoji-icon">'
                elif cell_str == 'Elisabetta Fiori':
                    cell_str = 'Elisabetta Fiori&nbsp;<img src="/static/tromba.gif" class="emoji-icon">'
                elif cell_str == 'Federico Buscemi':
                    cell_str = 'Federico Buscemi&nbsp;<img src="/static/maiale.gif" class="emoji-icon">'

            for col_idx in colonna_gialla:
                valore = df.iloc[r_idx, col_idx]
                if type(valore) == str:
                    if len(valore) == 1:
                        if valore in ['P', 'H', 'I']:
                            style_parts.append(
                                'color: #E50000; font-weight: bold;')
                            break

                    elif len(valore) > 1:
                        if valore == 'CP':
                            style_parts.append(
                                'color: #FF00FF; font-weight: normal;')
                            break

                        elif valore == 'SI':
                            style_parts.append(
                                'color: #E50000; font-weight: bold;')
                            break

                        elif valore.endswith('H'):
                            if valore[-2:] == 'CH':
                                style_parts.append(
                                    'color: #FF00FF; font-weight: normal;')
                                break

                            else:
                                style_parts.append(
                                    'color: #E50000; font-weight: bold;')
                                break

            ################################
            # Emoji in svg prese da qui https://github.com/adobe-fonts/noto-emoji-svg

            valore_cella = sheet.cell_value(r_idx, 0)
            # print(f'{valore_cella=} {r_idx=} {c_idx=}')
            # print(valore_cella)

            if (r_idx, c_idx) in celle_R:
                cell_str = '<img src="/static/radar.svg" class="emoji-icon">'
            elif (r_idx, c_idx) in celle_HR:
                cell_str = 'H<img src="/static/radar.svg" class="emoji-icon">'
                style_parts.append('color: #E50000; font-weight: bold;')
            elif (r_idx, c_idx) in celle_TR:
                cell_str = '<img src="/static/radar.svg" class="emoji-icon">T'
                style_parts.append('color: #2CA02C; font-weight: bold;')
            elif (r_idx, c_idx) in celle_PuntoR:
                cell_str = '<img src="/static/radar.svg" class="emoji-icon">•󠁏󠁏'
                style_parts.append('color: #1F77B4; font-weight: bold;')
            elif (r_idx, c_idx) in celle_MR:
                cell_str = '<img src="/static/radar.svg" class="emoji-icon"><img src="/static/razzo.gif" class="emoji-icon">'
                style_parts.append('color: #278C27; font-weight: bold;')
            elif (r_idx, c_idx) in celle_CPR:
                cell_str = '<img src="/static/uccellino.svg" class="emoji-icon"><img src="/static/radar.svg" class="emoji-icon">'
            elif (r_idx, c_idx) in celle_AP:
                cell_str = 'P<img src="/static/biberon.gif" class="emoji-icon">'
                style_parts.append('color: #E50000; font-weight: bold;')
            elif (r_idx, c_idx) in celle_ACP:
                cell_str = 'C<img src="/static/biberon.gif" class="emoji-icon">'
                style_parts.append('color: #E50000; font-weight: bold;')
            elif (r_idx, c_idx) in celle_Rep:
                cell_str = '<img src="/static/telefono.svg" class="emoji-icon">'
            elif (r_idx, c_idx) in celle_Mail_controllo:
                cell_str = '<img src="/static/buca_lettere.svg" class="emoji-icon">'
            elif (r_idx, c_idx) in celle_Mamma:
                cell_str = '<img src="/static/maternita.gif" class="emoji-icon">'

            if cell_str == 'TC':
                cell_html = (
                    '<span style="color:#000000; font-weight:bold;"><img src="/static/mondo.gif" class="emoji-icon"></span>'
                    '<span style="color:#2CA02C; font-weight:bold;">T</span>'
                )
            elif cell_str == 'FC':

                if 'Dicembre' in titolo or 'Gennaio' in titolo or 'Febbraio' in titolo:
                    a = '<span style="color:#000000; font-weight:bold;"><img src="/static/montagna.svg" class="emoji-icon"></span>'
                elif 'Giugno' in titolo or 'Luglio' in titolo or 'Agosto' in titolo:
                    a = '<span style="color:#000000; font-weight:bold;"><img src="/static/spiaggia.svg" class="emoji-icon"></span>'
                elif 'Marzo' in titolo or 'Aprile' in titolo or 'Maggio' in titolo:
                    a = '<span style="color:#000000; font-weight:bold;"><img src="/static/tenda.gif" class="emoji-icon"></span>'
                elif 'Settembre' in titolo or 'Ottobre' in titolo or 'Novembre' in titolo:
                    a = '<span style="color:#000000; font-weight:bold;"><img src="/static/funghi.gif" class="emoji-icon"></span>'
                cell_html = (
                    f"<span style='color:#000000; font-weight:bold;'><img src='/static/mondo.gif' class='emoji-icon'></span>{a}"
                )
                del (a)
            elif cell_str == 'MH' or cell_str == 'HM':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;">H</span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'CPN':
                cell_html = (
                    '<span style="color:#FF00FF; font-weight:bold;">CP<img src="/static/dormire.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'CPC':
                cell_html = (
                    '<span style="color:#FF00FF; font-weight:bold;">CP<img src="/static/mondo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'CPM':
                cell_html = (
                    '<span style="color:#FF00FF; font-weight:bold;">CP</span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'RON':
                cell_html = (
                    '<span style="color:#1F77B4; font-weight:bold;">RO<img src="/static/dormire.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'MI':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;">I</span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'XI' or cell_str == 'IX':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;">I</span>'
                    '<span style="color:#1F77B4; font-weight:bold;">✖</span>'
                )
            elif cell_str == 'PM':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;">P</span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'MN':
                cell_html = (
                    '<span style="color:#808080; font-weight:bold;"><img src="/static/dormire.gif" class="emoji-icon"></span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'IM':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;">I</span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'IT':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;">I</span>'
                    '<span style="color:#800000; font-weight:bold;">Tec</span>'
                )
            elif cell_str == 'MT':
                cell_html = (
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                    '<span style="color:#2CA02C; font-weight:bold;">T</span>'
                )
            elif cell_str == 'T.':
                cell_html = (
                    '<span style="color:#2CA02C; font-weight:bold;">T</span>'
                    '<span style="color:#1F77B4; font-weight:bold;">•󠁏󠁏</span>'
                )
            elif cell_str == 'X.':
                cell_html = (
                    '<span style="color:#1F77B4; font-weight:bold;">✖</span>'
                    '<span style="color:#1F77B4; font-weight:bold;">•󠁏󠁏</span>'
                )
            elif cell_str == 'MCH' or cell_str == 'CHM':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;"><img src="/static/pulcino.svg" class="emoji-icon"></span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'MC' or cell_str == 'CM':
                cell_html = (
                    '<span style="color:#000000; font-weight:bold;"><img src="/static/mondo.gif" class="emoji-icon"></span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'CP3':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;"><img src="/static/pulcino.svg" class="emoji-icon"></span>'
                    '<span style="color:#E50000; font-weight:bold;"><img src="/static/tv.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'M3':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;"><img src="/static/tv.gif" class="emoji-icon"></span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'P3':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;">P</span>'
                    '<span style="color:#E50000; font-weight:bold;"><img src="/static/tv.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'PT':
                cell_html = (
                    # '<span>🤬</span>'
                    '<span><img src="/static/arrabiato.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'TCH':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;"><img src="/static/pulcino.svg" class="emoji-icon"></span>'
                    '<span style="color:#2CA02C; font-weight:bold;">T</span>'
                )
            elif cell_str == 'MSH':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;"><img src="/static/SH.gif" class="emoji-icon"></span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/razzo.gif" class="emoji-icon"></span>'
                )
            elif cell_str == 'MIJ':
                cell_html = (
                    '<span style="color:#FF00FF; font-weight:bold;">I<img src="/static/biberon.gif" class="emoji-icon"></span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/bimbo.svg" class="emoji-icon"></span>'
                )
            elif cell_str == 'HJ':
                cell_html = (
                    '<span style="color:#E50000; font-weight:bold;">H</span>'
                    '<span style="color:#278C27; font-weight:bold;"><img src="/static/bimbo.svg" class="emoji-icon"></span>'
                )
            elif cell_str == 'M':
                # cell_html = '🚀'
                cell_html = '<img src="/static/razzo.gif" class="emoji-icon">'
                style_parts.append('color: #278C27; font-weight: bold;')
            elif cell_str == '3' and r_idx > 4:
                # cell_html = '📺'
                cell_html = '<img src="/static/tv.gif" class="emoji-icon">'
            elif cell_str == 'N':
                style_parts.append('color: #808080; font-weight: bold;')
                # cell_html = '💤'
                cell_html = '<img src="/static/dormire.gif" class="emoji-icon">'
            elif cell_str == 'C':
                # cell_html = '🌍'
                cell_html = '<img src="/static/mondo.gif" class="emoji-icon">'
                style_parts.append('color: #000000; font-weight: bold;')
            elif cell_str == 'T':
                cell_html = 'T'
                style_parts.append('color: #2CA02C; font-weight: bold;')
            elif cell_str == 'X':
                cell_html = '✖'
                style_parts.append('color: #1F77B4; font-weight: bold;')
            elif cell_str == 'RO':
                cell_html = 'RO'
                style_parts.append('color: #1F77B4; font-weight: bold;')
            elif cell_str == 'RP':
                cell_html = 'RP'
                style_parts.append('color: #1F77B4; font-weight: bold;')
            elif cell_str == 'RD':
                # cell_html = '⛪'
                cell_html = '<img src="/static/domenica.svg" class="emoji-icon">'
                style_parts.append('color: #1F77B4; font-weight: bold;')
            elif cell_str == 'F':
                if 'Dicembre' in titolo or 'Gennaio' in titolo or 'Febbraio' in titolo:
                    # cell_html = '🏔️'
                    cell_html = '<img src="/static/montagna.svg" class="emoji-icon">'
                elif 'Giugno' in titolo or 'Luglio' in titolo or 'Agosto' in titolo:
                    # cell_html = '🏖️'
                    cell_html = '<img src="/static/spiaggia.svg" class="emoji-icon">'
                elif 'Marzo' in titolo or 'Aprile' in titolo or 'Maggio' in titolo:
                    # cell_html = '🏕️'
                    cell_html = '<img src="/static/tenda.gif" class="emoji-icon">'
                elif 'Settembre' in titolo or 'Ottobre' in titolo or 'Novembre' in titolo:
                    # cell_html = '🍄'
                    cell_html = '<img src="/static/funghi.gif" class="emoji-icon">'
                style_parts.append('color: #1F77B4; font-weight: bold;')
            elif cell_str == 'π':
                style_parts.append('color: #E50000; font-weight: bold;')
                cell_html = 'π'
            elif cell_str == 'PJ':
                style_parts.append('color: #E50000; font-weight: bold;')
                cell_html = 'P<img src="/static/biberon.gif" class="emoji-icon">'
            elif cell_str == 'SI':
                cell_html = 'SI'
                style_parts.append('color: #E50000; font-weight: bold;')
            elif cell_str == 'SH':
                cell_html = 'SH'
                cell_html = '<img src="/static/SH.gif" class="emoji-icon">'
                style_parts.append('color: #E50000; font-weight: bold;')
            elif cell_str == 'P':
                cell_html = 'P'
                style_parts.append('color: #E50000; font-weight: bold;')
            elif cell_str == 'H':
                cell_html = 'H'
                style_parts.append('color: #E50000; font-weight: bold;')
            elif cell_str == 'I':
                cell_html = 'I'
                style_parts.append('color: #E50000; font-weight: bold;')
            elif cell_str == 'IJ':
                cell_html = cell_str
                style_parts.append('color: #FF00FF; font-weight: bold;')
            elif cell_str == 'CH':
                # cell_html = '🐤'
                cell_html = '<img src="/static/pulcino.svg" class="emoji-icon">'
                style_parts.append('color: #FF00FF; font-weight: bold;')
            elif cell_str == 'CP':
                # cell_html = '🐦'
                cell_html = '<img src="/static/uccellino.svg" class="emoji-icon">'
                style_parts.append('color: #FF00FF; font-weight: bold;')
            elif cell_str == '.':
                cell_html = '󠁯•󠁏󠁏'
                style_parts.append('color: #1F77B4; font-weight: bold;')
            elif cell_str == '..':
                cell_html = '󠁯•󠁏󠁏󠁯•󠁏󠁏'
                style_parts.append('color: #1F77B4; font-weight: bold;')
            elif cell_str == 'Ì':
                # cell_html = '🚑'
                cell_html = '<img src="/static/ambulanza.svg" class="emoji-icon">'
                style_parts.append('color: #E50000; font-weight: bold;')
            else:
                cell_html = cell_str

            # TODO Non capisco perché non funzioni...
            # if valore_cella == 'Davide Sacchetti' and cell_html == 'T':
            #     print(f'T, {valore_cella=} {r_idx=} {c_idx=}')
            #     cell_html = '<img src="/static/davide_smart.svg" class="emoji-icon">'

            stile = f' style="{" ".join(style_parts)}"' if style_parts else ""
            html_riga += f'<td{stile}{tooltip_attr}>{cell_html}</td>'

        html_riga += '</tr>\n'
        # print(f'{html_riga=}')

        try:
            if r_idx == 1:
                riga_numeri_giorni = html_riga

            if r_idx == 2:
                riga_nomi_giorni = html_riga
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

    # Inserisce righe con 'C' dopo la terza riga vuota
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

    # Se ci sono meno di 3 spazi, inseriscile alla fine
    if not inserito:
        corpo_tabella.extend(righe_con_C)
        corpo_tabella.append(riga_spazio)

    html_tabella += '  <tbody>\n' + \
        ''.join(corpo_tabella) + '  </tbody>\n</table>'

    # Rimuove righe vuote doppie consecutive
    html_tabella = re.sub(
        r'(<tr>\s*<td colspan="\d+" style="border:none; height: ?\d+px;"></td>\s*</tr>\s*){2,}',
        r'\1',
        html_tabella,
        flags=re.IGNORECASE
    )

    # print(f'{html_tabella=}')

    # Inserisce la tabella nell'HTML base
    percorso_index = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "index.html")
    with open(percorso_index, encoding="utf-8") as f:
        html_base = f.read()

    link_css = '<link rel="stylesheet" type="text/css" href="/static/stile.css">'
    html_base = html_base.replace("</head>", f"  {link_css}\n</head>")
    html_completo = html_base.replace(
        "<!-- La tabella verrà messa qui -->", html_tabella)
    # print(f'\n\n\n{html_completo}')

    return Response(html_completo, mimetype="text/html")


if __name__ == '__main__':
    app.run(debug=True)
