import os

lista_mesi = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre']
# !!! Fino a Settembre 2008 erano in formato .doc

cartella_orario = './link'
cartella_orario_OVG = './link_OVG'

os.makedirs(cartella_orario, exist_ok=True)
os.makedirs(cartella_orario_OVG, exist_ok=True)

cartella_madre_orario = '/mnt/isilon-arpal/orariocmi'
cartella_madre_orario_OVG = '/mnt/isilon-arpal/agcmi00a00_comune/archivio/Archivio_Progetti/2025-16_Progetto Vigilanza Gialla/Orario'

...

for anno in range(2008, 2028):

    for mese in lista_mesi:

        origine_underscore = f'{cartella_madre_orario_OVG}/{anno}/{mese}_{anno}.pdf'
        origine_spazio = f'{cartella_madre_orario_OVG}/{anno}/{mese} {anno}.pdf'
        origine = origine_underscore if os.path.exists(origine_underscore) else origine_spazio

        destinazione = f'{cartella_orario_OVG}/{mese} {anno}.pdf'
        if os.path.lexists(destinazione):
            os.remove(destinazione)
        os.symlink(origine, destinazione)
