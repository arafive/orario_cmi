
import pandas as pd

import requests
from requests.auth import HTTPBasicAuth

# Tutto l'archivio degli orari sta in /mnt/isilon-arpal/orariocmi
url = 'https://cmi-servizi.arpal.liguria.it/orario/luglio25.htm'

# Credenziali
username = "cmi"
password = "cmi"

# Richiesta con autenticazione base
response = requests.get(url, auth=HTTPBasicAuth(username, password))

tables = pd.read_html(response.text)
