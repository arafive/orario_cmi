README.txt

:Author: daniele.carnevale
:Email: daniele.carnevale@01588-lenovo.cfmi.arpal.org
:Date: 2025-08-03 08:48

Avviare flask in background:
nohup flask run --host=0.0.0.0 --port=5000 > log_web.log 2>&1 &

ip di questo computer ARPAL: 10.24.50.225 (trovato con >>> hostname -I)
Link: http://10.24.50.225:5000

>>> Cosa ho fatto per far vedere il link anche con la VPN <<<
sudo systemctl status firewalld
sudo firewall-cmd --permanent --add-port=5000/tcp
sudo firewall-cmd --reload
sudo firewall-cmd --list-ports

Maggiori dettagli su questa chat: https://chatgpt.com/c/688f02a0-cebc-8326-8022-c757edab1687
