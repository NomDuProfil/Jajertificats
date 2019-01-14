# -*- coding: utf-8 -*-

import socket
import ssl
import datetime
from termcolor import colored

FILE_DOMAINS = "domains.txt"
EXPIRATION_DAYS = 256

def sslExpirationDate(address):
    format_date = r'%b %d %H:%M:%S %Y %Z'

    lebeaucontext = ssl.create_default_context()
    labelconnexion = lebeaucontext.wrap_socket(
        socket.socket(socket.AF_INET),
        server_hostname=address,
    )
    labelconnexion.settimeout(5.0)
    labelconnexion.connect((address, 443))
    ssl_information = labelconnexion.getpeercert()
    return datetime.datetime.strptime(ssl_information['notAfter'], format_date)

def sslExpirationCalcul(address):
	expirationdate = sslExpirationDate(address)
	return expirationdate - datetime.datetime.utcnow()

def sslExpirationStatus(address, days_check):
	getdate = sslExpirationCalcul(address)
	result = {}
	result["date"] = getdate.days
	if getdate <= datetime.timedelta(days=0):
		result["status"] = "expired"
	elif getdate < datetime.timedelta(days=days_check):
		result["status"] = "expiresoon"
	else:
		result["status"] = "ok"
	return result

with open(FILE_DOMAINS, "r") as f:
    for currentdomain in f:
		currentdomain = currentdomain.replace('\n', '')
		result = sslExpirationStatus(currentdomain, EXPIRATION_DAYS)
		if result["status"] == "expired":
			print '\033[1;31m[!]\033[1;m'+" Certificat pour "+currentdomain+" expiré depuis "+str(result["date"])+" jours"
		elif result["status"] == "expiresoon":
			print '\033[1;33m[!]\033[1;m'+" Certificat pour "+currentdomain+" expire dans "+str(result["date"])+" jours"
		else:
			print '\033[1;34m[+]\033[1;m'+" Certificat pour "+currentdomain+" OK"
