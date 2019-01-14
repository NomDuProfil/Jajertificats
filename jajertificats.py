# -*- coding: utf-8 -*-

import socket
import ssl
import datetime
import OpenSSL
import socket

FILE_DOMAINS = "domains.txt"
EXPIRATION_DAYS = 30

def sslExpirationDate(address):
	cert=ssl.get_server_certificate((address, 443))
	x509 = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, cert)

	return datetime.datetime.strptime(x509.get_notAfter().decode('ascii'), '%Y%m%d%H%M%SZ')

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
			print '\033[1;31m[!]\033[1;m'+" Certificat pour "+currentdomain+" expiré"
		elif result["status"] == "expiresoon":
			print '\033[1;33m[!]\033[1;m'+" Certificat pour "+currentdomain+" expire dans "+str(result["date"])+" jours"
		else:
			print '\033[1;34m[+]\033[1;m'+" Certificat pour "+currentdomain+" OK"
