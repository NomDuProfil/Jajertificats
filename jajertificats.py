# -*- coding: utf-8 -*-

import socket
import ssl
import datetime
import OpenSSL
import socket
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import sys
import multiprocessing
from time import sleep
import re

DUREE_ENTRE_CHAQUE_REQUETE = 0 #En seconde
TIMEOUT_REQUEST = 5 #En seconde
EXPIRATION_DAYS = 30 #En nombre de jour
PERIODE_DE_VALIDITE = 825 #En nombre de jour

LIST_PORT = ['443', '8443']

def writeresult(result, pathsave="./"):
	wb = Workbook()
	ws = wb.active
	listwidthcolumn = [] 
	ws['A1'] = "Status certificat"
	ws['A1'].alignment = Alignment(horizontal='center')
	ws['A1'].font = Font(bold=True)
	ws.column_dimensions['A'].width = len("Status certificat")*1.3
	ws['B1'] = "Domaines"
	ws['B1'].alignment = Alignment(horizontal='center')
	ws['B1'].font = Font(bold=True)
	ws['C1'] = "Port"
	ws.column_dimensions['C'].width = len("Port")*1.3
	ws['C1'].alignment = Alignment(horizontal='center')
	ws['C1'].font = Font(bold=True)
	ws['D1'] = "Date de délivrance"
	ws.column_dimensions['D'].width = len("Date de délivrance")*1.3
	ws['D1'].alignment = Alignment(horizontal='center')
	ws['D1'].font = Font(bold=True)
	ws['E1'] = "Date d'expiration"
	ws.column_dimensions['E'].width = len("Date d'expiration")*1.3
	ws['E1'].alignment = Alignment(horizontal='center')
	ws['E1'].font = Font(bold=True)
	ws['F1'] = "Jours restants"
	ws.column_dimensions['F'].width = len("Jours restants")*1.3
	ws['F1'].alignment = Alignment(horizontal='center')
	ws['F1'].font = Font(bold=True)
	ws['G1'] = "Validité (nombre de jours)"
	ws.column_dimensions['G'].width = len("Validité (nombre de jours)")*1.3
	ws['G1'].alignment = Alignment(horizontal='center')
	ws['G1'].font = Font(bold=True)
	ws['H1'] = "Vérifié par"
	ws.column_dimensions['H'].width = len("Vérifié par")*1.3
	ws['H1'].alignment = Alignment(horizontal='center')
	ws['H1'].font = Font(bold=True)
	currentline = 2
	for current in result:
		if current["status"] == "ok":
			ws['A'+str(currentline)] = "OK"
			ws['A'+str(currentline)].fill = PatternFill(fill_type="solid", start_color='009933', end_color='009933')
		elif current["status"] == "expired":
			ws['A'+str(currentline)] = "EXPIRÉ"
			ws['A'+str(currentline)].fill = PatternFill(fill_type="solid", start_color='FF0000', end_color='FF0000')
		elif current["status"] == "expiresoon":
			ws['A'+str(currentline)] = "ÉXPIRE BIENTOT"
			ws['A'+str(currentline)].fill = PatternFill(fill_type="solid", start_color='FF9900', end_color='FF9900')
		elif current["status"] == "validitytoolong":
			if ws.column_dimensions['A'].width < len("VALIDITE TROP LONGUE")*1.3:
				ws.column_dimensions['A'].width = len("VALIDITE TROP LONGUE")*1.3
			ws['A'+str(currentline)] = "VALIDITÉ TROP LONGUE"
			ws['A'+str(currentline)].fill = PatternFill(fill_type="solid", start_color='FF9900', end_color='FF9900')
		elif current["status"] == "NOT_MATCH":
			if ws.column_dimensions['A'].width < len("HOSTNAME INVALIDE")*1.3:
				ws.column_dimensions['A'].width = len("HOSTNAME INVALIDE")*1.3
			ws['A'+str(currentline)] = "HOSTNAME INVALIDE"
			ws['A'+str(currentline)].fill = PatternFill(fill_type="solid", start_color='FF0000', end_color='FF0000')
		else:
			ws['A'+str(currentline)] = current["status"]
			ws['A'+str(currentline)].fill = PatternFill(fill_type="solid", start_color='FF0000', end_color='FF0000')
		ws['A'+str(currentline)].font = Font(color='FFFFFF')
		ws['A'+str(currentline)].alignment = Alignment(horizontal='center')
		ws['B'+str(currentline)] = current["domain"]
		if ws.column_dimensions['B'].width < len(current["domain"])*1.3:
			ws.column_dimensions['B'].width = len(current["domain"])*1.3
		ws['C'+str(currentline)] = current["port"]
		if ws.column_dimensions['C'].width < len(current["port"])*1.3:
			ws.column_dimensions['C'].width = len(current["port"])*1.3
		if (str(current["notBefore"]) != "ERROR"):
			ws['D'+str(currentline)] = str(current["notBefore"].year)+'-'+str(current["notBefore"].month)+'-'+str(current["notBefore"].day)
		else:
			ws['D'+str(currentline)] = "ERROR"
		if (str(current["notAfter"]) != "ERROR"):
			ws['E'+str(currentline)] = str(current["notAfter"].year)+'-'+str(current["notAfter"].month)+'-'+str(current["notAfter"].day)
		else:
			ws['E'+str(currentline)] = "ERROR"
		ws['F'+str(currentline)] = str(current["deltaToday"])
		ws['G'+str(currentline)] = str(current["periodevalidity"])
		ws['H'+str(currentline)] = current["deliver"]
		if ws.column_dimensions['H'].width < len(current["deliver"])*1.3:
			ws.column_dimensions['H'].width = len(current["deliver"])*1.3
		currentline+=1
	wb.save(pathsave+'StatusCertificates.xlsx')

def sslExpirationDate(address, port):
	if DUREE_ENTRE_CHAQUE_REQUETE != 0:
		sleep(DUREE_ENTRE_CHAQUE_REQUETE)
	try:
		cert=ssl.get_server_certificate((address, int(port)), ssl_version=ssl.PROTOCOL_TLSv1_2)
		certLoad = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, cert)
		certinfo = {}
		certinfo["notBefore"] = datetime.datetime.strptime(certLoad.get_notBefore().decode('ascii'), '%Y%m%d%H%M%SZ')
		certinfo["notAfter"] = datetime.datetime.strptime(certLoad.get_notAfter().decode('ascii'), '%Y%m%d%H%M%SZ')
		if certLoad.get_issuer().CN != None:
			certinfo["deliver"] = certLoad.get_issuer().CN
		elif certLoad.get_issuer().O != None:
			certinfo["deliver"] = certLoad.get_issuer().O
		else:
			certinfo["deliver"] = "Introuvable"
		return certinfo
	except ssl.SSLError as err:
		if ("SSLV3_ALERT" in str(err) or "WRONG_SIGNATURE_TYPE" in str(err) or "TLSV1_ALERT_INTERNAL_ERROR" in str(err)):
			try:
				context = ssl.create_default_context()
				conn = context.wrap_socket(socket.socket(socket.AF_INET), server_hostname=address)
				conn.connect((address, int(port)))
				cert = conn.getpeercert()
				certinfo = {}
				certinfo["notBefore"] = datetime.datetime.strptime(cert["notBefore"].decode('ascii'), '%b %d %H:%M:%S %Y %Z')
				certinfo["notAfter"] = datetime.datetime.strptime(cert["notAfter"].decode('ascii'), '%b %d %H:%M:%S %Y %Z')
				try:
					certinfo["deliver"] = cert["issuer"][3][0][1]
				except IndexError:
					certinfo["deliver"] = cert["issuer"][2][0][1]
				conn.close()
				return certinfo
			except socket.error as err:
				#print str(err)
				return "ERROR"
			except ssl.CertificateError as err:
				#print str(err)
				if ("doesn't match either of" in str(err)):
					return "NOT_MATCH"
				return "ERROR"
		elif ("WRONG_SSL_VERSION" in str(err) or "UNSUPPORTED_PROTOCOL" in str(err)):
			try:
				cert=ssl.get_server_certificate((address, int(port)), ssl_version=ssl.PROTOCOL_TLSv1)
				certLoad = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, cert)
				certinfo = {}
				certinfo["notBefore"] = datetime.datetime.strptime(certLoad.get_notBefore().decode('ascii'), '%Y%m%d%H%M%SZ')
				certinfo["notAfter"] = datetime.datetime.strptime(certLoad.get_notAfter().decode('ascii'), '%Y%m%d%H%M%SZ')
				if certLoad.get_issuer().CN != None:
					certinfo["deliver"] = certLoad.get_issuer().CN
				elif certLoad.get_issuer().O != None:
					certinfo["deliver"] = certLoad.get_issuer().O
				else:
					certinfo["deliver"] = "Introuvable"
				return certinfo
			except ssl.SSLError as err:
				#print str(err)
				return "ERROR"
			except socket.error as err:
				#print str(err)
				return "ERROR"
		#print str(err)
		return "ERROR"
	except socket.error as err:
		if "Name or service not known" in str(err):
			return "ERR RESOLUTION"
		if "Connection refused" in str(err):
			return "CONN REFUSED"
		if "[Errno 0]" in str(err):
			return "INVALID CERT"
		return "ERROR"

def sslExpirationCalcul(address, port):
	expirationdate = sslExpirationDate(address, port)
	if type(expirationdate) != str:
		certcalcul = {}
		certcalcul["notAfter"] = expirationdate["notAfter"]
		certcalcul["notBefore"] = expirationdate["notBefore"]
		certcalcul["deliver"] = expirationdate["deliver"]
		certcalcul["deltaToday"] = expirationdate["notAfter"] - datetime.datetime.utcnow()
		certcalcul["deltaValidity"] = expirationdate["notAfter"] - expirationdate["notBefore"]
		return certcalcul
	else:
		return expirationdate

def sslExpirationStatus(address, port, days_check, validity_period, return_value):
	getdate = sslExpirationCalcul(address, port)
	result = {}
	result["domain"] = address
	result["port"] = port
	if type(getdate) == str:
		result["status"] = getdate
		result["periodevalidity"] = "ERROR"
		result["deltaToday"] = "ERROR"
		result["deliver"] = "ERROR"
		result["notAfter"] = "ERROR"
		result["notBefore"] = "ERROR"
		return_value[0] = result
	else:
		result["notAfter"] = getdate["notAfter"]
		result["notBefore"] = getdate["notBefore"]
		result["deliver"] = getdate["deliver"]
		result["deltaToday"] = getdate["deltaToday"].days
		result["periodevalidity"] = getdate["deltaValidity"].days
		if getdate["deltaToday"] <= datetime.timedelta(days=0):
			result["status"] = "expired"
		elif getdate["deltaToday"] < datetime.timedelta(days=days_check):
			result["status"] = "expiresoon"
		else:
			if getdate["deltaValidity"] > datetime.timedelta(days=validity_period):
				result["status"] = "validitytoolong"
			else:
				result["status"] = "ok"
		return_value[0] = result

if len(sys.argv) < 2:
	print "usage: python jajertificats.py PATH_FICHIER_LISTE_DOMAINE"
	sys.exit(1)
allresult = []
with open(sys.argv[1], "r") as f:
    for currentdomain in f:
		currentdomain = currentdomain.replace('\n', '')
		for currentport in LIST_PORT:
			domaininprogress = currentdomain
			manager = multiprocessing.Manager()
			resultthread = manager.dict()
			p = multiprocessing.Process(target=sslExpirationStatus, args=(domaininprogress, currentport, EXPIRATION_DAYS, PERIODE_DE_VALIDITE, resultthread))
			p.start()
			p.join(TIMEOUT_REQUEST)
			if not p.is_alive():
				result = resultthread[0]
				if result["status"] == "ERR RESOLUTION":
					print u'\033[1;31m[ERR RÉSOLUTION]\033[1;m'+" Impossible de joindre "+domaininprogress.encode('utf8')
				elif result["status"] == "CONN REFUSED":
					print u'\033[1;31m[CONN REFUSÉE]\033[1;m'+u" Connexion refusée pour "+domaininprogress.encode('utf8')+" sur le port "+currentport
				elif result["status"] == "INVALID CERT":
					print '\033[1;31m[CERT INVALIDE]\033[1;m'+u" Certificat invalide pour "+domaininprogress.encode('utf8')+" sur le port "+currentport
				elif result["status"] == "ERROR":
					print '\033[1;31m[ERREUR]\033[1;m'+" Impossible de recuperer le certificat pour "+domaininprogress.encode('utf8')+" sur le port "+currentport
				elif result["status"] == "expired":
					print u'\033[1;31m[EXPIRÉ]\033[1;m'+" Certificat pour "+domaininprogress.encode('utf8')+u" expiré depuis "+str(result["deltaToday"]*-1)+" jours sur le port "+currentport
				elif result["status"] == "expiresoon":
					print u'\033[1;33m[EXPIRE BIENTÔT]\033[1;m'+" Certificat pour "+domaininprogress.encode('utf8')+" expire dans "+str(result["deltaToday"])+" jours sur le port "+currentport
				elif result["status"] == "validitytoolong":
					print u'\033[1;33m[VALIDITÉ TROP LONGUE]\033[1;m'+" Certificat pour "+domaininprogress.encode('utf8')+u" a une période de validité de "+str(result["periodevalidity"])+" jours sur le port "+currentport
				else:
					print '\033[1;34m[OK]\033[1;m'+" Certificat pour "+domaininprogress.encode('utf8')+" OK sur le port "+currentport
			else:
				p.terminate()
				result = {}
				result["domain"] = domaininprogress
				result["status"] = "TIMEOUT"
				result["port"] = currentport
				result["deliver"] = "ERROR"
				result["deltaToday"] = "ERROR"
				result["notAfter"] = "ERROR"
				result["notBefore"] = "ERROR"
				result["periodevalidity"] = "ERROR"
				print '\033[1;31m[TIMEOUT]\033[1;m'+" TIMEOUT pour "+domaininprogress.encode('utf8')+" sur le port "+currentport
			allresult.append(result)

print "\nSauvegarde des resultats en cours..."
writeresult(allresult)
print "Resultats disponible dans le fichier StatusCertificates.xlsx"
