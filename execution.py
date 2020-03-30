#!C:/python27/python
# -*-coding:utf-8 -*

#-------------------------------------------------------------------------
# SCRIPT D'EXECUTION - COMPOSITION DES MODELES AFP
# DEVELOPPE PAR : Bastien FERRAGU
# V1.0.0 - 2015-10-12 
# V1.1.0 - adapatation specifique CRT 
# V1.3.0 - 2019-09-XX - Ajout de la fillière Marketing
# V1.4.0 - 2020-01-03 - Optimisation 
#-------------------------------------------------------------------------

import os, sys, subprocess, shutil, glob, time
import shlex
import codecs
import base64
from subprocess	import PIPE, Popen
from time 		import gmtime
from functions	import *
import tarfile, gzip
import xml.etree.ElementTree as ET
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

ROOT_PATH 		= sys.argv[4]
cheminConfig 	= ROOT_PATH + os.sep +"param"

BUFSIZE = 4096
BOMLEN = len(codecs.BOM_UTF8)

""" CLASSES """

class LogFile:
	def __init__(self, nomlog, exitFile):
		self.__nomlog__ 	= nomlog
		self.__fichierlog__	= None
		self.__erreur__		= None
		self.exitFile 		= exitFile
		try:
			self.__fichierlog__ = open(self.__nomlog__, "w", 0)
		except Exception, e:
			self.__erreur__ == "Probleme a l'ouverture du fichier de log " + self.__nomlog__ + " : [" + str(e) + "]"
			return
		self.write("-- CREATION DU FICHIER --", 0)
		return	
	def write(self, message, exitCode):
		logtype, errorcode  = self.set_erreurType(exitCode)
		self.__fichierlog__.write(logtype + maintenant() + ' | ' + message + '\n')
		affichedebug(message, DEBUG)
		if (exitCode < 100):
			self.__fichierlog__.flush
		else:
			""" Fonction d'ecriture d'un message d'erreur dans la trace et gestion de fin du script """
			self.__fichierlog__.write(logtype + maintenant() + ' | ERRORCODE : ' + errorcode + ' | EXITCODE : ' + str(exitCode) + '\n')
			self.__fichierlog__.flush
			affichedebug("Exit code : " + str(exitCode), DEBUG)
			self.exitFile.exit(exitCode)	
			sys.exit(exitCode)		
		return
	def set_erreurType(self, exitCode):
		""" Definit le type du message à afficher """
		if (exitCode==0):
			logtype = "I | "
			return logtype, "0"
		else:
			logtype = "E | "
			errorcode = str(exitCode)
			return logtype, errorcode
	def get_erreur(self):
		""" Fonction de renvoi d'erreur liee a la gestion de la trace """
		return self.__erreur__
	def close(self):
		self.__fichierlog__.close()
		return

class Env:
	def __init__(self, nomFichier, logfile):
		self.__nomFichier__ = nomFichier
		self.__dico__       = {}
		self.__erreur__		= None
		self.logfile 		= logfile
		if not os.path.isfile(self.__nomFichier__):
			self.logfile.write( "Absence du fichier d'environnement : '" + nomFichier, 100 )
		self.logfile.write( "Chargement du fichier '" + nomFichier + "' : OK", 0 )

		self.load()
		return
	def load(self):
		""" Fonction de chargement du fichier de table """
		try:
			ficTMP = file(self.__nomFichier__, "rU")
			for ligne in ficTMP:
				ligne = ligne.replace("\n", "")
				if (len(ligne) < 1):
					break
				listeTMP = ligne.strip().split("\t")
				self.__dico__[ listeTMP[0] ] = listeTMP[1]
			ficTMP.close()
		except Exception, e:
			self.logfile.write( "Erreur detectee au chargement de '" + self.__nomFichier__ + " -- " + str(e) , 100 )
		return
	def get_erreur(self):
		""" Fonction de renvoi d'erreur """
		return self.__erreur__
	def get_element(self, cle):
		""" Fonction de recherche d'une entree dans la table de configuration de l'environnement """
		try:
			element = self.__dico__[cle]
		except Exception, e:
			self.logfile.write( "Pas d'entree dans la table " + self.__nomFichier__ + " pour la cle '" + cle + "'", 100 )
			element = "erreur"
		return element

class Type:
	def __init__(self, nomFichier, log, longueur):
		""" Classe de gestion de la table de configuration des types de traitement """
		self.__nomFichier__ = nomFichier
		self.__log__        = log
		self.__dico__       = {}
		self.__erreur__     = '-'
		self.__longLigne__  = longueur
		if not os.path.isfile(self.__nomFichier__):
			self.logfile.write( "Absence du fichier d'environnement : '" + nomFichier, 100 )
		self.logfile.write( "Chargement du fichier '" + nomFichier + "' : OK", 0 )
		self.load()
		return
		#
	def get_erreur(self):
		""" Fonction de renvoi d'erreur """
		return self.__erreur__
		#
	def load(self):
		""" Fonction de chargement du fichier de table """
		try:
			ficTMP = file(self.__nomFichier__, "rU")
			for ligne in ficTMP:
				ligne = ligne.replace("\n", "")
				if (len(ligne) < 1):
					break
				listeTMP = ligne.strip().split("\t")
				if ( len(listeTMP) != self.__longLigne__ ):
					raise( Exception("Probleme de structure, la ligne '" + str(listeTMP) + "' ne contient pas le nombre d'elements '" + str(self.__longLigne__) + "'") )
				self.__dico__[ listeTMP[0] ] = listeTMP
			ficTMP.close()
		except Exception, e:
			self.logfile.write( "Erreur detectee au chargement de '" + self.__nomFichier__ + " -- " + str(e) , 100 )
		return
		#
	def get_ligne(self, cle):
		""" Fonction de recherche d'une entree dans la table de configuration des types de traitement """
		try:
			element = self.__dico__[cle]
		except Exception, e:
			self.logfile.write( "Pas d'entree dans la table des 'Types' pour la cle '" + cle + "'", 100 )
			element = "erreur"
		return element

#		
class ExitFile():
	def __init__(self, fileName):
		self.fileName = open(fileName, "w")

	def exit(self, returncode):
		''' ecrit le return code dans un fichier temp '''
		self.fileName.write(str(returncode))
		self.fileName.flush()
		self.fileName.close()

class MiddleOffice():
	def __init__(self, dataFile, nomModele, confEnv, confProcess, confAppli):
		self.typeAssemblage		= "MO"
		self.__datafile__			= dataFile
		self.__nomModele__			= nomModele
		self.__logfile__			= dataFile + "_execution.log"
		self.__XMLFile__    		= os.path.basename(dataFile)
		self.__XMLFilesplitext__ = os.path.basename(os.path.splitext(dataFile)[0])
		self.__XMLFileOri__    		= (''.join(self.__XMLFile__.split('.')[2:-1]))+'.'+''.join(self.__XMLFile__.split('.')[-1:])
		self.__XMLFileOrisplitext__ = os.path.basename(os.path.splitext(self.__XMLFileOri__)[0])
		self.__workingkDir__		= self.__datafile__.split(self.__XMLFile__)[0] #contient le slash
		self.__workingkTrtFlux__	= self.__workingkDir__ + self.__XMLFile__ 
		self.__workingkDirPDF__ 	= self.__workingkTrtFlux__ + os.sep + "pdfTmp"
		self.__workingkDirPDFDuplicata__ 	= self.__workingkTrtFlux__ + os.sep + "pdfTmp"
		self.__workingkDirAFP__ 	= self.__workingkTrtFlux__ + os.sep + "afpTmp"
		self.__workingkDirZIPTMP__ 	= self.__workingkTrtFlux__ + os.sep + "zip"
		self.__workingkDirVPF__ 	= self.__workingkTrtFlux__ + os.sep + "vpf"
		self.__workingkDirECLATEMENT__ 	= self.__workingkTrtFlux__ + os.sep +"eclatement"
		self.__workingkDirHTML__ 	= self.__workingkTrtFlux__ + os.sep +"html"
		self.fichierVPF 			= self.__XMLFilesplitext__ + ".vpf"
		self.fichierVPFind			= self.__XMLFilesplitext__ + ".vpf.ind"
		self.exitCode				= None
		self.pid					= None
		self.stdOut					= None
		self.stdErr					= None
		self.__commandline__		= None
		self.envMiddleOffice		= None
		self.__opWD__				= None
		self.__opInstallDir__		= None
		self.__infoAppli__			= None
		self.RepFluxIn				= None
		self.RepTrt					= None
		self.repFluxOut				= None
		self.RepRessources			= None		
		self.Suivi 					= str("")
		self.ENVOI_MAIL			= str("")
		self.exitfile				= ExitFile(self.__workingkDir__ + self.__XMLFile__ + ".lock")
		self.logfile				= LogFile(self.__logfile__, self.exitfile)
		self.nbRepPDFOutput			= ""
		self.Htmlsource			= ""
		self.Config         = ""
		self.IndexFile			= []
		try:
			if not os.path.isdir(self.__workingkDirAFP__):
				os.makedirs(self.__workingkDirAFP__, mode=0777)
			self.logfile.write("Creation du repertoire : " + self.__workingkDirAFP__, 0)
			if not os.path.isdir(self.__workingkDirPDF__):
				os.makedirs(self.__workingkDirPDF__, mode=0777)
			self.logfile.write("Creation du repertoire : " + self.__workingkDirPDF__, 0)
			if not os.path.isdir(self.__workingkDirVPF__):
				os.makedirs(self.__workingkDirVPF__, mode=0777)
			self.logfile.write("Creation du repertoire : " + self.__workingkDirVPF__, 0)
			if not os.path.isdir(self.__workingkDirECLATEMENT__):
				os.makedirs(self.__workingkDirECLATEMENT__, mode=0777)
			self.logfile.write("Creation du repertoire : " + self.__workingkDirECLATEMENT__, 0)
			if not os.path.isdir(self.__workingkDirZIPTMP__):
				os.makedirs(self.__workingkDirZIPTMP__, mode=0777)
			self.logfile.write("Creation du repertoire : " + self.__workingkDirZIPTMP__, 0)
			if not os.path.isdir(self.__workingkDirHTML__):
				os.makedirs(self.__workingkDirHTML__, mode=0777)
			self.logfile.write("Creation du repertoire : " + self.__workingkDirHTML__, 0)
		except Exception, e:
			self.logfile.write("Erreur lors de la creation de l'arborescence : " + str(e), 101 )
		''' PARAMETRAGE '''
		self.logfile.write("parametrage execution",0)
		self.setConfiguration(confEnv, confProcess,confAppli)
		self.logfile.write("Fin Initialisation", 0)
		self.RepRessources	= self.repTrt + os.sep +"../ressources"
		#
		self.logfile.write("init env MiddleOffice ",0)
		self.__workingkDirIN__ 		= self.__workingkTrtFlux__ + os.sep +"entree"
		self.__workingkDirOUT__ 	= self.__workingkTrtFlux__ + os.sep +"sortie"
		self.__workingkDirTRC__ 	= self.__workingkTrtFlux__ + os.sep +"trace"
		self.__workingkDirPARAM__ 	= self.__workingkTrtFlux__ + os.sep +"parametres"
		self.__setWorkingDir__()
		try :
			filecopy = dataFile.replace(self.__workingkTrtFlux__,self.__workingkDirIN__ )
			if (os.path.isfile(filecopy)):
				self.logfile.write("suppression de %s" % filecopy ,0)
				os.remove(filecopy)
			self.logfile.write("copie de %s vers %s" % (dataFile,filecopy ),0)
			shutil.copy(dataFile, filecopy)
			# shutil.copy(confAppli, filecopy)
			pass
		except Exception,e: 
			self.logfile.write("probleme init env MO %s" % e,0)
		return
		
	def setCommandCompositionVPF(self):
		self.logfile.write("setCommandCompositionVPF", 0)
		self.logfile.write("XMLFilesplitext"+self.__XMLFilesplitext__,0)
		datafile = self.__workingkDirIN__ + os.sep + self.__XMLFile__
		self.__commandline__ = self.__opInstallDir__ + os.sep + "bin" +os.sep + "pydlexec" + \
			" " + self.__infoAppli__["dataloader"] + "_py" + \
			" -i " + datafile + \
			" -o " + self.__workingkDirVPF__ + os.sep + self.__XMLFilesplitext__ + ".vpf" + \
			" -E opWD=" + self.__opWD__ + \
			" -E opFam=" + self.__infoAppli__["opFam"]  + \
			" -E opAppli=" + self.__infoAppli__["opAppli"] + \
			" -E opInstallDir=" + self.__opInstallDir__ + \
			" -E opMaxWarning=0 -E opTecMacroTab=macros -V withVpfID=true -V wizard_action=sgml -V STD5_VPF_EMBED_IMG=YES -V STD5_WITHDATA=YES -V STD5_OUTMODE=VPF -E opEuro=WINDOWS" + \
			" -V STD5_TEMPLATE=" + self.__opWD__ + os.sep + self.__infoAppli__["opFam"] + os.sep + self.__infoAppli__["opAppli"] + os.sep +"template"+ os.sep + self.__infoAppli__["resdescid"] + ".xml" + \
			" -E opDoubleByte=1"
		return

	def setCommandCompositionVPFChorus(self):
		self.logfile.write("setCommandCompositionVPF", 0)
		self.logfile.write("XMLFilesplitext"+self.__XMLFilesplitext__,0)
		datafile = self.__workingkDirIN__ + os.sep + self.__XMLFile__
		self.__commandline__ = self.__opInstallDir__ + os.sep + "bin" +os.sep + "pydlexec" + \
			" " + self.__infoAppli__["dataloader"] + "_py" + \
			" -i " + datafile + \
			" -o " + self.__workingkDirVPF__ + os.sep + self.__XMLFilesplitext__ + ".vpf" + \
			" -E opWD=" + self.__opWD__ + \
			" -E opFam=" + self.__infoAppli__["opFam"]  + \
			" -E opAppli=" + self.__infoAppli__["opAppli"] + \
			" -E opInstallDir=" + self.__opInstallDir__ + \
			" -V chemin_sortie=" + self.__workingkDirPDF__ + \
			" -E opMaxWarning=0 -E opTecMacroTab=macros -V withVpfID=true -V wizard_action=sgml -V STD5_VPF_EMBED_IMG=YES -V STD5_WITHDATA=YES -V STD5_OUTMODE=VPF -E opEuro=WINDOWS" + \
			" -V STD5_TEMPLATE=" + self.__opWD__ + os.sep + self.__infoAppli__["opFam"] + os.sep + self.__infoAppli__["opAppli"] + os.sep +"template"+ os.sep + self.__infoAppli__["resdescid"] + ".xml" + \
			" -E opDoubleByte=1"
		return
	
	def setCommandCompositionVPF_Marketing(self):
		self.logfile.write("setCommandCompositionVPF Marketing", 0)
		self.Htmlsource = self.envAppli.get_element("HTML")
		self.Config = self.envAppli.get_element("CONFIG")
		datafile = self.__workingkDirIN__ + os.sep + self.__XMLFile__
		try :
			self.__commandline__ = self.__opInstallDir__ + os.sep + "bin" +os.sep + "pydlexec" + \
				" " + self.__infoAppli__["dataloader"] + "_py" + \
				" -i " + datafile + \
				" -o " + self.__workingkDirVPF__ + os.sep + self.__XMLFilesplitext__ + ".vpf" + \
				" -E opWD=" + self.__opWD__ + \
				" -E opFam=" + self.__infoAppli__["opFam"]  + \
				" -E opAppli=" + self.__infoAppli__["opAppli"] + \
				" -E opInstallDir=" + self.__opInstallDir__ + \
				" -E opMaxWarning=0 -E opTecMacroTab=macros -V withVpfID=true -V wizard_action=sgml -V STD5_VPF_EMBED_IMG=YES -V STD5_WITHDATA=YES -V STD5_OUTMODE=VPF -E opEuro=WINDOWS" + \
				" -V STD5_TEMPLATE=" + self.__opWD__ + os.sep + self.__infoAppli__["opFam"] + os.sep + self.__infoAppli__["opAppli"] + os.sep +"template"+ os.sep + self.__infoAppli__["resdescid"] + ".xml" + \
				" -E opDoubleByte=1" + \
				" -V chemin_html=" + self.Htmlsource + \
				" -V html_sortie="+self.__workingkDirHTML__
		except Exception,e: 
			Environnement.logfile.write("PROB COMPO SPECIFIQUE:"+str(e), 0)
		return
	""" FONCTIONS COMMUNES """

	def exit(self, returncode):
		self.exitfile.exit(returncode)

	def execCommande(self, typeAssemblage, erreur):
		self.logfile.write(typeAssemblage + " : ligne de commande [" + self.__commandline__ + "]", 0 )
		dateDebutStep = datetime.now()
		self.exitCode = runCommand(self.__commandline__, self)
		if typeAssemblage == "RAR":
			# WINRAR renvoie exit code 1 si warning : TODO prevoir le cas exit code 1 sur l'appel de la commande 
			# if ( self.exitCode != 0 ) and ( self.exitCode != 1 ):
			if ( self.exitCode != 0 ) :	
				self.logfile.write("Erreur detectee : ExitCode[" + str(self.exitCode) + "]", erreur )
		else :
			if ( self.exitCode != 0 ):
				self.logfile.write("Erreur detectee : ExitCode[" + str(self.exitCode) + "]", erreur )
		self.logfile.write("Resultat Commande : ExitCode[" + str(self.exitCode) + "] / Duree[" + str(datetime.now() - dateDebutStep) + "]", 0 )
		self.logfile.write("---------------", 0)
		return
		
	def setConfiguration(self, confEnv, confProcess, confAppli):
		self.envAppli			= Env(confAppli, self.logfile)
		self.OutputDirName	= self.envAppli.get_element("OUTPUTDIR")
		self.OutputDirNameBis	= self.envAppli.get_element("OUTPUTDIRBIS")
		if not os.path.isdir(self.OutputDirName):
			os.makedirs(self.OutputDirName, mode=0777)
			self.logfile.write("Creation du repertoire : " + self.OutputDirName, 0)
		if (self.OutputDirNameBis is not None and self.OutputDirNameBis != "-") and self.OutputDirNameBis != "" :
			if not os.path.isdir(self.OutputDirNameBis) :
				os.makedirs(self.OutputDirNameBis, mode=0777)
				self.logfile.write("Creation du repertoire : " + self.OutputDirNameBis, 0)	
		self.UserRunning		= "TEST"
		self.envMiddleOffice 	= Env(confEnv, self.logfile)
		self.__opInstallDir__	= self.envMiddleOffice.get_element("opInstallDir")
		self.repFluxOutTmp		= self.envMiddleOffice.get_element("RepFluxOutTmp")
		self.repTrt				= self.envMiddleOffice.get_element("RepFluxTrt")
		self.repLog				= self.envMiddleOffice.get_element("RepLog")
		self.RepWinRar    = self.envMiddleOffice.get_element("RepWinRar")
		self.Password = self.envMiddleOffice.get_element("Password")
		if (self.typeAssemblage == "MO"):
			self.__opWD__			= self.envMiddleOffice.get_element("opWDMO")
			self.__infoAppli__ 		= lectureApplisTab( self.__opWD__, self.__nomModele__, self.logfile )
		self.envProcess			= Env(confProcess, self.logfile)
		return
	
	def setZipfile(self):
		self.logfile.write("Paramétrage commande archivage",0)
		# si RepWinRar	C:\WinRAR
		# self.__commandline__ = self.RepWinRar + os.sep + "rar.exe a -ep1 -hp"+ self.Password +" " + self.__workingkDirZIPTMP__ + os.sep + self.__XMLFileOri__.rstrip('.CSU') + str(datetime.now().strftime("%H%M%S"))+".rar "+ self.__workingkDirAFP__+os.sep +"*.ixt "+ self.__workingkDirAFP__+ os.sep + "*.afp"
		self.__commandline__ = self.RepWinRar + os.sep + "7z.exe a -p"+ self.Password +" " + self.__workingkDirZIPTMP__ + os.sep + self.__XMLFileOri__.rstrip('.CSU') + str(datetime.now().strftime("%H%M%S"))+".7z "+ self.__workingkDirAFP__+os.sep +"*.ixt "+ self.__workingkDirAFP__+ os.sep + "*.afp"
	def Filesmove(self):
		if TYPE_TRT == "AFP":
			for file in os.listdir(self.__workingkDirZIPTMP__):
				try:
					filein  = self.__workingkDirZIPTMP__ + os.sep + file
					shutil.copy(filein, self.OutputDirName)
				except Exception, e:
					self.logfile.write("Erreur lors de la deplacement du fichier AFP de traitement : " + str(e), 108 )
				return
		elif TYPE_TRT == "PDF":
			for file in os.listdir(self.__workingkDirPDF__ ):
				try:
					filein  = self.__workingkDirPDF__  + os.sep + file
					shutil.copy(filein, self.OutputDirName)
				except Exception, e:
					self.logfile.write("Erreur lors de la deplacement du fichier PDF de traitement : " + str(e), 108 )
			pass
		elif TYPE_TRT == "MAIL" :
			if self.Suivi == "1" or self.OutputDirNameBis != "-" :	
				self.GenereSuivi(self.__workingkDirHTML__,self.OutputDirNameBis)
				for file in os.listdir(self.__workingkDirHTML__ ):
					try:
						filein  = self.__workingkDirHTML__  + os.sep + file
						fileout = self.OutputDirNameBis + os.sep + os.splitext(file)[0]+ ".html"
						shutil.copy(filein, fileout)
					except Exception, e:
						self.logfile.write("Erreur lors de la deplacement du fichier MAIL de traitement : " + str(e), 108 )
			pass
		elif TYPE_TRT == "MARKETING" :
			if self.Suivi == str("1") :
				# self.GenereSuivi(self.__workingkDirHTML__,self.OutputDirNameBis)
				self.GenereSuiviMarketing(self.__workingkDirHTML__,self.OutputDirNameBis)
				for file in os.listdir(self.__workingkDirHTML__ ):
					try:
						filein  = self.__workingkDirHTML__  + os.sep + file
						fileout = self.OutputDirNameBis + os.sep + file
						shutil.copy(filein, fileout)
					except Exception, e:
						self.logfile.write("Erreur lors de la deplacement du fichier MAIL de traitement : " + str(e), 108 )
		elif TYPE_TRT == "CHORUS" :
			for file in os.listdir(self.__workingkDirPDF__ ):
				try:
					if "Duplicata" in os.path.basename(file):
						filein  = self.__workingkDirPDF__  + os.sep + file
						shutil.copy(filein, self.OutputDirNameBis)
					else : 
						filein  = self.__workingkDirPDF__  + os.sep + file
						shutil.copy(filein, self.OutputDirName)
				except Exception, e:
					self.logfile.write("Erreur lors de la deplacement du fichier PDF Chorus de traitement : " + str(e), 108 )
		
		elif TYPE_TRT == "HYBRIDE" :
			if self.Suivi == "1" :
				if os.listdir(self.__workingkDirHTML__ ) != [] : # si on doit generé un suivi et que le dossier html est vide alors on fait le suivi du dossier pdf
					self.GenereSuivi(self.__workingkDirHTML__,self.OutputDirNameBis)
					for file in os.listdir(self.__workingkDirHTML__ ):
						try:
							filein  = self.__workingkDirHTML__  + os.sep + file
							fileout = self.OutputDirNameBis + os.sep + os.path.splitext(file)[0]+ ".html"
							shutil.copy(filein, fileout)
						except Exception, e:
							self.logfile.write("Erreur lors de la deplacement du fichier HTML de traitement : " + str(e), 108 )
						pass
					
				else :		
					self.GenereSuivi(self.__workingkDirPDF__,self.OutputDirNameBis)
					for file in os.listdir(self.__workingkDirPDF__ ):
						try:
							filein  = self.__workingkDirPDF__  + os.sep + file
							fileout  = self.OutputDirNameBis + os.sep + file
							shutil.copy(filein, fileout)
							filein  = self.__workingkDirPDF__  + os.sep + file
							shutil.copy(filein, self.OutputDirName)	
						except Exception, e:
							self.logfile.write("Erreur lors de la deplacement du fichier PDF de traitement : " + str(e), 108 )
			return

	def movelogfile(self, directory):
		if not os.path.isdir(directory):
			os.makedirs(directory, mode=0777)
		outputfile = self.__logfile__.replace(self.__workingkDir__, directory)
		outputfile = outputfile.replace(self.__XMLFile__ + os.sep, "")
		self.logfile.write("Deplacement du fichier log vers : " + outputfile, 0 )
		if os.path.isfile(outputfile):
			os.remove(outputfile)
		self.logfile.close()
		shutil.copy(self.__logfile__, outputfile)
		return

	def GenereSuivi(self,DirCourant,DirCible):
		fichier_suivi = DirCible + os.sep + os.path.splitext(self.__XMLFileOri__)[0] +".suivi.xml"
		fic_suivi = open(fichier_suivi,"wb")
		fic_suivi.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
		fic_suivi.write("<Root>\n")
		i=1
		for file in os.listdir(DirCourant):
			fic_suivi.write("<traceCourrier IdCourrier=\""+self.IndexFile[i][6]+"\" Utilisateur=\""+self.IndexFile[i][7]+"\" Chemin=\""+DirCible+"\" Fichier=\""+os.path.basename(file)+"\" CodeSociete=\"CESU\">\n")
			i=+1
		fic_suivi.write("</Root>")	
		fic_suivi.close()

	def GenereSuiviMarketing(self,DirCourant,DirCible):
		fichier_suivi = DirCible + os.sep + os.path.splitext(self.__XMLFileOri__)[0] +".suivi.xml"
		tree= ET.parse(self.Config)
		root = tree.getroot()
		for Param in root.getiterator('Root') :
			for UserParams in Param.getiterator('UserParams') :
				Utilisateur = UserParams.find('Utilisateur').text
				CodeOperation = UserParams.find('CodeOperation').text
				Commentaire= UserParams.find('Commentaire').text
				TypeEvenement = UserParams.find('TypeEvenement').text
				CodeSociete = UserParams.find('CodeSociete').text
		fic_suivi = open(fichier_suivi,"wb")
		fic_suivi.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
		fic_suivi.write("<Root>\n")
		i=1
		for file in os.listdir(DirCourant):
			 #<traceMarketing CodeOperation="EDEM1903" Affilie="0" TypeEvenement="1010" CommentaireEvenement="Campagne email RIB pour transactions UP" Utilisateur="SYS" Chemin="D:\PROD\Archivage\ArchivagePNI\CampagneEmail\CRT\SuiviXML\Attente" Fichier="0_EDEM1903_0_2019-03-15_14.23.15_EDEM1903.html" CodeSociete="CRT"/>
			fic_suivi.write("<traceMarketing CodeOperation=\""+CodeOperation+"\" Affilie=\"" +self.IndexFile[i][2]+"\" TypeEvenement=\""+TypeEvenement+" \"CommentaireEvenement=\""+Commentaire+"\" Utilisateur=\""+Utilisateur+"\" Chemin=\""+DirCible+"\" Fichier=\""+os.path.basename(file)+"\" CodeSociete=\""+CodeSociete+"\">\n")
			i=+1
		fic_suivi.write("</Root>")	
		fic_suivi.close()		
	''' Partie composition commune 
		Traitement des vpf et index (techsort) '''

	def __setWorkingDir__(self):
		try:
			if not os.path.isdir(self.__workingkDirIN__):
				os.makedirs( self.__workingkDirIN__, mode=0777 )
			if not os.path.isdir(self.__workingkDirOUT__ ):
				os.makedirs( self.__workingkDirOUT__ , mode=0777 )
			if not os.path.isdir(self.__workingkDirTRC__):
				os.makedirs( self.__workingkDirTRC__, mode=0777 )
			if not os.path.isdir(self.__workingkDirPARAM__):
				os.makedirs( self.__workingkDirPARAM__, mode=0777 )
		except Exception, e:
			# Probleme detecte
			self.logfile.write("Assemblage : Erreur detectee a la creation d'un sous-repertoire de travail : " + str(e) , 102)		
		return
	
	def setCommandTechsort(self):
		self.__commandline__ = self.__opInstallDir__ + "/bin/techsort" + \
			" -c " + self.RepRessources + os.sep + "exec.cmd" + \
			" -i " + self.__workingkDirVPF__ + os.sep + self.__XMLFilesplitext__ + ".vpf" + \
			" -S OUTPUT_DIR=" + self.__workingkDirECLATEMENT__ + \
			" -S REPRESSOURCE=" + self.RepRessources + \
			" -E opInstallDir=" + self.__opInstallDir__ + \
			" -S FILE=" + self.__XMLFileOri__.rstrip('.CSU') + " -nomerge -NamedIn -NamedOut -NoTechInfo"

	def FormatageTechSort(self, idxErreur):
		''' modification des index pour respecter le format attendu '''
		try :
			self.logfile.write("creation index :",0)
			idx_files = []	
			list_fic = os.listdir(self.__workingkDirECLATEMENT__)	
			for file in list_fic :
				if (os.path.isfile(self.__workingkDirECLATEMENT__+ os.sep + file) and file.endswith(".ind")) :
					idx_files.append(file)
			for element in idx_files :
				self.logfile.write("bouclage file index:",0)
				txtData = open(self.__workingkDirECLATEMENT__ + os.sep +element, 'r')
				ixtData = open(self.__workingkDirAFP__+ os.sep + element.rstrip("_s.ind") +".ixt", "w")
				self.logfile.write("creation fichier :"+self.__workingkDirAFP__+ os.sep + element.rstrip("_s.ind") +".ixt",0)
				numPli=-1
				for row in txtData:
					numPli += 1
					if not(row.startswith("0_-4")):
						idx = (row.rstrip('\r\n')).split('\t')
						if numPli <= 0 :
							lenIdx = len(idx)
							if idx[lenIdx-1] == '' :
								lenIdx -= 1
							ixtData.write(unicode('L:          Job_name    Check_NumberPage_Number Position    Size        BC01_CODECL         Date_Impression     Cpt_NumPag_Dos      Cpt_Pag_Dossier     Cpt_Dossier         Code_Post_BC02      Code_Post_LM01      Code_Post_LP01      Code_Post_FA01      Nom_Fichier                   FA01_CODCLI         LP01_CODEAF         LM01_CODEAF         '))
							if lenIdx > 20 :
								ixtData.write(unicode('AN\n'))
							else:
								ixtData.write(unicode('\n'))
						else: 
							idx[4]=idx[4].rjust(10,'0')
							idx[9]=idx[9].ljust(20)
							idx[10]=idx[10].ljust(20)
							#probleme espace sur idx11
							#idx[11]=idx[11].lstrip(' ')
							idx[17]=idx[17].lstrip('0').ljust(20,' ')
							ixtData.write(unicode(''.join(idx[1:lenIdx])+'\n'))
				ixtData.close()
			self.logfile.write("FormatageTechSort terminé",0)
		except Exception,e:
			self.logfile.write("erreur formatage index : "+str(e),0)

	def AjoutPJGenere(self):
		if os.listdir(self.__workingkDirPDF__) != []:
			try :
				old_index= self.__workingkDirVPF__ + os.sep + self.fichierVPFind
				new_index= self.__workingkDirVPF__ + os.sep + self.fichierVPFind+"new"
				buffer_old_index = open(old_index,"r")
				buffer_new_index = open(new_index,"w")
				new_data = ""
				i=0
				for line in buffer_old_index :
					if i != 0 :
						tab_line=line.strip('\n').split('\t')
						if tab_line[2] == "" :
							buffer_new_index.write(tab_line[0]+"\t"+tab_line[1]+"\t"+self.__workingkDirPDF__+ os.sep + os.path.basename(os.path.splitext(self.__XMLFileOri__)[0]) + ".pdf"+"\t"+tab_line[3]+"\t"+tab_line[4]+"\t"+tab_line[5]+"\t"+tab_line[6]+"\t"+tab_line[7]+"\t"+tab_line[8]+"\n")
						else :
							buffer_new_index.write(tab_line[0]+"\t"+tab_line[1]+"\t"+tab_line[2]+";"+self.__workingkDirPDF__+ os.sep +os.path.basename(os.path.splitext(self.__XMLFileOri__)[0]) + ".pdf"+"\t"+tab_line[3]+"\t"+tab_line[4]+"\t"+tab_line[5]+"\t"+tab_line[6]+"\t"+tab_line[7]+"\t"+tab_line[8]+"\n")
					else : 
						buffer_new_index.write(line)
					i=+1
				buffer_old_index.close()
				buffer_new_index.close()
				shutil.copy(new_index,old_index)
			except Exception,e : print e

	def setCommandRenderingAFP(self): #rendering
		try :
			self.__commandline__ = self.__opInstallDir__ + "/bin/techcodr" + \
				" -E opWD=" + self.__opWD__ + \
				" -E opFam=" + self.__infoAppli__["opFam"]  + \
				" -E opAppli=" + self.__infoAppli__["opAppli"]  + \
				" -E opInstallDir=" + self.__opInstallDir__ + \
				" -E opMaxWarning=0 -E opDoubleByte=1" + \
				" -iv " + self.__workingkDirECLATEMENT__ + os.sep + self.__XMLFileOri__.rstrip('.CSU') + "-MA_s.vpf" + \
				" -od " + self.__workingkDirAFP__ + os.sep + self.__XMLFileOri__.rstrip(".CSU") + "-MA.afp" + \
				" -dn afp8"
			self.execCommande("RENDERING MAITRE", 104)
			if not self.__XMLFileOri__.startswith('BRC') :
				self.__commandline__ = self.__opInstallDir__ + "/bin/techcodr" + \
					" -E opWD=" + self.__opWD__ + \
					" -E opFam=" + self.__infoAppli__["opFam"]  + \
					" -E opAppli=" + self.__infoAppli__["opAppli"]  + \
					" -E opInstallDir=" + self.__opInstallDir__ + \
					" -E opMaxWarning=0 -E opDoubleByte=1" + \
					" -iv " + self.__workingkDirECLATEMENT__ + os.sep + self.__XMLFileOri__.rstrip('.CSU') + "-E1_s.vpf" + \
					" -od " + self.__workingkDirAFP__ + os.sep + self.__XMLFileOri__.rstrip(".CSU") + "-E1.afp" + \
					" -dn afp8"
				self.execCommande("RENDERING ESCLAVE", 104)
		except Exception,e : 
			self.logfile.write("Probleme Rendering AFP : %s" % str(e),1)

	def setCommandRenderingPDF(self): #rendering
		try :
			self.__commandline__ = self.__opInstallDir__ + "/bin/techcodr" + \
				" -E opWD=" + self.__opWD__ + \
				" -E opFam=" + self.__infoAppli__["opFam"]  + \
				" -E opAppli=" + self.__infoAppli__["opAppli"]  + \
				" -E opInstallDir=" + self.__opInstallDir__ + \
				" -E opMaxWarning=0 -E opDoubleByte=1" + \
				" -iv " + self.__workingkDirVPF__ + os.sep + self.__XMLFilesplitext__ + ".vpf" + \
				" -od " + self.__workingkDirPDF__ + os.sep + os.path.basename(os.path.splitext(self.__XMLFileOri__)[0]) + ".pdf" + \
				" -dn pdfuni"
			self.execCommande("RENDERING", 104)
		except Exception,e : 
			self.logfile.write("Probleme Rendering PDF : %s" % str(e),1)

	def setCommandRenderingPDFChorus(self): #rendering
		try :
			self.__commandline__ = self.__opInstallDir__ + "\\bin\\techcodr" + \
				" -E opWD=" + self.__opWD__ + \
				" -E opFam=" + self.__infoAppli__["opFam"]  + \
				" -E opAppli=" + self.__infoAppli__["opAppli"]  + \
				" -E opInstallDir=" + self.__opInstallDir__ + \
				" -E opMaxWarning=0 -E opDoubleByte=1" + \
				" -iv " + self.__workingkDirVPF__ + os.sep + self.__XMLFilesplitext__ + ".vpf" + \
				" -od " + os.path.basename(os.path.splitext(self.__XMLFileOri__)[0]) + ".pdf" + \
				" -dn pdfeclatuni"
			self.execCommande("RENDERING", 104)
		except Exception,e : 
			self.logfile.write("Probleme Rendering PDF CHORUS: %s" % str(e),1)
			
	def setCommandRenderingHTML(self): #rendering
		try :
			self.__commandline__ = self.__opInstallDir__ + "/bin/techcodr" + \
				" -E opWD=" + self.__opWD__ + \
				" -E opFam=" + self.__infoAppli__["opFam"]  + \
				" -E opAppli=" + self.__infoAppli__["opAppli"]  + \
				" -E opInstallDir=" + self.__opInstallDir__ + \
				" -E opMaxWarning=0 -E opDoubleByte=1" + \
				" -iv " + self.__workingkDirVPF__ + os.sep + self.__XMLFilesplitext__ + ".vpf" + \
				" -oe " + self.__workingkDirHTML__ + os.sep +self.__XMLFileOrisplitext__ + ".html" + \
				" -de emlHTML"
				# " -dn emlHTML"
			self.execCommande("RENDERING HTML", 104)
		except Exception,e : 
			self.logfile.write("Probleme Rendering HTML : %s" % str(e),1)

	def SendMail(self):
		"Lecture des informations de connexion"
		runConfigMail(self.RepRessources + os.sep + "cesu.cfg",self.__workingkDirHTML__ + os.sep + self.__XMLFileOrisplitext__ + "_001.html", self.__workingkDirVPF__ + os.sep +  self.__XMLFilesplitext__+ ".vpf.ind", self.logfile )

	def SendMailMarketing(self):
		"Lecture des informations de connexion"
		i = 1 
		Code_moins1 = ""
		for line in self.IndexFile[i:-1] :
			Code = self.IndexFile[i][7]
			while Code_moins1 != Code :
				runConfigMailMarketing(self.RepRessources + os.sep + "cesu.cfg",self.__workingkDirHTML__ + os.sep + Code +".html",self.IndexFile[i],self.logfile)
				Code_moins1 = Code
			i=i+1	
		
	def ChargeIndex(self):
			my_ind = open(self.__workingkDirVPF__+ os.sep +self.fichierVPFind,'rb')
			i = 0 
			for line in my_ind.readlines():
				tab_line=line.strip('\n').split('\t')
				self.IndexFile.append([tab_line[1],tab_line[2],tab_line[3],tab_line[4],tab_line[5],tab_line[6],tab_line[7],tab_line[8]])
				i=i+1
			my_ind.close()
		
def lectureApplisTab(opWD,modele,logfile):
	"""
	Recuperation des parametres du modele depuis '/opWD/common/tablei/applis.tab'
	"""
	applistab = {}
	# Chargement de la table
	filename = "%s/common/tablei/applis.tab" % (opWD)
	try :
		logfile.write( "Lecture '" + filename + "' OK", 0 )
		fic = open(filename, "r")
	except Exception, e:
		logfile.write( "Probleme d'ouverture en lecture de la table '" + filename + "' : " + str(e) + "'", 100)
	#
	# Lecture ligne a ligne de la table avec recherche du template (modele)
	found = False
	try:
		for line in fic :
			tab = line.strip("\n").split("\t")
			if ( tab[0] == modele ) :
				applistab["resdescid"]  = tab[1]
				applistab["opAppli"]    = tab[3]
				applistab["opFam"]      = tab[4]
				applistab["dataloader"] = tab[5]
				found = True
				break
	except Exception, e:
		logfile.write( "Probleme de lecture de la table '" + filename + "' avec la cle '" + modele + "' : '" + str(e) + "'", 100)
	#
	# Fermeture du fichier
	fic.close()
	#
	# Test si modele trouve dans la table
	if not found :
		logfile.write( "Absence du modele '" + modele + "' dans la table '" + filename + "'", 100 )
	#
	# Ecriture dans la trace des parametres trouves
	logfile.write( "Parametres de modeles trouves ", 0) #+ str(applistab) ,0 )
	#
	return applistab

def runCommand(commandline, Environnement):
	'''Fonction de lancement d'une commande'''
	try:
		exitCode = 0
		if (os.name == 'posix'): 
			print "commandline:%s" % commandline
			process = Popen(commandline, shell=True, stdin=PIPE, stdout=PIPE, stderr=PIPE)
			streamdata = process.communicate()[0]
			exitCode = process.returncode
		else:
			# process = Popen(shlex.split(commandline), stdout=PIPE, stderr=PIPE)
			process = Popen(commandline, shell=True, stdin=PIPE, stdout=PIPE, stderr=PIPE)
			streamdata,streamerrdata = process.communicate()
			exitCode = process.returncode
	except Exception,e :
		self.logfile.write(str(e),1)
	return exitCode

def getInfoIndex(file,num_ligne):
	IdEmet = -1
	"Destinataires	PiecesJointes	IdEmetteur	SujetMail	GenerationMail Suivi IdCourrier CodeUtilisateur #IdAffilie# "
	Destinataires = ""
	PiecesJointes = ""
	IdEmet = ""
	SujetMail = ""
	GenerationMail= ""
	Suivi = ""
	IdCourrier= ""
	CodeUtilisateur = ""
	my_ind_file=open(file,'r')
	next(my_ind_file)
	i=1
	for line in my_ind_file :
		if i == num_ligne :
			tab_line=line.strip('\n').split('\t')
			if  len(tab_line) == 10 :
				Destinataires,PiecesJointes,IdEmet,SujetMail,GenerationMail,Suivi,IdCourrier,CodeUtilisateur,IdAffilie = tab_line[1],tab_line[2],tab_line[3],tab_line[4],tab_line[5],tab_line[6],tab_line[7],tab_line[8],tab_line[9]
			else:
				Destinataires,PiecesJointes,IdEmet,SujetMail,GenerationMail,Suivi,IdCourrier,CodeUtilisateur,IdAffilie = tab_line[1],tab_line[2],tab_line[3],tab_line[4],tab_line[5],tab_line[6],tab_line[7],tab_line[8],""
			i=+1
	my_ind_file.close()
	return Destinataires,PiecesJointes,IdEmet,SujetMail,GenerationMail,Suivi,IdCourrier,CodeUtilisateur,IdAffilie



def getInfoIndexMarketing(file,num_ligne,num_col):
	IdEmet = -1
	"Destinataires	PiecesJointes	IdEmetteur	SujetMail	GenerationMail Suivi IdCourrier CodeUtilisateur"
	my_ind_file=open(file,'r')
	next(my_ind_file)
	i=1
	for line in my_ind_file :
		if i == num_ligne :
			tab_line=line.strip('\n').split('\t')
			result = tab_line[num_col]
			i=+1
	my_ind_file.close()
	return result
	
	
def __runConfigMail(fileConfig,filehtml,index,logfile):
	#renvoie le destinataire
	temp_dict = {}
	SujetMail = ""
	# lecture html
	data = open(filehtml, 'r').read()
	tree= ET.parse(fileConfig)
	root = tree.getroot()
	#recuperation donnés indexés
	receivers,PJs,IdEmetteur,SujetMail,GenerationMail,Suivi,IdCourrier,CodeUtilisateur,IdAffilie = getInfoIndex(index,1)
	#lecture XML config dans dossier commun
	Cle=CalcIdCle(IdAffilie)
	logfile.write( "Cle:"+str(Cle), 0)
	for CI in root.getiterator('ConfigurationInternet') :
		for config in CI.getiterator('Config') :
			for servers in config.getiterator('Servers') :
				i = 0
				for server in config.getiterator('Server') :
					temp_dict[i]=server.attrib
					i=i+1
	server = ''
	sender = ''
	for k,v in 	temp_dict.iteritems():
		print "recherche serveur avec IDEmetteur:%s" %  str(IdEmetteur)
		if v['idEmetteur'] == str(IdEmetteur) :
			server = v['smtpServer']
			sender = v['senderEmail']
			# print 'trouve'
	try:
		msg = MIMEMultipart()
		msg['From'] = sender
		msg['To'] = receivers
		msg['Subject'] = SujetMail
		try :
			if Cle != "":
					new_data = data.replace("#VarIdCle#",str(Cle))
					data = new_data
		except Exception,e : print "prob rempalcement cle:" + str(e)
		msg.attach(MIMEText(data,"html","utf-8")) 
		if PJs != '' :
			for PJ in PJs.split(';'):
				if os.path.isfile(PJ) :
					part = MIMEBase('application', "octet-stream")
					part.set_payload(open(PJ, "rb").read())
					encoders.encode_base64(part)
					part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(PJ))
					msg.attach(part)
				else : 
					logfile.write( "Probleme Pièce jointe  '" + PJ + "' absente : " "'", 110)
		my_config_mail = os.getcwd() + os.sep + "config_mail.txt"
		if not os.path.isfile(my_config_mail) :
			logfile.write( "Probleme config_mail.txt", 110)
		gmail_user = getConfAcc(os.getcwd() + os.sep + "config_mail.txt",u"user")
		gmail_password = getConfAcc(os.getcwd() + os.sep + "config_mail.txt",u"password")
		try :
			s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
			s.ehlo()
			s.login(gmail_user, gmail_password)
			for receiver in receivers.split(';') :
				s.sendmail(sender, receiver, msg.as_string())
			s.close()
		except  Exception,e:  
			print 'Something went wrong...'+str(e)
		
	except Exception,e:
		logfile.write( "Probleme envoi mail "+str(e), 0)
	pass

def runConfigMail(fileConfig,filehtml,index,logfile):
	#renvoie le destinataire
	temp_dict = {}
	SujetMail = ""
	# lecture html
	data = open(filehtml, 'r').read()
	tree= ET.parse(fileConfig)
	root = tree.getroot()
	#recuperation donnés indexés
	receivers,PJs,IdEmetteur,SujetMail,GenerationMail,Suivi,IdCourrier,CodeUtilisateur,IdAffilie = getInfoIndex(index,1)
	#lecture XML config dans dossier commun
	Cle=CalcIdCle(IdAffilie)
	logfile.write( "Cle:"+str(Cle), 0)
	for CI in root.getiterator('ConfigurationInternet') :
		for config in CI.getiterator('Config') :
			for servers in config.getiterator('Servers') :
				i = 0
				for server in config.getiterator('Server') :
					temp_dict[i]=server.attrib
					i=i+1
	server = ''
	sender = ''
	for k,v in 	temp_dict.iteritems():
		print "recherche serveur avec IDEmetteur:%s" %  str(IdEmetteur)
		if v['idEmetteur'] == str(IdEmetteur) :
			server = v['smtpServer']
			sender = v['senderEmail']
			# print 'trouve'
	try:
		msg = MIMEMultipart()
		msg['From'] = sender
		msg['To'] = receivers
		msg['Subject'] = SujetMail
		try :
			if Cle != "":
					new_data = data.replace("#VarIdCle#",str(Cle))
					data = new_data
		except Exception,e : print "prob rempalcement cle:" + str(e)
		msg.attach(MIMEText(data,"html","utf-8")) 
		if PJs != '' :
			for PJ in PJs.split(';'):
				if os.path.isfile(PJ) :
					part = MIMEBase('application', "octet-stream")
					part.set_payload(open(PJ, "rb").read())
					encoders.encode_base64(part)
					part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(PJ))
					msg.attach(part)
				else : 
					logfile.write( "Probleme Pièce jointe  '" + PJ + "' absente : " "'", 110)
		s = smtplib.SMTP(server)
		s.sendmail(sender, receivers.split(';'), msg.as_string())		
	except Exception,e:
		logfile.write( "Probleme envoi mail "+str(e), 1)
		return -1
	pass

def runConfigMailMarketing(fileConfig,filehtml,index,logfile):
	#renvoie le destinataire
	temp_dict = {}
	SujetMail = ""
	data = open(filehtml, 'r').read()
	tree= ET.parse(fileConfig)
	root = tree.getroot()
	receivers,PJs,IdEmetteur,SujetMail,GenerationMail,Suivi,IdCourrier,CodeUtilisateur = index
	for CI in root.getiterator('ConfigurationInternet') :
		for config in CI.getiterator('Config') :
			for servers in config.getiterator('Servers') :
				i = 0
				for server in config.getiterator('Server') :
					temp_dict[i]=server.attrib
					i=i+1
	server = ''
	sender = ''
	for k,v in 	temp_dict.iteritems():
		if v['idEmetteur'] == str(IdEmetteur) :
			server = v['smtpServer']
			sender = v['senderEmail']
	try:
		if receivers != "" :
			msg = MIMEMultipart()
			msg['From'] = sender
			msg['To'] = receivers
			msg['Subject'] = SujetMail
			msg.attach(MIMEText(data,"html","utf-8")) 
			s = smtplib.SMTP(server)
			s.sendmail(sender, receivers.split(';'), msg.as_string())		
	except Exception,e:
		logfile.write( "Probleme envoi mail "+str(e), 1)
		return -1
		

	
def CheckEnvoiParMail(file):
	try :
		my_ind_file=open(file,'r')
		next(my_ind_file)
		i=1
		for line in my_ind_file :
			if i == 1:
				tab_line=line.split('\t')
				envoiMail = tab_line[5].strip('\r\n')
				i=+1
		return envoiMail
	except : 
		return -1
		pass

def getConfAcc(file,cle):
	try :
		buffer=open(file,'r')
		for line in buffer:
			cle_courante,valeur_courante = line.strip('\r\n').split("\t")
			if cle_courante == cle :
				return valeur_courante
		buffer.close()
	except Exception,e :
		print"Probleme lecture fichier: "+str(e)

def CalcIdCle(IdAffilie):
	mult = 128
	cle = 0
	try : 
		VarIdAffilie=str(IdAffilie.rjust(7,'0'))
		A1 = int(VarIdAffilie[0])
		A2 = int(VarIdAffilie[1])
		A3 = int(VarIdAffilie[2])
		A4 = int(VarIdAffilie[3])
		A5 = int(VarIdAffilie[4])
		A6 = int(VarIdAffilie[5])
		A7 = int(VarIdAffilie[6])
	  
		cle = A1*128 + A2*64 + A3*32 + A4*16 + A5*8 + A6*4 + A7*2
		print str(cle)
		cle = cle - (11*(cle/11))
		if cle == 10 :
			cle = 0
	 	elif cle == 11 : 
	 		cle = 1
		VarIdCle = str(cle)
		return VarIdCle
	except : 
	  VarIdCle = ""
	  return VarIdCle


""" MAIN PROGRAM """		
if __name__=='__main__' :
	''' Initialisation '''
	ENVOI_MAIL 	= 0
	INPUT_FILE 		= sys.argv[1]
	TYPEMODELE 		= sys.argv[2]
	TYPE_TRT 		= os.path.basename(sys.argv[1]).split('.')[0]
	nomModele       = os.path.basename(sys.argv[1]).split('.')[1]
	confEnv      	= cheminConfig + os.sep +"env.tab"
	confProcess   	= cheminConfig + os.sep +"process.tab"
	Environnement 	= None
	confAppli		= os.path.splitext(INPUT_FILE)[0] + ".acc"
	""" Creation de l'objet principal """
	Environnement 	= MiddleOffice(INPUT_FILE, nomModele, confEnv, confProcess, confAppli)
	''' MODE DEBUG '''
	DEBUG = 0
	if TYPE_TRT == "MARKETING" :
		'''COMPO VPF avec ligne de commande spécifique'''
		Environnement.setCommandCompositionVPF_Marketing()
		Environnement.execCommande("COMPOSITION VPF Marketing (spécifique)", 104)
	elif TYPE_TRT == "CHORUS":
		Environnement.setCommandCompositionVPFChorus()
		Environnement.execCommande("COMPOSITION VPF CHORUS", 104)
	else :
		''' COMPOSITION VPF '''
		Environnement.setCommandCompositionVPF()
		Environnement.execCommande("COMPOSITION VPF", 104)
		''' RECUPERATION INDEX'''
	if TYPE_TRT == "HYBRIDE" or TYPE_TRT == "MAIL" or TYPE_TRT == "MARKETING": 
		Environnement.ChargeIndex()
		Environnement.ENVOI_MAIL,Environnement.Suivi = Environnement.IndexFile[1][4:6]
		if Environnement.ENVOI_MAIL == "1" :
			Environnement.Suivi= str(getInfoIndex(Environnement.__workingkDirVPF__+ os.sep +Environnement.fichierVPFind,1)[5])
			Environnement.logfile.write("ENVOI_MAIL:"+str(Environnement.ENVOI_MAIL), 0)
	''' TECHSORT '''
	if TYPE_TRT == "AFP" :
		Environnement.setCommandTechsort()
		Environnement.execCommande("TECHSORT", 105)
		Environnement.FormatageTechSort(0)
	''' RENDERING '''
	if TYPE_TRT == "CHORUS" :
		Environnement.setCommandRenderingPDFChorus()
	if TYPE_TRT == "AFP" :
		Environnement.setCommandRenderingAFP()
	if TYPE_TRT == "PDF" :
		Environnement.setCommandRenderingPDF()
	if TYPE_TRT == "MAIL" :
		Environnement.setCommandRenderingHTML()	
	if TYPE_TRT == "HYBRIDE":
		if Environnement.ENVOI_MAIL == "1" :
			Environnement.setCommandRenderingHTML()
		Environnement.setCommandRenderingPDF()
	''' COMPRESSION ZIP '''
	if TYPE_TRT == "AFP":
		Environnement.setZipfile()
		Environnement.execCommande("RAR", 106)
	''' ENVOIE MAIL '''
	if TYPE_TRT == "MAIL" :	
		Environnement.SendMail()
		Environnement.logfile.write("MAIL ENVOYE", 0)		
	if TYPE_TRT == "HYBRIDE" :	
		if Environnement.ENVOI_MAIL == "1" :
			Environnement.AjoutPJGenere()
			Environnement.logfile.write("Ajout Pieces Jointes", 0)
			Environnement.SendMail()
			Environnement.logfile.write("MAIL ENVOYE", 0)
	if  TYPE_TRT == "MARKETING" :
		Environnement.SendMailMarketing()
		Environnement.logfile.write("MAIL ENVOYE", 0)
	''' FIN DU TRAITEMENT ''' 
	Environnement.Filesmove()
	Environnement.logfile.write("", 0)
	Environnement.logfile.write("FIN DU TRAITEMENT", 0)
	Environnement.movelogfile(Environnement.repLog + os.sep)
	Environnement.exit(0)


