#!C:/Python27_64/python
# -*-coding:utf-8 -*

#-------------------------------------------------------------------------
# SCRIPT DE SCRUTATION 
# DEVELOPPE PAR : BASTIEN FERRAGU (SEFAS)
# V1.0.0 - 2015-10-12 -
# V1.0.1 - 2015-10-14 - Suppresion de nombreux commentaires inutiles 
# V1.1.0 - 2019-06-XX - Version CRT
# V1.2.0 - 2019-09-XX - Ajout de la fillière Marketing
#-------------------------------------------------------------------------

import os, sys, subprocess, shutil, glob, time
import shlex
import tarfile, zipfile

from threading 	import Thread
from Queue 		import Queue
from datetime 	import datetime
from random 	import randint
from time 		import sleep
from functions	import *

ROOT_PATH	= sys.argv[1]
PYTHON_PATH = sys.argv[2]

cheminConfig 	= ROOT_PATH + os.sep + "param"
SCRIPT_PATH		= ROOT_PATH + os.sep + "applis"+ os.sep + "execution.py"

class LogFile:
	def __init__(self, nomlog):
		self.__nomlog__ 	= nomlog
		self.__fichierlog__	= None
		self.__erreur__		= None
		try:
			self.__fichierlog__ = open(self.__nomlog__, "a+", 0)
		except Exception, e:
			self.__erreur__ == "Probleme a l'ouverture du fichier de log " + self.__nomlog__ + " : [" + str(e) + "]"
			return
		return	
	def write(self, message, exitCode):
		logtype, errocode  = self.set_erreurType(exitCode)
		self.__fichierlog__.write(logtype + maintenant() + ' | ' + message + '\n')
		if (exitCode < 100):
			# 0 = OK
			self.__fichierlog__.flush
		else:
			""" Fonction d'ecriture d'un message d'erreur dans la trace et gestion de fin du script """
			self.__fichierlog__.write(logtype + maintenant() + ' | ERRORCODE : ' + errocode + ' | EXITCODE : ' + str(exitCode) + '\n')
			self.__fichierlog__.flush
			affichedebug(message, DEBUG)
			affichedebug("Exit code : " + str(exitCode), DEBUG)
			sys.exit(exitCode)			
		return
	def set_erreurType(self, exitCode):
		""" Definit le type du message à afficher """
		if (exitCode==0):
			logtype = "I | "
			return logtype, "0"
		else:
			logtype = "E | "
			errocode = str(exitCode)
			return logtype, errocode
	def get_erreur(self):
		""" Fonction de renvoi d'erreur liee a la gestion de la trace """
		return self.__erreur__
	def close(self):
		self.__fichierlog__.close()
		return

class Loadfile:
	def __init__(self, nomFichier):
		self.__nomFichier = nomFichier
		self.__dico       = {}
		self.__erreur		= None

		self.load()
		return
	def load(self):
		""" Fonction de chargement du fichier de table """
		try:
			ficTMP = file(self.__nomFichier, "rU")
			for ligne in ficTMP:
				listeTMP = ligne.strip().split("\t")
				self.__dico[ listeTMP[0] ] = listeTMP[1]
			ficTMP.close()
		except Exception, e:
			pass
		return
	def get_erreur(self):
		""" Fonction de renvoi d'erreur """
		return self.__erreur
	def get_element(self, cle):
		""" Fonction de recherche d'une entree dans la table de configuration de l'environnement """
		try:
			element = self.__dico[cle]
		except Exception, e:
			self.__erreur = "Pas d'entree dans la table d 'Environnement' pour la cle '" + cle + "'"
			element = "erreur"
		return element

class LoadBusFile():
	def __init__(self, nomFichier):
		self.listeecanaux		= []
		self.inputdirs      	= []
		self.archivedirs			= []
		self.templatesname		= []
		self.prefixesflux		= []
		self.outputdirs			= []
		self.prefixeMaquette 	= []
		self.outputdirsbis      = []
		self.__nomFichier		= nomFichier
		# self.__userrunning	= UserRunning
		self.__erreur			= None
		self.load()
		return

	def load(self):
		""" Fonction de chargement du fichier de table """
		try:
			ficTMP = file(self.__nomFichier, "rU")
			for ligne in ficTMP:
				listeTMP = ligne.strip().split("\t")
				self.listeecanaux.append(listeTMP[0])
				self.inputdirs.append(listeTMP[1])
				self.archivedirs.append(listeTMP[2])
				self.templatesname.append(listeTMP[3])
				self.prefixesflux.append(listeTMP[4])
				self.outputdirs.append(listeTMP[5])
				self.prefixeMaquette.append(listeTMP[6])
				self.outputdirsbis.append(listeTMP[7])
			ficTMP.close()
		except Exception, e:
			print str(e)
			pass
		return
	def get_erreur(self):
		""" Fonction de renvoi d'erreur """
		return self.__erreur
	def get_element(self, cle):
		""" Fonction de recherche d'une entree dans la table de configuration de l'environnement """
		try:
			element = self.__dico[cle]
		except Exception, e:
			self.__erreur = "Pas d'entree dans la table d 'Environnement' pour la cle '" + cle + "'"
			element = "erreur"
		return element
		
		
class SousProcess(Thread):
	def __init__(self, command):
		Thread.__init__(self)
		self.command = command
		self.__sp = None
		self.coderetour = ""		
	def run(self):	
		self.__sp = subprocess.Popen(self.command,shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
		streamdata = self.__sp.communicate()
	def stop(self):
		self.Terminated = True
		# fin du sous process (pour liberer les repertoires)
	def getRetour(self):
		""" """
		self.coderetour = self.__sp.returncode 
		return self.coderetour


def isSpecifique(nomFichier):
	if nomFichier.startswith("SPECIFIQUE"):
		return True
	else:
		return False

def CleanFichiersAnnexes(nomFichier) :
	if os.path.isfile(nomFichier) :
		my_file = open(nomFichier,'r')
		my_file_buffer = my_file.readlines()
		for line in my_file_buffer :
			if (line.startswith('HTML') or line.startswith('CONFIG')) :
				line_clean= line.strip('\n').split('\t')[1]
				if os.path.isfile(line_clean) :
					logfile.write("suppression du fichier :" + line_clean, 0)
					os.remove(line_clean)
				else : 
					logfile.write("fichier absent :" + line_clean, 0)
		my_file.close()
	else :
		logfile.write("Absence du fichier ACC:" + nomFichier, 1)	



''' EXECUTION DU THREAD '''
def do_stuff(QueueExecution, idxthread, file, nom_complet):
	while True:
		IdTrtFlux		= ""
		FicAcc 			= ""
		QueueExecution.get()
		QueueExecution.task_done()
		''' creation du fichier lock sur le thread '''
		fichierLock = os.path.abspath(RepFluxTrt) + os.sep + "thread" + str(idxthread+1) + ".lock"
		open(fichierLock , "w", 0)
		IdTrtFlux 		= file
		rep_traitement = RepFluxTrt  + os.sep + IdTrtFlux
		fichier_trt = rep_traitement + os.sep + IdTrtFlux
		nom_complet_acc = os.path.splitext(nom_complet)[0] + ".acc"
		FicAcc 			= os.path.splitext(file)[0] + ".acc"
		fichier_trt_acc = rep_traitement + os.sep + FicAcc
		cmd			= PYTHON_PATH + os.sep + "python " + SCRIPT_PATH + " " + fichier_trt + " MO DEFAULT" + " " + ROOT_PATH
		# deplacement du fichier recu dans le repertoire de traitement
		if not os.path.isdir(rep_traitement):
			# os.makedirs(rep_traitement, 0777)
			os.mkdir(rep_traitement, 0777)
		else:
			""" le repertoire existe deja => que faire """
			pass
		''' deplacement des fichiers dans le repertoire de traitement '''
		if os.path.isfile(fichier_trt):
			os.remove(fichier_trt)
		shutil.move(nom_complet, os.path.abspath(fichier_trt))
		time.sleep(0.2)
		if os.path.isfile(fichier_trt_acc):
			os.remove(fichier_trt_acc)
		shutil.move(nom_complet_acc, os.path.abspath(fichier_trt_acc))	
		logfile.write("Lancement de la commande : " + cmd, 0)
		coderetour = ""
	 	""" Thread qui fonctionne """
		P = SousProcess(cmd)
		P.start()
		P.join()
		returnlockfile = RepFluxTrt + os.sep + IdTrtFlux + ".lock"
		try:
			coderetour = open(returnlockfile, "r").readlines()[0]
			os.remove(returnlockfile)
		except:
			coderetour = "-1"
		logfile.write("Code retour " + str(coderetour) + " pour le fichier " + file, 0)
		''' Suppression du fichier lock (debloquer le thread) '''
		if os.path.isfile(fichierLock):
			os.remove(fichierLock)
			logfile.write("Suppression du fichir lock " + rep_traitement, 0)
		'''GESTION DU RETOUR DU SOUS PROCESS'''
		if int(coderetour) == 0:
			''' Le process s'est execute correctement on supprime le repertoire de traitement '''
			CleanFichiersAnnexes(fichier_trt_acc)
			shutil.rmtree(rep_traitement)
			logfile.write("Suppression du repertoire de traitement " + rep_traitement, 0)

		else:
			''' on deplace le repertoire dans le repertoire d'erreur '''
			pass
			rep_dest = rep_traitement.replace(RepFluxTrt,RepFluxErreur)
			logfile.write("Deplacement du repertoire de traitement " + rep_traitement + " vers " + rep_dest, 0)
			if os.path.isfile(rep_dest):
				logfile.write("Dossier "+rep_dest+" deja présent : purge",0)
				os.remove(rep_dest)
			try:
				shutil.copytree(rep_traitement, rep_dest)
			except OSError:
				pass
			CleanFichiersAnnexes(fichier_trt_acc)
			# shutil.rmtree(rep_traitement)
		sys.exit(0)
	return


if __name__=='__main__':
	
	envfile      	= cheminConfig + os.sep + "env.tab"
	processfile   	= cheminConfig + os.sep + "process.tab"
	busfile		   	= cheminConfig + os.sep + "bus.tab"
	env 			= Loadfile(envfile)
	process 		= Loadfile(processfile)
	num_threads 	= int(process.get_element("NbProcess"))
	intervale 		= int(process.get_element("TScrut"))
	extension		= process.get_element("extension")
	buslist 		= LoadBusFile(busfile)

	''' traitement des autres repertoires de traitements '''
	RepFluxTrt		= env.get_element("RepFluxTrt")
	RepFluxErreur	= env.get_element("RepFluxErreur")
	RepArchives		= env.get_element("RepFluxArchives")
	RepLog			= env.get_element("RepLog")
	RepFluxInData	= env.get_element("RepFluxInTmp")

	''' MODE DEBUG '''
	DEBUG = 0
	if not os.path.isdir(RepFluxInData):
		os.makedirs(RepFluxInData, 0777)
	now 			= datetime.now()
	logfilename		= RepLog + "/composition_" + now.strftime('%Y-%m-%d') + ".log"
	logfile			= LogFile(logfilename)	
	""" boucle principale """
	while True:
		now = datetime.now()
		logfilename		= RepLog + "/composition_" + now.strftime('%Y-%m-%d') + ".log"
		if not os.path.isfile(logfilename):
			logfile			= LogFile(logfilename)	
		index = 0
		logfile.write("Demarrage traitement",0)
		for DirIn in buslist.inputdirs: # on parcourt la liste des repertoires d'entrée	
			if os.path.isdir(DirIn): # si le repertoire existe
				if DEBUG == 1:
					logfile.write("Debut scrutation repertoire entree : " + DirIn, 0)
				files = os.listdir(DirIn)
				DirOut = buslist.outputdirs[index]
				ArchiveDir = buslist.archivedirs[index]
				TemplateName = buslist.templatesname[index] # nom_maquette_editique
				FluxPrefixe = buslist.prefixesflux[index]
				nomCanal = buslist.listeecanaux[index]
				prefixeMaquette = buslist.prefixeMaquette[index]
				DirOutBis = buslist.outputdirsbis[index]
				for file in files:
					# logfile.write("Lecture fichier en entree: " + file, 0)
					fileinput = file
					filein = os.path.join(DirIn, file)
					taille1 = 0
					taille2 = 1
					while taille1!=taille2 :
						taille1 = os.path.getsize(filein)
						time.sleep(0.2)
						taille2 = os.path.getsize(filein)
					fileout = ArchiveDir
					if not os.path.isdir(fileout):
						os.makedirs(fileout)
					fileout = fileout + os.sep + file
					if (os.path.isfile(filein)) :
						if (os.path.splitext(filein)[1] != ".acc") and (os.path.splitext(filein)[1] != ".html") and (os.path.splitext(filein)[1] != ".config") and (os.path.splitext(filein)[1] != ".chorus"):
							try:
								logfile.write("Deplacement du fichier " + filein + " dans : " + fileout, 0)
								shutil.copy(filein, fileout)
							except Exception, e:
								logfile.write("probleme : " + str(e), 0)
							# Partie preparation flux a traiter avec gestion exclusion 
							try:
								logfile.write("Lecture des flux en entree: " + file, 0)
								# files = os.listdir(RepFluxInData)
								#on recupère uniquement les fichiers a traiter et non les fichiers annexe
								if (file.lower().endswith('.csu') or file.lower().endswith('.cou') or file.lower().endswith('.xml') or file.lower().endswith('.csv')):
									# pas besoin d'ouvrir le fichier pour conaitre la filliere
									DirOut = ""
									TemplateName = ""
									FluxPrefixe = ""
									outputdirs = ""
									prefixeMaquette = ""
									DirOutBis = ""
									idxCanal = 0
									for canal in buslist.listeecanaux:
										if buslist.inputdirs[idxCanal] == DirIn: # necessaire mais pas suffisant si plusieurs fillieres dans un meme dossier de scrutation
											if buslist.archivedirs[idxCanal] == ArchiveDir :
												DirOut = buslist.outputdirs[idxCanal]
												TemplateName = buslist.templatesname[idxCanal]
												FluxPrefixe = buslist.prefixesflux[idxCanal]
												prefixeMaquette = buslist.prefixeMaquette[idxCanal]
												DirOutBis = buslist.outputdirsbis[idxCanal]
												if FluxPrefixe == "oui" : # si plusieurs filieres dans le meme dossier il faut un prefixe sur le nom de fichier (cas papier)
													if canal == "AFP" or canal == "PDF" :
														if file.startswith(prefixeMaquette) :
															fileout = os.path.join(RepFluxInData, canal +"."+ TemplateName +"."+ file)
															try :
																logfile.write("Deplacement du fichier " + filein + " dans : " + fileout, 0)
																shutil.copy(filein, fileout)
																#creation du fichier d'accompagnement
																FichierAcc 			= os.path.splitext(fileout)[0] + ".acc"
																fileacc 			= open(FichierAcc, "w", 0)
																fileacc.write("OUTPUTDIR\t" 	+ DirOut + "\n")
																fileacc.write("OUTPUTDIRBIS\t" 	+ DirOutBis + "\n")
																if canal == "MARKETING" :
																	for file in os.listdir(DirIn) :
																		if os.path.splitext(file)[1] == ".html" :
																			my_file_html = os.path.join(DirIn,file)
																			fileacc.write("HTML\t" 	+ my_file_html + "\n")
																		if os.path.splitext(file)[1] == ".config" :
																			my_file_xml = os.path.join(DirIn,file)
																			fileacc.write("CONFIG\t" 	+ my_file_xml + "\n")
																fileacc.close()
																logfile.write("Création du fichier " + FichierAcc +"dans Dir:"+DirIn+ " Canal :" +str(idxCanal) + " Préfixe:" + prefixeMaquette, 0)
																#acc deja present dans archivage
															except Exception,e : 
																	logfile.write("probleme copie fichier %s " % (filein),0)	
												else : #prefixe = non (cas mail)
													if canal == "MAIL" or canal == "HYBRIDE" or canal == "PDF" :
														fileout = os.path.join(RepFluxInData, canal +"."+ TemplateName +"."+ file)
													else : # specifique marketing
														fileout = os.path.join(RepFluxInData, canal +"."+ TemplateName +"."+ file)
													try :
														logfile.write("Deplacement du fichier " + filein + " dans : " + fileout, 0)
														shutil.copy(filein, fileout)
														#creation du fichier d'accompagnement
														FichierAcc 			= os.path.splitext(fileout)[0] + ".acc"
														fileacc 			= open(FichierAcc, "w", 0)
														fileacc.write("OUTPUTDIR\t" 	+ DirOut + "\n")
														fileacc.write("OUTPUTDIRBIS\t" 	+ DirOutBis + "\n")
														if canal == "MARKETING" :
															for file in os.listdir(DirIn) :
																if os.path.splitext(file)[1] == ".html" :
																	my_file_html = os.path.join(DirIn,file)
																	fileacc.write("HTML\t" 	+ my_file_html + "\n")
																if os.path.splitext(file)[1] == ".config" :
																	my_file_xml = os.path.join(DirIn,file)
																	fileacc.write("CONFIG\t" 	+ my_file_xml + "\n")
														elif canal == "CHORUS" :
															for file in os.listdir(DirIn) :
																if os.path.splitext(file)[1] == ".chorus" :
																	my_file_xml = os.path.join(DirIn,file)
																	fileacc.write("XML\t" 	+ my_file_xml + "\n")
														fileacc.close()
														logfile.write("Création du fichier " + FichierAcc +" dans Dir:"+DirIn+ " Canal :" +str(idxCanal) + " Préfixe:" + prefixeMaquette, 0)
														#acc deja present dans archivage
													except Exception,e : 
														logfile.write("probleme copie fichier %s " % (filein),0)
												execute = True
										idxCanal = idxCanal + 1
							except Exception, e:
								logfile.write("Erreur lors du traitement du fichier : " + file, 0)
								logfile.write("     => " + str(e), 0)
								fileout = RepFluxErreur + os.sep + file
							''' suppression du fichier initial de base '''
							if (os.path.splitext(filein)[1] != ".html") and (os.path.splitext(filein)[1] != ".pdf"):
								os.remove(filein)
			index = index + 1
		files = os.listdir(RepFluxInData)
		''' creation de la file d'attente '''
		QueueExecution = Queue(maxsize=0)
		'''conversion de la liste de fichiers en tableau (on ne sélectionne pas les fichiers d'accompagnements '''
		myvars = []
		for file in files :
			if not (file.endswith('.html') or file.endswith('.acc')):
				myvars.append(file)
		i = 0
		''' compter les fichiers en attente '''
		nbFichierAttente = len(files)/2
		affichedebug("SEFAS | Reception fichier donnees | still alive | "  + str(nbFichierAttente) + " fichier(s) en attente", DEBUG)
		while i < num_threads and len(myvars) != 0:
			try:
				file = myvars[i]
				nom_complet = os.path.join(RepFluxInData, file)
				""" pour chaque fichier """
				if (os.path.isfile(nom_complet)):
					''' Test si la taille du fichier varie => fichier en cours de chargement '''
					taille1 = os.path.getsize(nom_complet)
					time.sleep(0.2)
					taille2 = os.path.getsize(nom_complet)
					if (taille1 != taille2):
						break
					''' Verification de l'utilisation des thread deja en cours '''
					ThreadDispo = False
					while not ThreadDispo and i < num_threads :
						fichierLock = RepFluxTrt + os.sep + "thread" + str(i+1) + ".lock"
						if os.path.isfile(fichierLock):
							''' thread non dispo '''
							i=i+1
						else:
							''' thread dispo '''
							ThreadDispo = True
							affichedebug("     Thread : "+ str(i+1) + " | Fichier : " + file +"", DEBUG)
							worker = Thread(target=do_stuff, args=(QueueExecution,i,file,nom_complet))
							worker.setDaemon(True)
							worker.start()
							QueueExecution.put(i)
							QueueExecution.join()
					# execution de la queue 		
			except IndexError:
				pass
			i = i+1	
		# SCRUTATION EN BOUCLE
		time.sleep(intervale)
		# sys.exit(0)
		if DEBUG == 1:
			logfile.write("Fin scrutation repertoire entree", 0)
	# FIN DE LA BOUCLE PRINCIPALE 
# FIN DU MAIN