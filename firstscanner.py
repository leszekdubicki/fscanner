# -*- coding: cp1250 -*-
import os, sys
import os.path
import win32com.client
import win32api
from win32api import *
import pythoncom
import pyPdf
import copy
import fnmatch
import tempfile
import ConfigParser
import time
import subprocess

USER = win32api.GetUserName()
#USER_EX = win32api.GetUserNameEx(3)Q
USER_EX = ""
pathname = os.path.dirname(sys.argv[0])        
SCRIPTPATH = os.path.abspath(pathname)	#sciezka do katalogu gdzie jest skrypt

#sciezka do pdfsam-a:
SAMPATH = SCRIPTPATH+"\\pdfsam\\bin"
#SAMFILE = SCRIPTPATH+"\\pdfsam\\bin\\run-console.bat"
SAMFILE = "run-console.bat"
#print SAMPATH


def process_files(DOCUMENT, PAGE, PATH = None):
	global SAMPATH, SAMFILE
	if PATH == None:
		PATH = os.path.dirname(DOCUMENT)
	elif PATH[-1] == "\\":
		PATH = PATH[:-1]
	#import subprocess
	print("\n\nrozdzielenie plikow...")
	os.chdir(SAMPATH)
	#OUTPUT1 = subprocess.call([SAMFILE,"-f",PDF_NAME,"-o",PATH,"-s","SPLIT","-n","1","-overwrite","split"]) #nie dziala po skompilowaniu :(
	#print OUTPUT1
	#win32api.ShellExecute(0, "open", SAMFILE, " -f "+"\""+PDF_NAME+"\""+" -o "+"\""+PATH+"\""+" -s SPLIT -n 1 -overwrite split" , SAMPATH, 1)
	os.system(SAMFILE+" -f "+"\""+DOCUMENT+"\""+" -o "+"\""+PATH+"\""+" -s SPLIT -n 1 -overwrite split")
	#raw_input("........................jesli pdfsam zakonczyl rozdzielac nacisnij enter!")
	print("\n\nwstawienie pierwszej strony...")
	PDF_NAME_1 = PATH+"\\1_"+os.path.basename(DOCUMENT)
	PDF_NAME_2 = PATH+"\\2_"+os.path.basename(DOCUMENT)
	os.system(SAMFILE+" -f "+"\""+PAGE+"\""+" -f "+"\""+PDF_NAME_2+"\""+" -o "+"\""+PATH + "\\" + os.path.basename(DOCUMENT)+"\""+" -overwrite concat")
	#win32api.ShellExecute(0, "open", SAMFILE, " -f \""+FILE+"\" -f "+"\""+PDF_NAME_2+"\""+" -o "+"\""+PDF_NAME+"\""+" -overwrite concat", SAMPATH, 1) #nie dziala po skompilowaniu :(
	#OUTPUT2 = subprocess.call([SAMFILE,"-f",FILE,"-f",PDF_NAME_2,"-o",PDF_NAME,"-overwrite","concat"])
	#print OUTPUT2
	#raw_input("pykkkkkk")
	
	#usuniecie plikow tymczasowych utworzonych przez pdfsam:
	if os.path.isfile(PDF_NAME_1) and os.path.basename(PDF_NAME_1)[:2]=="1_":
		os.remove(PDF_NAME_1)
	if os.path.isfile(PDF_NAME_2) and os.path.basename(PDF_NAME_2)[:2]=="2_":
		os.remove(PDF_NAME_2)


def check_write(filename): 
	if os.path.isfile(filename):
		directory = os.path.dirname(filename)
	else:
		directory = filename
	if os.path.isdir(directory):
		try: 
			fd, fn = tempfile.mkstemp(dir=directory) 
			os.close(fd) 
			os.remove(fn) 
			return True 
		except: 
			return False 
	else:
		 return False

def isfileok(F, FirstPage = False):
	#funkcja zwraca w. logiczną mówiącą czy można użyć pliku i ew. komunikaty błędów
	KOM = u"Błąd: "; OK = True
	#warunek formatu - pdf dla pierwszej str. i pdf albo doc dla drugiej:
	if os.path.isdir(F):
		KOM+=u"wpisana ścieżka jest katalogiem, "
		OK = False
	elif not os.path.isfile(F):
		KOM+=u"plik nie istnieje, "
		OK = False
	elif not fnmatch.fnmatch(F,"*.pdf"):
		KOM+=u"plik w formacie innym niż pdf, "
		OK = False
	else:
		if not os.path.isfile(F):
			KOM+=u"plik nie istnieje, "
			OK = False
		else:
			if not FirstPage and not check_write(F):
				KOM+=u"plik tylko do odczytu, "
				OK = False
			n = pages(F)
			if FirstPage and n>1:
				KOM+=u"skan musi mieć jedną stronę, ma "+str(n)+", "
				OK = False
			elif not FirstPage and n == 1:
				KOM+=u"skan powinien mieć więcej niż jedną stronę, "
				OK = False
	if KOM == u"Błąd: ":
		KOM = ""
	return OK, KOM[:-2]
					
def getword():
		pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
		#myWord = win32com.client.DispatchEx('Word.Application')
		myWord = win32com.client.Dispatch('Word.Application')
		#myWord.Visible = True
		#from pyPdf import PdfFileWriter, PdfFileReader
		#import shutil
		
		#scan = PdfFileReader(file())
		#doc = PdfFileReader(file())
		#page1 = scan.getPage(0)
		#output = PdfFileWriter()
	
		#odczytanie otwartego dokumentu worda:
		if myWord.Documents.Count == 1:
			WORDDOC = myWord.Documents.Item(1)
			NAME = WORDDOC.Name
			PATH = WORDDOC.Path
			return PATH+"\\"+NAME
		else:
			return ""
		
def saveword(File = None):
		pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
		#myWord = win32com.client.DispatchEx('Word.Application')
		myWord = win32com.client.Dispatch('Word.Application')
		myWord.Visible = True
		#from pyPdf import PdfFileWriter, PdfFileReader
		#import shutil
		
		#scan = PdfFileReader(file())
		#doc = PdfFileReader(file())
		#page1 = scan.getPage(0)
		#output = PdfFileWriter()
	
		#odczytanie otwartego dokumentu worda:
		if File == "":
			File = None
		if not File == None:
			WORDDOC = myWord.Documents.Open(File)
		else:
			WORDDOC = myWord.Documents.Item(1)
		NAME = WORDDOC.Name
		PATH = WORDDOC.Path
		PDF_NAME = PATH+"\\"+NAME[:NAME.rindex(".")]+".pdf"
		PDF_NAME_2 = PATH+"\\2_"+NAME[:NAME.rindex(".")]+".pdf"
		PDF_NAME_1 = PATH+"\\1_"+NAME[:NAME.rindex(".")]+".pdf"
		#print PDF_NAME
		#raw_input()
		MESSAGE = ""
		if not check_write(os.path.dirname(PDF_NAME)):
			MESSAGE = "Katalog tylko do odczytu"
			print MESSAGE
			return False, MESSAGE
		elif not os.path.isfile(PDF_NAME):
			print("zapisanie jako pdf...")
			WORDDOC.ExportAsFixedFormat(OutputFileName=PDF_NAME, ExportFormat=17, OpenAfterExport=False, OptimizeFor=0, Range=0, From=1, To=1, Item=0, IncludeDocProps=True, KeepIRM=True, CreateBookmarks=0, DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False)
			return True, PDF_NAME
		else:
			MESSAGE = "plik pdf juz istnieje!"
			print(MESSAGE)
			return False, MESSAGE
def getallworddocs():
	pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
	myWord = win32com.client.Dispatch('Word.Application')
	DOCS = []
	for i in range(1, myWord.Documents.Count+1):
		WORDDOC = myWord.Documents.Item(i)
		DOCS.append(WORDDOC.Path+"\\"+WORDDOC.Name)
	return DOCS
	
#utworzenie maila:
def mail(Subject, Body):
	o = win32com.client.Dispatch("Outlook.Application")

	Msg = o.CreateItem(0)
	#Msg.To = recipient
	Msg.Subject = Subject
	Msg.Body = Body
	Msg.Display()
	return Msg
def ConvertTime(SentOn):
	#zwraca czas wyslania podany przez wlasciwosc MailItem.SentOn w sek. od EPOCH
	#_%Y-%m-%d_%H-%M
	date = str(SentOn)
	date2 = time.strptime(date, "%m/%d/%y %H:%M:%S")
	T = time.mktime(date2)
	return T

class Mail:
	def __init__(self, Subject, Body, Filename, Folder):
		self._subject = Subject
		self._body = Body
		self._filename = Filename
		self._folder = Folder
		try:
			self._o = win32com.client.Dispatch("Outlook.Application")
		except:
			self._o = None
	def create(self):
		if not self._o == None:
			Msg = self._o.CreateItem(0)
			#Msg.To = recipient
			Msg.Subject = self._subject
			Msg.Body = self._body
			self._creationtime = time.time()
			Msg.Display()
	def path(self):
		PATH = self._folder + "\\" + self._filename
		PATH = PATH.replace("/","\\")
		PATH = PATH.replace("\\\\","\\")
		PATH+=".msg"
		return PATH
	def find(self):
		#zwraca True jak mail jest w wyslanych, a false jak nie
		if self._o == None:
			return False
		result = False
		N = self._o.GetNamespace("MAPI")
		myFolder = N.GetDefaultFolder(5)
		for i in range(len(myFolder.Items),len(myFolder.Items)-20,-1):
			if i==len(myFolder.Items):
				try:
					MM = myFolder.Items[i] #na wypadek jakby ostatni index byl rowny dlugosci listy a nie -1
				except:
					continue
			else:
				MM = myFolder.Items[i]
			date = ConvertTime(MM.SentOn)
			if (MM.Subject == self._subject) and (date - self._creationtime > 0):
				self._sentMail = MM
				result = True
				break
		return result
	def save(self):
		#poszukanie maila w wyslanych i zapisanie na dysku jako _filename w katalogu _path
		if self._o == None:
			return 4 #kod bledu 4
		result = 1 #z zalozenia nie znaleziono maila
		if os.path.isfile(self.path()):
			#plik o nazwie docelowej juz istnieje!
			result = 7
		elif not self.find():
			result = 11
		elif self.find() and (check_write(os.path.dirname(self.path()))):
			MM = self._sentMail
			date = ConvertTime(MM.SentOn)
			if (MM.Subject == self._subject) and (time.time() - date > 0) and (time.time() - date < 60*60) and (os.path.isdir(os.path.dirname(self.path()))):
				MM.SaveAs(self.path())
				result = 0
		elif not check_write(os.path.dirname(self.path())):
			result = 5
		if result == 0 and not os.path.isfile(self.path()):
			result+=2
		return result
	def explore(self):
		#otwiera okno exploratora tam gdzie zapisal maila
		subprocess.Popen(r'explorer /n,/select,'+self.path())
		

def pages(F):
	#zwraca ilosc stron pliku o scierzce F
	FILEOBJ = file(F,"rb")
	PDF = pyPdf.PdfFileReader(FILEOBJ)
	n = PDF.getNumPages()
	FILEOBJ.close()
	return n
	

def get_args(ARGS, Word = True):
	FILE2 = None
	try:
		FILE = ARGS[0]
		if len(ARGS)==2:
			FILE2 = ARGS[1]
	except:
		raw_input("Brak pliku, nacisnij enter i sprobuj ponownie :D")
		#sys.exit()
		return None, None
		#FILE = "c:\\temp\DOC120711.pdf"
	#print FILE
	
	if FILE2==None:
		if Word:
			print("zapisanie jako pdf...")
			PDF_NAME = saveword()
	else:
		#weryfikacja, czy mamy do czynienia z dwoma pdf-ami:
		if FILE[-4:].lower()==".pdf" and FILE2[-4:].lower()==".pdf":
			print(u"oba pliki maja rozszerzenie pdf, sprawdzanie ilosci stron")
		else:
			print(u"oba pliki muszą byc w formacie pdf.\nNaciśnij enter żeby zakończyć")
			#raw_input()
			return None, None
			#sys.exit()
		PDFPAGES = pages(FILE); PDF2PAGES = pages(FILE2)
		if PDFPAGES>1 and PDF2PAGES==1:
			#zamiana kolejności
			PDF_NAME = copy.deepcopy(FILE)
			FILE = FILE2
		elif PDF2PAGES>1 and PDFPAGES==1:
			PDF_NAME = copy.deepcopy(FILE2)
		else:
			print(u"jeden z plikow pdf musi mieć jedna stronę, drugi więcej.\nNaciśnij enter żeby zakończyć")
			return None, None
			#raw_input()
			#sys.exit()
	return PDF_NAME, FILE

class config:
	#definiuje konfiguracje
	def configfile(self, ARG0 = None):
		#ARG0 to pierwszy argument sys.argv
		SELFPATH = os.path.dirname(ARG0)
		if SELFPATH in [None,""]:
			SELFPATH= os.path.dirname(sys.argv[0])
		if SELFPATH.strip() == "":
			configfile = "config.txt"
		else:
			configfile = SELFPATH+"\\config.txt"
		return configfile
	def __init__(self, ARG0 = None):
		#ARG0 to pierwszy argument sys.argv
		configfile = self.configfile(ARG0)
		self._config = ConfigParser.ConfigParser()
		self._config.read(configfile)
		self._configfile = configfile
	def has(self, option, option2 = ""):
		if not option2 == "":
			section = option
			option = option2
		else:
			section = "mail"
		if self._config.has_section(section):
			return self._config.has_option(section, option)
		else:
			return False
	def get(self, option, option2 = ""):
		#pobiera z pliku conf opcje option, jesli option2 jest podane to option jest section :)
		if not option2 == "":
			section = option
			option = option2
		else:
			section = "mail"
		if self._config.has_section(section) and self._config.has_option(section,option):
			return self._config.get(section, option)
		else:
			return ""
	def Subject(self):
		return self.get("subject")
	def FileName(self):
		return self.get("filename")
	def MailPath(self):
		return self.get("mailpath")
	def Body(self):
		B = ""
		for i in range(0,100):
			if self.has("body"+str(i)):
				B+=self.get("body"+str(i))+"\n"
		return B
	def set(self, section, option, string = ""):
		#zapisuje opcję do pliku konfiguracyjnego
		if string == "":
			string = option
			option = section
			section = "mail"
		if not self._config.has_section(section):
			self._config.add_section(section)
		self._config.set(section, option, string)
	def set_Subject(self, Subject):
		self.set("subject",Subject)
	def set_FileName(self, FileName):
		self.set("filename", FileName)
	def set_Body(self, Body):
		BB = Body.splitlines()
		BBB = [""]
		j = 0
		for i in range(0,len(BB)):
			if not BBB[j] == "" and BB[i].strip() == "":
				BBB[j] = BBB[j][:-1]
				j+=1
				BBB.append("\n")
				j+=1
				BBB.append("")
					
			else:
				BBB[j]+=BB[i]
				BBB[j]+="\n"
		BBB[j] = BBB[j][:-1] #pozbycie się ostatniego znaku pystego wiersza
		for i in range(0, len(BBB)):
			self.set("body"+str(i),BBB[i])
		for i in range(len(BBB),100):
			if self._config.has_option("mail","body"+str(i)):
				self._config.remove_option("mail","body"+str(i))
	def set_MailPath(self, MailPath):
		self.set("mailpath", MailPath)
	def save(self):
		CFILE = open(self._configfile,"w")
		self._config.write(CFILE)
		CFILE.close()
def doccode(FNAME):
	#zwraca kod dokumentu
	IND1 = FNAME.index('_')
	IND2 = FNAME.index('_',IND1+1)
	return FNAME[:IND2]
def zeros(String, n):
	while len(String)<n:
		String = "0"+String
	return String
def plus1(String):
	if String.isdigit():
		return zeros(str(int(String)+1), len(String))
	else:
		return String
def minus1(String):
	if String.isdigit():
		return zeros(str(int(String)-1), len(String))
	else:
		return String
	
def resolve(string, PATH, VERSION = None, DATE = None, DATABATCH = None, VAR = None , USER = None):
	#zamienia zmienne w postaci "<ZMIENNA>" (zawarte w słowniku DICT) na odpowiednie slowa
	if USER == None: USER = win32api.GetUserName()
	if DATE == None: DATE = time.strftime("%Y-%m-%d")
	if VERSION == None: VERSION = ""
	VERSION1 = minus1(VERSION)
	if DATABATCH == None:
		DATABATCH = ""
	if VAR == None:
		VAR = ""
	
	DOCNAME = os.path.basename(PATH); DOCNAME = DOCNAME[:DOCNAME.rindex(".")]
	DOCCODE = doccode(DOCNAME)
	DICT = {"<DOCFOLDER>":os.path.dirname(PATH),
		"<FILENAME>":os.path.basename(PATH),
		"<DOCNAME>":DOCNAME,
		"<DOCPATH>":PATH,
		"<USER>":USER,
		"<VERSION>": VERSION,
		"<VERSION-1>": VERSION1,
		"<DATE>":DATE,
		"<DATEPREFIX>":DATABATCH,
		"<DOCCODE>":DOCCODE,
		"<VAR>": VAR
	}
	for VAR in DICT:
		if VAR in string:
			string2 = string.replace(VAR,DICT[VAR])
			string = string2
	return string
	
if __name__ == "__main__":
	
	#okreslenie pliku podanego jako argument:
	ARGS = sys.argv
	PDF_NAME, FILE = get_args(ARGS[1:])
	if PDF_NAME == None or FILE == None:
		sys.exit()
	#print PDF_NAME
	#print FILE
	CONFIG = config(sys.argv[0])
	
	process_files(PDF_NAME, FILE)
	
	DATE = ""
	VER = ""
	if "<DATE>" in CONFIG.Subject()+CONFIG.Body()+CONFIG.FileName()+CONFIG.MailPath():
		DATE = raw_input("Podaj Date: ")
	if "<VERSION>" in CONFIG.Subject()+CONFIG.Body()+CONFIG.FileName()+CONFIG.MailPath():
		VER = raw_input("Podaj wersje: ")
	M = Mail(resolve(CONFIG.Subject(),PDF_NAME, VER, DATE, USER), resolve(CONFIG.Body(),PDF_NAME, VER, DATE, USER), resolve(CONFIG.FileName(),PDF_NAME, VER, DATE, USER), resolve(CONFIG.MailPath(),PDF_NAME, VER, DATE, USER))
 	M.create()
	raw_input("Po wyslaniu maila nacisnij enter aby go zapisac")
	iter = 0
	while not M.find():
		raw_input("Brak maila w wyslanych, po wyslaniu nacisnij enter aby zapisac maila")
		iter+=1
		if iter>10:
			print("Mail nie zapisany, przerywam.....")
			break
	if M.find():
		M.save()
		M.explore()
	#os.system(SAMFILE+" -f \""+FILE+"\" -f \""+PDF_NAME+"\" -o \""+PDF_NAME+"\" concat")
	#print(SAMFILE+" -f "+FILE+" -f "+PDF_NAME+" -o "+PDF_NAME+" -overwrite concat")
	#os.system("\""+SAMFILE+" -f "+FILE+" -f "+PDF_NAME+" -o "+PDF_NAME+" -overwrite concat"+"\"")
	
	#przyklad z http://www.codeproject.com/KB/system/newbiespawn.aspx
	#SHELLEXECUTEINFO
	#ShExecInfo = {0};
	#ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	#ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
	#ShExecInfo.hwnd = NULL;
	#ShExecInfo.lpVerb = NULL;
	#ShExecInfo.lpFile = SAMFILE;		
	#ShExecInfo.lpParameters = " -f "+"\""+PDF_NAME+"\""+" -o "+"\""+PATH+"\""+" -s SPLIT -n 1 -overwrite split";	
	#ShExecInfo.lpDirectory = NULL;
	#ShExecInfo.nShow = SW_SHOW;
	#ShExecInfo.hInstApp = NULL;	
	#ShellExecuteEx(&ShExecInfo);
	#WaitForSingleObject(ShExecInfo.hProcess,INFINITE);
