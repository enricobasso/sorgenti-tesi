Option Public
Option Declare
Use "FunzioniComuni"
Use "SIGED_XML_ENGINE"

Sub Initialize
	
	' Agent: Invio Notifiche
	
	Dim Session As New NotesSession
	Dim dbNotifiche As NotesDatabase
	Dim viewCodaNotifiche As NotesView
	Dim viewConfigurazione As NotesView
	Dim dcNotifiche As NotesDocumentCollection
	Dim docNotifica As NotesDocument
	Dim errore As String
	Dim chiaviRicercaConf(1) As String
	Dim docConf As NotesDocument
	Dim flag As Boolean
	
	Dim dbOrigineNotifica As New NotesDatabase("", "")
	Dim viewUNICOD As NotesView
	Dim docOrigineNotifica As NotesDocument
	Dim nomeDbOrigine As String
	Dim annoOrigine As String
	
	Set dbNotifiche = session.CurrentDatabase
	Set viewCodaNotifiche = dbNotifiche.GetView("(#CODA NOTIFICHE)")
	Set viewConfigurazione = dbNotifiche.GetView("(#CONFIGURAZIONE)")
	Set dcNotifiche = viewCodaNotifiche.GetAllDocumentsByKey("IN CODA", True)
	
	Set docNotifica = dcNotifiche.GetFirstDocument()
	While Not(docNotifica Is Nothing)
		' recupero configurazione per il docNotifica
		If chiaviRicercaConf(0) <> RitornaValore(docNotifica, "codice_comunicazione", 0) Or chiaviRicercaConf(0) <> StrToken(RitornaValore(docNotifica, "Chiave_Origine", 0), "{}", 2) Then	
			chiaviRicercaConf(0) = RitornaValore(docNotifica, "codice_comunicazione", 0)
			chiaviRicercaConf(1) = StrToken(RitornaValore(docNotifica, "Chiave_Origine", 0), "{}", 2)
			
			Set docConf = viewConfigurazione.GetDocumentByKey(chiaviRicercaConf(), True)
		End If
		
		' inizializzo campi mancanti in docNotifica
		docNotifica.tipo_notifica = docConf.tipo_notifica
		docNotifica.chiave_procedura = docConf.chiave_procedura
		docNotifica.template_xml = docConf.template_xml
		docNotifica.chiaveWS = docConf.chiaveWS
		docNotifica.chiaveWSGetProfiloUtente = docConf.chiaveWSGetProfiloUtente
		docNotifica.DbDestinatari = docConf.DbDestinatari
		Call docNotifica.Save(True, false)
	
		' recupero documento origine della notifica
		nomeDbOrigine = StrToken(RitornaValore(docNotifica, "Chiave_Origine", 0), "{}", 2)
		annoOrigine = StrToken(RitornaValore(docNotifica, "Chiave_Origine", 0), "{}", 1)
		flag = dbOrigineNotifica.Open(RitornaServerAnno(nomeDbOrigine, annoOrigine), RitornaDatabaseAnno(nomeDbOrigine, annoOrigine))
		
		If flag Then
			Set viewUNICOD = dbOrigineNotifica.Getview("(#UNICOD)")
			Set docOrigineNotifica = viewUNICOD.Getdocumentbykey(StrToken(RitornaValore(docNotifica, "Chiave_Origine", 0), "{}", 3), True)
			
			' creazione file xml per WSC
			Call creaFileInputXMLNotifica(Docnotifica, Docoriginenotifica)
			' esecuzione richiesta HTTP
			Call creaChiamateWSC(docNotifica)
			
			Docnotifica.status = "PROCESSATO"
			Call Docnotifica.Save(True, False)
			
		Else
			Print "Errore apertura db di origine della notifica."
		End If
		
		Set docNotifica = dcNotifiche.GetNextDocument(docNotifica)
	Wend
End Sub

%REM
	Function creaFileInputXMLWSC
%END REM
Function creaFileInputXMLWSC(destinatario As String, oggettoNotifica As String, testoNotifica As String, errore As String)
	Dim rootXML As XMLTag
	Dim fileInputXML As New XMLFile(CStr("nuovo_messaggio"))
	Dim tagXML As XMLTag
	Dim oggetto As XMLTag
	Dim testo As XMLTag
	Dim pathName As String
	
	Set rootXML = fileInputXML.documentTag
	
	' Aggiunge destinatario
	Set tagXML = rootXML.AppendChild("codice_fiscale")
	tagXML.InnerText = destinatario
	
	' Aggiunge messaggio
	Set tagXML = rootXML.AppendChild("corpo_messaggio")
	Set oggetto = tagXML.AppendChild("oggetto")
	Set testo = tagXML.AppendChild("testo")
	
	oggetto.InnerText = oggettoNotifica
	testo.InnerText = testoNotifica
	
	' Scrittura file su disco
	pathname = scriviFileXmlSuDisco(fileInputXML, "input_notifica.xml")
	
	creaFileInputXMLWSC = pathName
End Function

%REM
	Function creaChiamateWSC
%END REM
Function creaChiamateWSC(docNotifica As NotesDocument)
	Dim errore As String
	' Percorso dove salvare XML da scheda notifica
	Dim pathXML As String
	' Percorso dell'XML input per WSC
	Dim inputXML As String
	Dim fileXML As NotesRichTextItem
	' File di risposta
	Dim fileRes As String
	
	' Dati per invio messaggio
	Dim destinatario As String
	Dim oggettoNotifica As String
	Dim testoNotifica As String
	
	Dim destFlag As Boolean
	
	' Download file input xml in locale
	Set fileXML = docNotifica.Getfirstitem("input_xml")
	
	ForAll att In fileXML.EmbeddedObjects
		If att.Type = EMBED_ATTACHMENT Then
			pathXML = CalcolaNomeFile("input.xml", True)
			Call att.ExtractFile(pathXML)
		End If
	End ForAll
	
	Call getContenutoNotifica(pathXML, destinatario, oggettoNotifica, testoNotifica, errore)
	If Not(errore = "") Then
		Print errore
		Exit Function
	End If
	
	If IsNull(destinatario) Then
		errore = "Questa notifica non ha destinatario."
		Print errore
		docNotifica.errore = errore		
		Exit function
	End If
	
	
	destFlag = checkDestinatarioDaLista(destinatario, RitornaValore(docNotifica, "codice_comunicazione", 0))
	
	If Not(Destflag) Then
		destflag = getProfiloUtente(destinatario, docNotifica)
	End If
	
	If Destflag then
		inputXML = creaFileInputXMLWSC(destinatario, oggettoNotifica, testoNotifica, errore)
		fileRes = eseguiChiamataWSC(RitornaValore(docNotifica, "ChiaveWS", 0), StrToken(RitornaValore(docNotifica, "Chiave_Origine", 0),"{}",2), StrToken(RitornaValore(docNotifica, "Chiave_Origine", 0),"{}",3), inputXML, errore)
		
		If errore <> "" Then
			Print errore
			docNotifica.errore = "Errore durante invio del messaggio."
		Else
			Dim flagMsg As Boolean
			Dim idMsg As String
			Dim erroreMsg As String
			
			idMsg = controllaRispostaMsg(fileRes, erroreMsg)
			
			If idMsg <> "" Then
				Call docNotifica.GetFirstItem("idMessaggio").AppendToTextList(idMsg)
			Else
				Call docNotifica.GetFirstItem("errore").AppendToTextList(erroreMsg)
			End If
		End If
	Else
		Call docNotifica.GetFirstItem("errore").AppendToTextList(destinatario + "#" + "Destinatario non abilitato a invio notifiche IO.italia.it")
	End If
End Function

%REM
	Function creaFileInputNotifica
	Prende il template xml collegato alla notifica e crea il file xml di input per il WSC.
%END REM
Function creaFileInputXMLNotifica(docNotifica As NotesDocument, docOrigine As NotesDocument)
	Dim xmlInput As XMLFile
	Dim pathName As String
	Dim inputItem As NotesRichTextItem
	
	Set xmlInput = engine_simple_start(RitornaValore(docNotifica, "template_xml", 0), docOrigine, Nothing)
	
	If IsNull(xmlInput) Then
		Print "Errore durante la creazione del file xml di input della notifica tramite template xml."
	Else
		pathname = scriviFileXmlSuDisco(xmlInput, "input_notifica.xml")
		Set inputItem = docNotifica.getFirstItem("input_xml")
		Call inputItem.EmbedObject(EMBED_ATTACHMENT, "", pathName)
		Call docNotifica.Save(True, False)
	End If
End Function

%REM
	Function controllaRispostaMsg
	Legge file xml di risposta dal WSC.
%END REM
Function controllaRispostaMsg(fileRes As String, erroreMsg As String)
	Dim docXML As NotesDOMDocumentNode
	Set docXML = apriFileXML(fileRes)
	
	If Not docXML.isnull Then			
		Dim rootelement As NotesDOMElementNode
		Set rootElement = docXML.DocumentElement

		If (rootElement.GetElementsByTagName("id_messaggio").Numberofentries = 1) Then
			controllaRispostaMsg = rootElement.GetElementsByTagName("id_messaggio").Getitem(1).Firstchild.Nodevalue
		Else
			erroreMSG = rootElement.GetElementsByTagName("status").Getitem(1).Firstchild.Nodevalue 
		End If
	End If
End Function

%REM
	Function getProfiloUtente
	Esegue chiamata WSC per verificare se il destinatario è abilitato all'invio di notifiche.
%END REM
Function getProfiloUtente(utente As String, docNotifica As NotesDocument) 
	Dim rootXML As XMLTag
	Dim fileInputXML As New XMLFile(CStr("getProfiloUtente"))
	Dim tagXML As XMLTag
	Dim fileRes As String
	Dim errore As String
	
	Set rootXML = fileInputXML.documentTag
	
	' Aggiunge CF destinatario
	Set tagXML = rootXML.AppendChild("codice_fiscale")
	tagXML.InnerText = utente
	
	' Scrittura file su disco
	Dim pathName As String
	Dim wFile As String
	
	pathname = CalcolaNomeFile("inputWSC.xml", True)
	
	wFile = Dir$(pathname)
	If wFile <> "" Then
		Kill pathname
	End If
	
	Call fileInputXML.WriteFile(pathname, False)

	fileRes = eseguiChiamataWSC(RitornaValore(docNotifica, "ChiaveWSGetProfiloUtente", 0), StrToken(RitornaValore(docNotifica, "Chiave_Origine", 0),"{}",2), StrToken(RitornaValore(docNotifica, "Chiave_Origine", 0),"{}",3), pathName, errore)
	
	If errore <> "" Then
		Print errore
		Exit function
	End If

	' Lettura file risultato
	Dim session As New NotesSession
	Dim inputStream As NotesStream
	Dim domParser As NotesDOMParser
	Dim docXML As NotesDOMDocumentNode
	
	Set inputStream = session.CreateStream
	inputStream.Open(fileRes)
	If inputStream.Bytes = 0 Then
		Call inputStream.Close
		errore = "Errore apertura del file"
		Exit Function
	End If
	
	Set domParser = session.CreateDOMParser(inputStream)
	domParser.InputValidationOption = 0
	domParser.Parse
	Set docXML = domParser.Document
	
	If Not docXML.isnull Then			
		Dim rootelement As NotesDOMElementNode
		Set rootElement = docXML.DocumentElement		

		If (rootElement.GetElementsByTagName("mittente_abilitato").Numberofentries = 0) Then
			getProfiloUtente = False
		Else
			getProfiloUtente = rootElement.GetElementsByTagName("mittente_abilitato").Getitem(1).Firstchild.Nodevalue 
		End If
	End If
End Function

%REM
	Function getContenutoNotifica
%END REM
Function getContenutoNotifica(nomeFile As String, destinatario As String, oggettoNotifica As String, testoNotifica As String, errore As String) 
	Dim docXML As NotesDOMDocumentNode
	Set docXML = apriFileXML(nomeFile)
	
	If Not docXML.isnull Then			
		Dim rootelement As NotesDOMElementNode
		Dim itemList As NotesDOMNodeList
		Dim i As Integer
		
		Set rootElement = docXML.DocumentElement		
		Set itemList = rootElement.GetElementsByTagName("destinatario")
		
		If (itemlist.Numberofentries > 0) Then
			destinatario = itemList.Getitem(i).Firstchild.Nodevalue
			
		End If

		If (rootElement.GetElementsByTagName("oggetto").Numberofentries = 0) Then
			errore = "Nessun oggetto specificato nel file di input xml."
			Print errore
			Exit Function
		Else
			oggettoNotifica = rootElement.GetElementsByTagName("oggetto").Getitem(1).Firstchild.Nodevalue
		End If
		
		If (rootElement.GetElementsByTagName("testo").Numberofentries = 0) Then
			errore = "Nessun testo specificato nel file di input xml."
			Print errore
			Exit Function
		Else
			testoNotifica = rootElement.GetElementsByTagName("testo").Getitem(1).Firstchild.Nodevalue
		End If
	End If
End Function

%REM
	Function apriFileXML
%END REM
Function apriFileXML(nomeFile As String)
	Dim session As New NotesSession
	Dim inputStream As NotesStream
	Dim domParser As NotesDOMParser
	Dim errore As String
	
	Set inputStream = session.CreateStream
	inputStream.Open(nomefile)
	If inputStream.Bytes = 0 Then
		Call inputStream.Close
		errore = "Errore apertura del file: " + nomeFile
		Exit Function
	End If
	
	Set domParser = session.CreateDOMParser(inputStream)
	domParser.InputValidationOption = 0
	domParser.Parse

	Set apriFileXML = domParser.Document
End Function

%REM
	Function scriviFileXmlSuDisco
%END REM
Function scriviFileXmlSuDisco(fileXml As XMLFile, nomeFile As String) 
	Dim pathName As String
	Dim wFile As String
	
	pathname = CalcolaNomeFile(nomeFile, True)
	
	wFile = Dir$(pathname)
	If wFile <> "" Then
		Kill pathname
	End If
	
	Call fileXml.WriteFile(pathname, False)
	
	scriviFileXmlSuDisco = pathname
End Function

%REM
	Function checkDestinatario
	Controlla nel db DESTINATARI se il destinatario è abilitato all'invio.
%END REM
Function checkDestinatarioDaLista(destinatario As String, codiceComunicazione As String)
	Dim dbDestinatari As New NotesDatabase("", "")
	Dim viewDestinatari As NotesView
	Dim dcDestinatari As NotesDocumentCollection
	Dim chiaviRicerca(1) As String
	Dim flag As Boolean
	Dim hash As String
	
	flag = dbDestinatari.Open(RitornaServerAnno("DESTINATARI NOTIFICHE", Year(Today())), RitornaDatabaseAnno("DESTINATARI NOTIFICHE", Year(Today())))
	
	If flag Then
		chiaviRicerca(0) = LCase$(calcolaSHA256stringa(UCase$(destinatario)))
		chiaviRicerca(1) = codiceComunicazione
		
		Set viewDestinatari = dbDestinatari.GetView("(#DESTINATARI)")
		Set dcDestinatari = viewDestinatari.GetAllDocumentsByKey(chiaviRicerca(), True)
		
		If dcDestinatari.Count = 1 Then
			checkDestinatarioDaLista = True
		Else			
			checkDestinatarioDaLista = False
		End If
	Else
		Print "Impossibile controllare il destinatario della notifica."
		checkDestinatarioDaLista = False
	End If
	
End Function
