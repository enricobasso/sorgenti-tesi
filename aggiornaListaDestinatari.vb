%REM
	Agent AGGIORNA LISTA DESTINATARI
%END REM
Option Public
Option Declare
Use "FunzioniComuni"
Use "funzionixmlls"

Sub Initialize
	Dim dbNotifiche As New NotesDatabase("", "")
	Dim viewConfNotifiche As NotesView
	Dim dcConf As NotesDocumentCollection
	Dim docConf As NotesDocument
	
	Dim flag As Boolean
	Dim pathFileXMLInput As String
	Dim errore As String
	Dim fileRes As String
	
	Dim iscritti List As String
	Dim disiscritti List As String
	
	Dim session As New NotesSession
	Dim dbDestinatari As NotesDatabase
	Dim viewDestinatari As NotesView
	Dim doc As NotesDocument
	
	flag = dbNotifiche.Open(RitornaServerAnno("NOTIFICHE", Year(Today)), RitornaDatabaseAnno("NOTIFICHE", Year(Today)))
	
	If flag Then
		Set viewConfNotifiche = dbNotifiche.Getview("(#SERVIZI AGENTE SCHEDULATO)")
		Set dcConf = viewConfNotifiche.Getalldocumentsbykey("Attivo", True)
		Set docConf = dcConf.Getfirstdocument()
		
		While Not(docConf Is Nothing)
			pathFileXMLInput = creaFileInputXMLWSC()
			fileRes = eseguiChiamataWSC(RitornaValore(docConf, "ChiaveWSListaDestinatari", 0), "DESTINATARI NOTIFICHE", "", pathFileXMLInput, Errore)
			
			If errore <> "" Then
				Print "Errore durante l'esecuzione della chiamata WS:" + RitornaValore(docConf, "ChiaveWSListaDestinatari", 0) + "."
			Else
				errore = leggiRispostaWS(fileRes, Iscritti, Disiscritti)
				
				If errore <> "" Then
					Print "Errore durante la lettura della risposta della chiamata WS: " + RitornaValore(docConf, "ChiaveWSListaDestinatari", 0) + "."
				Else
					Set dbDestinatari = session.Currentdatabase
					Set viewDestinatari = dbDestinatari.Getview("(#DESTINATARI)")
					
					' creo documenti iscritti
					ForAll iscritto In iscritti
						' controllo se il destinatario esiste già prima di inserirlo
						Set doc = viewDestinatari.Getdocumentbykey(iscritto, True)
						
						If  doc Is NOTHING Then
							Set doc = dbDestinatari.Createdocument()
							doc.destinatario = iscritto
							doc.nome_servizio = RitornaValore(docConf, "codice_comunicazione", 0)
							doc.data_iscrizione = Today - 1
							doc.form = "Destinatario"
							Call doc.Save(True, False)
							Set doc = Nothing
						End If
					End ForAll
					
					' elimino destinatari disiscritti
					ForAll disiscritto In disiscritti
						Set doc = viewDestinatari.Getdocumentbykey(disiscritto, true)
						doc.Remove(True)
					End ForAll
				End If
			End If
			
			Set docConf = dcConf.GetNextDocument(docConf)
		Wend
	Else
		Print "Impossibile aprire il db SIGED NOTIFICHE."
	End If
	
	
End Sub

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
	Function creaFileInputXML
%END REM
Function creaFileInputXMLWSC()
	Dim rootXML As XMLTag
	Dim fileInputXML As New XMLFile(Cstr("GetListaIscrizioni"))
	Dim tagXML As XMLTag
	Dim oggetto As XMLTag
	Dim testo As XMLTag
	Dim pathName As String
	
	Set rootXML = fileInputXML.documentTag
	
	' Aggiunge data
	Set tagXML = rootXML.AppendChild("data")
	tagXML.InnerText = Format$(Today - 1, "yyyy-mm-dd")
	
	' Scrittura file su disco
	pathname = scriviFileXmlSuDisco(fileInputXML, "input_notifica.xml")
	
	creaFileInputXMLWSC = pathName
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
	Function leggiRispostaWS
%END REM
Function leggiRispostaWS(fileRes As String, iscritti List As String, disiscritti List As String)
	Dim docXML As NotesDOMDocumentNode
	Set docXML = apriFileXML(fileRes)
	
	If Not docXML.isnull Then			
		Dim rootelement As NotesDOMElementNode
		Dim itemList As NotesDOMNodeList
		Dim i As Integer
		
		Set rootElement = docXML.DocumentElement
		Set itemList = rootElement.GetElementsByTagName("iscritto")
		
		If (itemList.Numberofentries > 0) Then
			For i = 1 To itemlist.Numberofentries
				iscritti(i) = itemList.Getitem(i).Firstchild.Nodevalue
			Next
		End If
		
		Set itemList = rootElement.GetElementsByTagName("disiscritto")
		If (itemList.Numberofentries > 0) Then
			For i = 1 To itemlist.Numberofentries
				disiscritti(i) = itemList.Getitem(i).Firstchild.Nodevalue
			Next
		End If
	Else
		leggiRispostaWS = "Impossibile aprire file XML di risposta."
	End If
End Function
