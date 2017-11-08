Sub Click(Source As Button)
	Dim s As New NotesSession	
	Dim ws As New NotesUIWorkspace
	Dim db As NotesDatabase
	Set db = s.CurrentDatabase
	Dim uidoc As NotesUIDocument
	Dim doc As NotesDocument
	Dim revdoc As NotesDocument
	Dim url As Variant
	Dim email As Variant
	Dim maildoc As NotesDocument
	Dim setauthor As NotesItem
	Dim publicacc As NotesItem
	Dim reader As NotesItem
	Dim SaveOptions As NotesItem
	Dim caseResult As String
	Dim arrayString As String
	Dim summary As Variant
	Dim mbresult As Integer
	
	Set uidoc = ws.CurrentDocument
	Set doc = uidoc.document
	uidoc.EditMode = True
	
	mbresult = Msgbox ("Important: Did you set up access in the Setup tab and Save as Draft yet?", 4 + 32 + 0 + 0, "Access + Save Check")
	
	If mbresult = 6 And uidoc.IsNewDoc = False Then
		'do nothing
	Else
		Msgbox ("If you do not set up access and Save as Draft now, users you add later will not be able to access the Peer Evals. Please go back and set up access first, then Save as Draft.")
		Exit Sub
	End If
	
	Dim peer1 As New Peer(doc.peer1(0), doc.peer1Email(0), doc.peer1Setup(0), doc.peer1flag(0))
	Dim peer2 As New Peer(doc.peer2(0), doc.peer2Email(0), doc.peer2Setup(0), doc.peer2flag(0))
	Dim peer3 As New Peer(doc.peer3(0), doc.peer3Email(0), doc.peer3Setup(0), doc.peer3flag(0))
	Dim peer4 As New Peer(doc.peer4(0), doc.peer4Email(0), doc.peer4Setup(0), doc.peer4flag(0))
	Dim peer5 As New Peer(doc.peer5(0), doc.peer5Email(0), doc.peer5Setup(0), doc.peer5flag(0))
	'created 5 peer objects using the custom class defined in declarations with the New function
	
	If uidoc.Document.MgrSecFlag(0) = "" And uidoc.Document.MgrNameSec(0) = "" Then
		Messagebox "Please fill out the Manager Conducting Appraisal section before continuing...", 16, "You Forgot Something!"
		Exit Sub
	Else
	     'do nothing...
	End If
	
	Dim peerArray(4) As Variant
	Set peerArray(0) = peer1
	Set peerArray(1) = peer2
	Set peerArray(2) = peer3
	Set peerArray(3) = peer4
	Set peerArray(4) = peer5
	
	'*********************************************************(Start Readers Field Creation Code)**************
	'the following lines basically turn the 4 author fields into arrays and then join all 3 arrays into
	'readerArrayFinal, which is used to create the Readers field for each peer request.
	MgrName = doc.GetItemValue("Authors")
	MgrNameSec = doc.GetItemValue("Authors2")
	
	Dim readerArray(5) As Variant
	readerArray(0) = "Administrators"
	readerArray(1) = "LocalDomainServers"
	readerArray(2) = "HR PA Admins"
	readerArray(3) = "HR PA All Access"
	readerArray(4) = MgrName(0)
	readerArray(5) = MgrNameSec(0)
	
	Dim Auth3 As Variant
	Auth3 = doc.GetItemValue("Authors3")
	
	Dim Auth4 As Variant
	Auth4 = doc.GetItemValue("Authors4")
	
	Dim hrread As Variant
	hrread = doc.GetItemValue("HRReaders")
	
	Dim readerArrayMid As Variant
	readerArrayMid = Arrayappend(readerArray, Auth3)
	
	Dim readerArrayMid2 As Variant
	readerArrayMid2 = Arrayappend(readerArrayMid, Auth4)
	
	Dim readerArrayFinal As Variant
	readerArrayFinal = Arrayappend(readerArrayMid2, hrread)
	'*********************************************************(End Readers Field Creation Code)**************
	
	
	'We don't need the next 2 lines anymore...they refer to the old way of doing things.
	'Dim item As NotesItem
	'Set item = doc.GetFirstItem("fldPeerEmail")
	Dim summaryArray() As String
	Dim counter As Integer
	counter = -1
	Dim counter2 As Integer
	counter2 = 1
	Dim fieldstr As String
	Dim emailstr As String
	Dim emailresult As Variant
	
	Forall x In peerArray
		fieldstr = "peer" + Cstr(counter2) + "flag" 'this creates the field name reference - A string
		emailstr = "peer" + Cstr(counter2) + "Email"  'this creates the field name reference for email address.
		emailresult = uidoc.FieldGetText(emailstr)
		commstr = "peer" + Cstr(counter2) + "Comments"
		
		If x.peerName <> "" And x.peerEmail <> "" And x.peerSetup <> "" And x.peerflag = x.peerEmail Then
			caseResult = "3"  'Run the rest of the code downstream in the Forall loop...
		Elseif x.peerName <> "" And x.peerEmail <> "" And x.peerSetup <> "" Then
			caseResult = "4" 'This has already been sent. Throw an error stating they already sent this and it won't be sent again.
		Elseif  x.peerName = "" And x.peerEmail <> "" And x.peerSetup <> "" Then
			caseResult = "2" 'Throw an error, with a message stating that they missed something and this won't be sent.
		Elseif x.peerName = "" And x.peerEmail = "" And x.peerSetup <> "" Then
			caseResult = "2"
		Elseif x.peerName <> "" And x.peerEmail = "" And x.peerSetup = "" Then
			caseResult = "2"
		Elseif x.peerName = "" And x.peerEmail <> "" And x.peerSetup = "" Then
			caseResult = "2"
		Elseif x.peerName <> "" And x.peerEmail = "" And x.peerSetup <> "" Then
			caseResult = "2"
		Elseif x.peerName <> "" And x.peerEmail <> "" And x.peerSetup = "" Then
			caseResult = "2"
		Else
			caseResult = "1" 'Do nothing...go on to the next iteration of the loop.
		End If
		
		'Select case will now run code depending upon the result of the if statement block above.
		
		Select Case caseResult
		Case "1":
			' do nothing
		Case "2":
			Msgbox "A peer review will not be sent if all 3 fields have not been filled in. The Peer Review for " & x.peerName & " (" & x.peerEmail & ") will not be sent until you fill in all the required fields."
		Case "3":
			Msgbox "You have already sent a Peer Review to " & x.peerEmail & ". Another will not be sent."
		Case "4":
			'code to build array of sent to addresses
			Redim Preserve summaryArray(counter + 1)
			summaryArray(counter + 1) = x.peerEmail
			counter = counter + 1
			
			Set revdoc = db.CreateDocument
			revdoc.form = "peerReview17"
			revdoc.flagEval = "New" 'Find out what this does.
			revdoc.FirstName = doc.FirstName
			revdoc.LastName = doc.LastName
			revdoc.fldReviewer = x.peerName
			revdoc.peerEmail = emailresult
			revdoc.commentsFlag = commstr
			Set setauthor = New NotesItem(revdoc, "Author", "[Web]", AUTHORS)
			Set publicacc = New Notesitem(revdoc, "$PublicAccess", "1")
			Set reader = New NotesItem(revdoc, "Readers1", readerArrayFinal, READERS)'This is the state of Readers1 at the time of sending. It will not take into account any subsequent changes.
			Set saveoptions = New NotesItem(revdoc, "SaveOptions","1")
			Call revdoc.MakeResponse(doc)
			Call uidoc.FieldSetText(fieldstr, x.peerEmail)
			Call revdoc.save(False,False)
			'At this point, the peer review document has been created and saved.
			
			
			revdoc.docUrl = "http://inotes.medicorp.org/databases/human%20resources/mwhc-pe.nsf/0/" & revdoc.universalID & "?EditDocument"
			revdoc.docUrlint = "http://dapps1.mhs.org/databases/human%20resources/mwhc-pe.nsf/0/" & revdoc.universalID & "?EditDocument"
			
			url = revdoc.docUrl
			urlint = revdoc.docUrlint
			email = x.peerEmail
			
			'Now we construct the email that will go out to the reviewer.
			Set maildoc = New NotesDocument(db)
			Dim richtext As New NotesRichTextItem(maildoc, "Body")
			Dim richStyle As NotesRichTextStyle
			Set richStyle = s.CreateRichTextStyle
			Dim subtext As String
			subtext = revdoc.FirstName(0) + " " + revdoc.LastName(0) 
			maildoc.form = "Memo"
			maildoc.SendTo = email
			maildoc.Subject = "Peer Review request for " + subtext
			'body text
			Call richtext.AppendText("A Manager / Supevisor has requested that you fill out a Peer Review for " & subtext & " with the following criteria:")
			Call richtext.AddNewline(2)
			richStyle.Bold = True
			Call richtext.AppendStyle(richStyle)
			Call richtext.AppendText(x.peerSetup)
			Call richtext.AddNewline(2)
			richStyle.Bold = False
			Call richtext.AppendStyle(richStyle)
			Call richtext.AppendText("Please fill out the Peer Review web form by clicking one of the links below:")
			Call richtext.AddNewline(2)
			Call richtext.AppendText("If you are accessing this while AT WORK:")
			Call richtext.AddNewline(2)
			Call richtext.AppendText(urlint(0))
			Call richtext.AddNewline(2)
			Call richtext.AppendText("If you are accessing this while AT HOME:")
			Call richtext.AddNewline(2)
			Call richtext.AppendText(url(0))
			Call richtext.AddNewline(2)
			richStyle.Bold = True
			Call richtext.AppendStyle(richStyle)
			Call richtext.AppendText("PLEASE DO NOT DELETE THIS EMAIL: If you need to access this Peer Review again, you will need the link in this email.")
			
			Call maildoc.Send(False)
		End Select
		counter2 = counter2 + 1 'increment the counter that takes care of the peerXflag
	End Forall
	
	'code constructs the contents of final msgbox detailing who got emails sent to them
	'This will throw an error 200 (uninitialized array) if button is pressed with no data in any fields. If it becomes a problem, check for uninitialized array (will need to create a function)	
	If IsArrayEmpty(summaryArray) Then
		Msgbox("No Peer Reviews Sent.")
	Else
		summary = summaryArray
		arrayString = ""
		Forall x In summary
			If arrayString = "" Then
				arrayString = x
			Else
				arrayString = arrayString & Chr(13) & x
			End If
		End Forall
		Msgbox("Send Summary:" & Chr(13) & Chr(13) & "Emails Successfully sent to:" & Chr(13) & Chr(13) & arrayString & Chr(13) & Chr(13) & _
		"Recipients will receive a web link in their email to complete their Peer Review.")
	End If
	
	
	
End Sub