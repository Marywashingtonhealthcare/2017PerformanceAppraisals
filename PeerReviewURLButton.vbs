Sub Click(Source As Button)
	Dim s As New NotesSession
	Dim ws As New NotesUIWorkspace
	
	Dim db As NotesDatabase
	Set db = s.CurrentDatabase
	
	Dim uidoc As NotesUIDocument
	Dim doc As NotesDocument
	Set uidoc = ws.CurrentDocument
	Set doc = uidoc.Document
	uidoc.EditMode = True
	
	Dim text1 As String
	Dim text2 As String
	
	
    text1 = "At Home:" + Chr(13) + Chr(13) + "http://inotes.medicorp.org/databases/human%20resources/mwhc-pe.nsf/0/" & doc.universalID & "?EditDocument"
    text2 = "At Work:" + Chr(13) + Chr(13) + "http://dapps1.mhs.org/databases/human%20resources/mwhc-pe.nsf/0/" & doc.universalID & "?EditDocument"
	
	Call uidoc.FieldSetText("URLs",  text1 + Chr(13) + Chr(13) + text2)
	
	Call uidoc.GoToField("URLs")
	Call uidoc.SelectAll
	Call uidoc.Copy
	
	Msgbox text1 + Chr(13) + Chr(13) + text2, 0+64+0+1, "URLs copied to clipboard"
	
	
End Sub