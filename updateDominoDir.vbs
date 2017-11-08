%REM
	Agent a-commToShortName
	Created Oct 11, 2017 by Shawn Sillman/MWH
	Description: Comments for Agent
%END REM

Option Public
Option Declare

Sub Initialize
	Dim session As New NotesSession
	Dim db As NotesDatabase
	Dim doc As NotesDocument
	Dim view As NotesView
	Dim filePath As String  'This will be the path to the text document that gets appended when we test this thing out for the first time.
	Dim counter As Integer  'Initialize the counter for the on-screen looping feedback in the status bar
	Dim fileNum As Integer
	Dim fileName As String
	Dim item As NotesItem

	Set db = session.CurrentDatabase
	Set view = db.GetView("PerfAppImport")
	Set doc = view.GetFirstDocument()

	counter = 1
	'fileNum% = FreeFile()
	'fileName$ = "E:\crystal\domino-dir.csv"

	'Open fileName$ For Output As fileNum%
	'Write #fileNum%, "ShortName", "Comment", "User"
	'Here you create a while loop that loops through all of the documents in the domino directory and copies the contents of the Comment field into the ShortName field if the length of the string in the Comment field is exactly six characters (the length of a common EMP#).

	While Not(doc Is Nothing)
		If(Len(doc.Comment(0)) = 6) Then
			'Write #fileNum%, doc.ShortName(0), doc.Comment(0), "Will be changed"
			Set item = doc.ReplaceItemValue("ShortName", doc.Comment(0)) 'copy Comment to ShortName field
            Call doc.Save(True, False)'save the document
			Print "Found " & counter & " documents..."
			counter = counter + 1
		End If'write if statement
			Set doc = view.GetNextDocument(doc)'Get a handle to the next document.
		Wend

		'Close fileNum%
		Print "Finished looping through the documents, Mr. Sillman...Have a nice day."
	
End Sub
