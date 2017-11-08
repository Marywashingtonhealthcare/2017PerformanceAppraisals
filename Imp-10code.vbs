Option Public
Sub Initialize
    'Call clean sub right off the bat to clean up all existing documents in the view
    Call clean

    Dim xlFilename As String
 '// This is the name of the Excel file that will be imported
    xlFilename = "E:\crystal\PerfAppList"

    Dim session As New NotesSession
    Dim db As NotesDatabase
    Dim view As NotesView
    Dim viewSL As NotesView
    Dim doc As NotesDocument
    Dim docSL As NotesDocument
    Dim row,xx, counter As Integer
    Dim written As Integer
    Dim help,test As Variant

    Set db = session.CurrentDatabase
    Set doc = New NotesDocument(db)

    Set viewSL = db.GetView("MgrListSideLoad")
    'This will loop through the sideload view to add documents manually from the MgrListSideLoad view.
    Set docSL = viewSL.GetFirstDocument
    While Not (docSL Is Nothing)
        Set doc = db.Createdocument 'creates the new document
        doc.form = "frmMgrList"
        doc.ID = docSL.ID(0)
        doc.First = docSL.First(0)
        doc.Middle = docSL.Middle(0)
        doc.Last = docSL.Last(0)
        doc.Pos = docSL.Pos(0)
        doc.Department = docSL.Department(0)
        doc.NotesName = docSL.NotesName(0)
        doc.Email = docSL.Email(0)
        doc.JobCode = docSL.JobCode(0)
        Call doc.save(True,True)
        Set docSL = viewSL.GetNextDocument(docSL)
    Wend

'//---------------------------------------------------------------------------------------

'// Next we connect to Excel and open the file. Then start pulling over the records.
    Dim Excel As Variant
    Dim xlWorkbook As Variant
    Dim xlSheet As Variant
    Dim testJC As Integer
    Dim testPos As String
    Print "Connecting to Excel..."
    Set Excel = CreateObject( "Excel.Application" )
    Excel.Visible = False '// Don't display the Excel window
    Print "Opening " & xlFilename & "..."
    Excel.Workbooks.Open xlFilename '// Open the Excel file
    Set xlWorkbook = Excel.ActiveWorkbook
    Excel.Sheets("Sheet1").Select
    Set xlSheet = xlWorkbook.ActiveSheet

        '// Cycle through the rows of the Excel file, pulling the data over to Notes

        Records:
        row = 1 '// These integers intialize to zero anyway
        written = 0
        xx = 5

        Print "DBREPORT: Starting import from Excel file" & xlFilename & "......."
        counter = 0
        Do While True '// We use a test a few lines down to exit the loop
            Finish:
            With xlSheet
                row = row + 1 'starting out on row 2 of the Excel file. Excluding the header.

                Set view= db.GetView("MgrList")

                doc.Form = "frmMgrList"

                Do While xx=5
	                testJC = CInt(.Cells(row, 16).Value)
	                testPos = .Cells(row, 9).Value
	                Do While (testJC > 2101) And (testPos <> "Medical Technologist")
                        row = row + 1
		                testJC = CInt(.Cells(row, 16).Value)
		                testPos = .Cells(row, 9).Value
	                Loop
                    Set doc = db.CreateDocument '// Create a new doc
                    doc.Form = "frmMgrList"
                    doc.ID= .Cells( row, 1 ).Value
                    temp = .Cells( row, 1 ).Value
                    doc.FIRST = .Cells(row,5).Value

                    If .Cells(row,5).Value="" Then
                        GoTo Done
                    End If

                    doc.MIDDLE = .Cells(row,6).Value
                    doc.LAST = .Cells(row,7).Value
                    doc.POS = .Cells(row,9).Value
                    doc.DEPARTMENT = .Cells(row,11).Value
                    doc.JOBCODE = .Cells(row,16).Value 'Added for Job Code
                    
                    row= row+1

                    'lookup notes Name And managers Name From nab
                    Dim dbNab As NotesDatabase
                    Set dbNab = session.GetDatabase( "Dapps1/MWH", "names.nsf" )
                    Set viewNab = dbNab.GetView ("PerfAppImport" )
                    Set docNab = viewNab.GetDocumentByKey (CStr(temp))

                    If Not (docNab Is Nothing) Then
                        Dim commonname As New NotesName(docNab.FullName(0))
                        Dim emailAddr As String
                        strName = commonname.Abbreviated
                        emailAddr = docNab.InternetAddress(0)
                        doc.NotesName = strName
                        doc.Email = emailAddr
                    End If

                    Call doc.Save( True, True ) '// Save the new doc
                    counter = counter + 1
                    written = written + 1
                    test = doc.ID
                    If counter > 15000 Then
                        GoTo Done
                    Else
                    End If
                Loop

            End With
        Loop
        Return

        Done:
        xlWorkbook.Close False '// Close the Excel file without saving (we made no changes)
        Excel.Quit '// Close Excel
        Set Excel = Nothing '// Free the memory that we'd used
        Print "DBREPORT: FINISHED import from Excel file..."
        
        'In this section, we will remove all entries where the Job Code is above a 1900, which will leave only managers and supervisors.------------------------------
        
        
        
        
        

'Now we clean duplicates-----------------------------------------------------------------------------
        Print "Cleaning Duplicates"
        Dim doc2 As NotesDocument
        Dim doc3 As NotesDocument
        Dim view2 As NotesView
        Dim s As New NotesSession
        Dim num As Integer
        num = 0


        Set view2 = db.GetView("MgrList")
        Set doc3 = view2.getfirstdocument
        Set doc2 = view2.getnextdocument(doc3)

        While Not doc3 Is Nothing
            If Not doc2 Is Nothing Then
                If doc3.ID(0) = doc2.ID(0) Then
                    Call doc2.remove(True)
                ElseIf doc2.POS(0) = "Medical Technologist" Then
                	Call doc2.remove(True)
                ElseIf doc2.POS(0) = "Cytotechnologist" Then
	                Call doc2.remove(True)
                ElseIf doc2.POS(0) = "Senior Technologist" Then
	                Call doc2.remove(True)
                ElseIf doc2.POS(0) = "Sleep Lab Team Leader" Then
	                Call doc2.remove(True)
                ElseIf doc2.POS(0) = "Lead Therapist" Then
	                Call doc2.remove(True)
                ElseIf doc2.POS(0) = "Senior Tech - Hematology" Then
	                Call doc2.remove(True)
                ElseIf doc2.POS(0) = "Point of Care Technologist" Then
	                Call doc2.remove(True)
                ElseIf doc2.POS(0) = "CSR Technician, Lead" Then
	                Call doc2.remove(True)
                ElseIf doc2.POS(0) = "Displaced Associate - Manager" Then
	                Call doc2.remove(True)
                Else
                    Set doc3 = doc2
                End If
                num = num + 1
                Set doc2 = view2.getnextdocument(doc3)
                On Error GoTo EN
            Else
                Set doc3 = doc2
                On Error GoTo EN
            End If
        Wend
        EN:
        Print "Agent Finished"
    End Sub



'Clean subroutine...this runs at the beginning of the main sub and removes everything in the view in preparation for the import 
    Sub clean
        Dim session As New NotesSession
        Dim db As NotesDatabase
        Dim view As NotesView
        Dim collection As NotesViewEntryCollection
        Dim entry As NotesViewEntry
        Dim doc As NotesDocument
        Set db = session.CurrentDatabase
        Set view = db.GetView("MgrList")
        Set collection = view.AllEntries
        Set entry = collection.GetFirstEntry()
        While Not(entry Is Nothing)
            Set doc = entry.Document
            doc.Remove(True)
            Set entry = collection.GetNextEntry(entry)
        Wend
    End Sub