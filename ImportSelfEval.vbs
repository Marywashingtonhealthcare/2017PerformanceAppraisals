Sub Click(Source As Button)
    Dim s As New notessession
    Dim ws As New NotesUIWorkspace
    Dim db As NotesDatabase
    Dim uidoc As NotesUIDocument
    Dim doc As NotesDocument
    Dim evaldoc As NotesDocument
    Dim dc As NotesDocumentCollection

    Set db = s.CurrentDatabase
    Set uidoc = ws.CurrentDocument
    Set doc = uidoc.Document
    Set dc = doc.Responses

    Dim searchunid As String
    'Dim searchtext As String
    searchunid = doc.selfEvalUNID(0)
    'searchtext = {([UNID] = searchunid)}
    'searchtext = {([Form] = "selfEvalAssoc") And ([flagEval] = "Complete")}
    'Call dc.FTSearch(searchtext, 0)
    'Set evaldoc = dc.GetFirstDocument
    Set evaldoc = db.GetDocumentByUNID(searchunid)'is this the best way to find this? Don't UNID's change under some circumstances?

    uidoc.EditMode = True

    If evaldoc.flagEval(0) <> "Complete" Then
        Messagebox "The Self Eval is not in Complete Status. Make sure the Associate has clicked the Submit button on their Self Eval before continuing...", 16, "Error"
        Exit Sub
    Else
        Print "Attempting to import Self Eval..."
    End If


    doc.icaOverall1 = evaldoc.icaOverall1
    doc.icaAssoc = evaldoc.icaAssoc
    doc.depSR1 = evaldoc.depSR1
    doc.depSR2 = evaldoc.depSR2
    doc.depSR3 = evaldoc.depSR3
    doc.depSR4 = evaldoc.depSR4
    doc.depSR5 = evaldoc.depSR5
    doc.depFd1Assoc = evaldoc.depFd1Assoc
    doc.depFd2Assoc = evaldoc.depFd2Assoc
    doc.depFd3Assoc = evaldoc.depFd3Assoc
    doc.depFd4Assoc = evaldoc.depFd4Assoc
    doc.depFd5Assoc = evaldoc.depFd5Assoc
    doc.fldComments = evaldoc.fldComments
    doc.lmRat1 = evaldoc.lmRat1
    doc.lmRat2 = evaldoc.lmRat2
    doc.lmRat3 = evaldoc.lmRat3
    doc.lmRat4 = evaldoc.lmRat4
    doc.lmRat5 = evaldoc.lmRat5
    doc.lmRat6 = evaldoc.lmRat6
    doc.lmRat7 = evaldoc.lmRat7
    doc.lmRat8 = evaldoc.lmRat8
    doc.lmRat9 = evaldoc.lmRat9
    doc.lmRat10 = evaldoc.lmRat10
    doc.lmRat11 = evaldoc.lmRat11
    doc.lmRat12 = evaldoc.lmRat12
    doc.lmRat13 = evaldoc.lmRat13
    doc.lmRat14 = evaldoc.lmRat14
    doc.lmRat15 = evaldoc.lmRat15
    doc.lmComm1 = evaldoc.lmComm1
    doc.lmComm2 = evaldoc.lmComm2
    doc.lmComm3 = evaldoc.lmComm3
    doc.lmComm4 = evaldoc.lmComm4
    doc.lmComm5 = evaldoc.lmComm5
    doc.lmComm6 = evaldoc.lmComm6
    doc.lmComm7 = evaldoc.lmComm7
    doc.lmComm8 = evaldoc.lmComm8
    doc.lmComm9 = evaldoc.lmComm9
    doc.lmComm10 = evaldoc.lmComm10
    doc.lmComm11 = evaldoc.lmComm11
    doc.lmComm12 = evaldoc.lmComm12
    doc.lmComm13 = evaldoc.lmComm13
    doc.lmComm14 = evaldoc.lmComm14
    doc.lmComm15 = evaldoc.lmComm15

    Call doc.ReplaceItemValue("fldSave", "Yes")
    'Call doc.Save(False,False)
    Call uidoc.Save
    Call uidoc.refresh

    Msgbox "SUCCESS: Import Action Run" & Chr(13) & Chr(13) & "THIS IS NOT AN ERROR MESSAGE!" & Chr(13) & Chr(13) & "Note: If you think data is missing, please check the self eval document to see what the Associate has written. If you think there is an import problem, please contact HR."

End Sub