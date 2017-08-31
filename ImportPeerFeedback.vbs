Sub Click(Source As Button)
    Dim s As New notessession
    Dim ws As New NotesUIWorkspace
    Dim db As NotesDatabase
    Dim uidoc As NotesUIDocument
    Dim doc As NotesDocument
    Dim dc As NotesDocumentCollection
    Dim peerdoc As NotesDocument

    Dim commFlag As Variant

    Set db = s.CurrentDatabase
    Set uidoc = ws.CurrentDocument
    Set doc = uidoc.Document
    Set dc = doc.Responses
    Set peerdoc = dc.GetFirstDocument

    uidoc.EditMode = True

    While Not (peerdoc Is Nothing)
        If (peerdoc.fldReviewer(0) <> "") Then 'This is what we use to differentiate between self evals and peer review docs.
            commFlag = peerdoc.GetItemValue("commentsFlag") 'set value of commFlag from commentsFlag field on peer review (commentsFlag holds field name text for each peer comment field)
            Call uidoc.FieldAppendText(commFlag(0), "REVIEW DATE:   " & peerdoc.fldRevDate(0))
            Call uidoc.FieldAppendText(commFlag(0), " " & Chr(13) & Chr(10) & " ")
            Call uidoc.FieldAppendText(commFlag(0), " " & Chr(13) & Chr(10) & " ")
            Call uidoc.FieldAppendText(commFlag(0), "RATINGS:")'The ratings also need to stay in some capacity.
            Call uidoc.FieldAppendText(commFlag(0), " " & Chr(13) & Chr(10) & " ")
            Call uidoc.FieldAppendText(commFlag(0), "Quality: " & peerdoc.fldQuality(0))
            Call uidoc.FieldAppendText(commFlag(0), " " & Chr(13) & Chr(10) & " ")
            Call uidoc.FieldAppendText(commFlag(0), "Service: " & peerdoc.fldService(0))
            Call uidoc.FieldAppendText(commFlag(0), " " & Chr(13) & Chr(10) & " ")
            Call uidoc.FieldAppendText(commFlag(0), "Teamwork: " & peerdoc.fldService(0))
            Call uidoc.FieldAppendText(commFlag(0), " " & Chr(13) & Chr(10) & " ")
            Call uidoc.FieldAppendText(commFlag(0), " " & Chr(13) & Chr(10) & " ")
            Call uidoc.FieldAppendText(commFlag(0), "COMMENTS: ")
            Call uidoc.FieldAppendText(commFlag(0), " " & Chr(13) & Chr(10) & " ")
            Forall vals In peerdoc.fldcomments 'Forall loop because the text field is encoded as an array.
                Call uidoc.FieldAppendText(commFlag(0), vals)
                Call uidoc.FieldAppendText(commFlag(0), " " & Chr(13) & Chr(10) & " ")
            End Forall
        End If
        Set peerdoc = dc.GetNextDocument(peerdoc)
    Wend

    Msgbox ("Import operation complete." & Chr(13) & Chr(13) & "If no text showed up in any Peer Comments fields, it may be because:" & Chr(13) & Chr(13) & _
    "* The user did not click Submit on their Peer Review form, and it is still in Draft Status" & Chr(13) & Chr(13) & "You can check by going to the To Do List and see " & _ 
    "what their peer review status is." & Chr(13) & Chr(13) & "If the peer review is not in completed status, it will not import.")

End Sub