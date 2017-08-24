Sub Click(Source As Button)
    Dim s As New notessession
    Dim ws As New NotesUIWorkspace
    Dim db As NotesDatabase
    Dim uidoc As NotesUIDocument
    Dim doc As NotesDocument
    Dim selfEval As NotesDocument
    Dim dc As NotesDocumentCollection
    Dim askme As Integer
    Dim address As String
    Dim findSelfEvalUnid As String
    Dim FirstName As String
    Dim LastName As String

    'Set EnvVar's
    Set db = s.CurrentDatabase
    Set uidoc = ws.CurrentDocument
    Set doc = uidoc.Document

    ' Get the UNID to tie the parent doc to the sibling eval doc
    findSelfEvalUnid = doc.selfEvalUNID(0)

    ' Get the first and last name of the associate for email subject
    FirstName = doc.FirstName(0)
    LastName = doc.LastName(0)

    'Validate email address to make sure it's valid.
    address = uidoc.FieldGetText("fldEmail")
    
    If ValidateNotesAddr(address) = True Then
        Print "Notes Address Validated..."
    Elseif ValidateEmail(address) = True Then
        Print "Email Address Validated..."
    Else
        Msgbox "The Email Address in the Email Address field didn't pass validation. Please verify that you have the correct address and try again.", 0 + 16 + 0 + 0, "Invalid Email Address!"
        Call uidoc.GotoField("fldEmail")
        Exit Sub
    End If

    ' User verification to see if they truly want to rest the eval doc
    askme = ws.Prompt (PROMPT_YESNO,
    "Update/Changes", "Are you sure you would like to reset the goals and competencies on the linked associates self evaluation document with the changes made to this document?  Click ""Yes"" to Proceed or ""No"" to cancel.")
    If askme = 1 Then
        'Get the self eval doc based on the UNID
        Set selfEval = db.GetDocumentByUNID(findSelfEvalUnid)'
        uidoc.EditMode = True
        
        ' Set the fields on the self eval document with the parent doc's field values
        Call selfEval.ReplaceItemValue("depComp1",doc.depComp1)
        Call selfEval.ReplaceItemValue("depComp2",doc.depComp2)
        Call selfEval.ReplaceItemValue("depComp3",doc.depComp3)
        Call selfEval.ReplaceItemValue("depComp4",doc.depComp4)
        Call selfEval.ReplaceItemValue("depComp5",doc.depComp5)
        Call selfEval.ReplaceItemValue("lmGoal1",doc.LM1)
        Call selfEval.ReplaceItemValue("lmGoal2",doc.LM6)
        Call selfEval.ReplaceItemValue("lmGoal3",doc.LM11)
        Call selfEval.ReplaceItemValue("lmGoal4",doc.LM16)
        Call selfEval.ReplaceItemValue("lmGoal5",doc.LM21)
        Call selfEval.ReplaceItemValue("lmGoal6",doc.LM26)
        Call selfEval.ReplaceItemValue("lmGoal7",doc.LM31)
        Call selfEval.ReplaceItemValue("lmGoal8",doc.LM36)
        Call selfEval.ReplaceItemValue("lmGoal9",doc.LM41)
        Call selfEval.ReplaceItemValue("lmGoal10",doc.LM46)
        Call selfEval.ReplaceItemValue("lmGoal11",doc.LM51)
        Call selfEval.ReplaceItemValue("lmGoal12",doc.LM56)
        Call selfEval.ReplaceItemValue("lmGoal13",doc.LM61)
        Call selfEval.ReplaceItemValue("lmGoal14",doc.LM66)
        Call selfEval.ReplaceItemValue("lmGoal15",doc.LM71)
        Call selfEval.ReplaceItemValue("flagEval","Draft")
        Call selfEval.ReplaceItemValue("$PublicAccess","1")

        'The following code adds the Goals field data to the evaluation *********
        Dim ctractgoal, ctrginfo, ctrpill, ctrgood, ctrbetter, ctrbest, ctrwt As Integer 'dim all counters
        Dim actgoal, ginfo, pill, good, better, best, wt As String
        Dim factgoal, fpill, fgood, fbetter, fbest, fwt As Variant
        Dim fginfo As NotesItem
        Dim strpill, strgood, strbetter, strbest, strwt As String
        Dim strginfo As String
        ctractgoal = 1 'actual goal text
        ctrginfo = 1 'populate goal info field
        ctrpill = 2 'actual pillar text
        ctrgood = 3 'actual good text
        ctrbetter = 4 'actual better text
        ctrbest = 5 'actual best text
        ctrwt = 1 'actual weight text

        While ctractgoal < 72
            'build the field names as text
            actgoal = "LM" & Cstr(ctractgoal)
            ginfo = "gInfo" & Cstr(ctrginfo)
            pill = "LM" & Cstr(ctrpill) 'this is now strategic priority
            good = "LM" & Cstr(ctrgood) 'repurposed from "thresh"
            better = "LM" & Cstr(ctrbetter) 'repurposed from "stretch"
            best = "LM" & Cstr(ctrbest) 'repurposed from "actual"
            wt = "lmwt" & Cstr(ctrwt)
            'get values of fields on main document (these are arrays that must be converted to strings before they can be used)
            factgoal = doc.GetItemValue(actgoal)
            fpill = doc.GetItemValue(pill)
            fgood = doc.GetItemValue(good)
            fbetter = doc.GetItemValue(better)
            fbest = doc.GetItemValue(best)
            fwt = doc.GetItemValue(wt)
            'convert field values to usable strings using join function
            strpill = Join(fpill, Chr(13))
            strgood = Join(fgood, Chr(13))
            strbetter = Join(fbetter, Chr(13))
            strbest = Join(fbest, Chr(13))
            strwt = Join(fwt, Chr(13))
            'concatenate the strings together into a variable and set the field value on the child doc.
        strginfo = "Strategic Priority: " & strpill & Chr(13) & Chr(13) & "Good: " & strgood & Chr(13) & Chr(13) & "Better: " & strbetter & Chr(13) & Chr(13) & "Best: " & strbest & Chr(13) & Chr(13)  & "Weight: " & strwt
            Set fginfo = selfEval.ReplaceItemValue(ginfo, strginfo)
            'now increment each counter
            ctractgoal = ctractgoal + 5
            ctrginfo = ctrginfo + 1
            ctrpill = ctrpill + 5
            ctrgood = ctrgood + 5
            ctrbetter = ctrbetter + 5
            ctrbest = ctrbest + 5
            ctrwt = ctrwt + 1
        Wend

'**********************************************************(End Goals field additions)****


' By pass save validation on parent document
        Call doc.ReplaceItemValue("fldSave", "Yes")

           'Save uidoc with values that have been updated by the user on the parent doc
        Call uidoc.Save
        Call selfEval.Save(False,False)

        ' Send Email
        Call  sendMail(address,findSelfEvalUnid,"Self Evaluation Modification for - "+FirstName+" "+Lastname +"  * Attention Needed","Please re-evaluate the Goals and Competencies as they have changed.")
    End If

End Sub