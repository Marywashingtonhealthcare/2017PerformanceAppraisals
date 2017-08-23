Sub Click(Source As Button)
    Dim s As New NotesSession
    Dim ws As New NotesUIWorkspace
    Dim db As NotesDatabase
    Set db = s.CurrentDatabase
    Dim uidoc As NotesUIDocument
    Dim evaldoc As NotesDocument
    Dim setauthor As NotesItem
    Dim publicacc As NotesItem
    Dim reader As NotesItem
    Dim unid As NotesItem
    Dim saveoptions As NotesItem
    Dim ret As Notesitem
    Dim mbresult As Integer
    Dim mbend As Integer

    Set uidoc = ws.CurrentDocument
    Set doc = uidoc.Document
    uidoc.EditMode = True
    
    mbresult = Msgbox ("Important: Did you set up access in the Setup tab and Save as Draft yet?", 4 + 32 + 0 + 0, "Access + Save Check")
    
    If mbresult = 6 And uidoc.IsNewDoc = False Then
        'do nothing
    Else
        Msgbox ("If you do not set up access and Save as Draft now, users you add later will not be able to access the Self Eval. Please go back and set up access first, then Save as Draft.")
        Exit Sub
    End If
    
    If uidoc.Document.MgrSecFlag(0) = "" And uidoc.Document.MgrNameSec(0) = "" Then
        Messagebox "Please fill out the Manager Conducting Appraisal section before continuing...", 16, "You Forgot Something!"
        Exit Sub
    Else
        Print "..."
    End If

    If doc.selfEvalUNID(0) <> "" Then
        Messagebox "Self-Eval Form has already been created for this Appraisal!!!", 16, "Warning!"
        Exit Sub
    End If

    If doc.depComp1(0) <> "" Then
        'do nothing
    Elseif doc.depComp2(0) <> "" Then
        'do nothing
    Elseif doc.depComp3(0) <> "" Then
        'do nothing
    Elseif doc.depComp4(0) <> "" Then
        'do nothing
    Elseif doc.depComp5(0) <> "" Then
        'do nothing
    Else
        Messagebox "You must fill out at least 1 Department Competency before sending the Self Eval. Please go back and add a competency."
        Exit Sub
    End If

    If doc.LM1(0) <> "" Then
        'do nothing
    Elseif doc.LM6(0) <> "" Then
        'do nothing
    Elseif doc.LM11(0) <> "" Then
        'do nothing
    Elseif doc.LM16(0) <> "" Then
        'do nothing
    Elseif doc.LM21(0) <> "" Then
        'do nothing
    Elseif doc.LM26(0) <> "" Then
        'do nothing
    Elseif doc.LM31(0) <> "" Then
        'do nothing
    Elseif doc.LM36(0) <> "" Then
        'do nothing
    Elseif doc.LM41(0) <> "" Then
        'do nothing
    Elseif doc.LM46(0) <> "" Then
        'do nothing
    Elseif doc.LM51(0) <> "" Then
        'do nothing
    Elseif doc.LM56(0) <> "" Then
        'do nothing
    Elseif doc.LM61(0) <> "" Then
        'do nothing
    Elseif doc.LM66(0) <> "" Then
        'do nothing
    Elseif doc.LM71(0) <> "" Then
        'do nothing
    Else
        Messagebox "You need to fill out at least one goal to continue. Please go back and fill one out before sending this Self Eval"
        Exit Sub
    End If

    'validate email address to make sure it's valid.
    Dim address As String
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


    '*********************************************************(Start Readers Field Creation Code)**************
    'the following lines basically turn the 4 author fields into arrays and then join all 3 arrays into
    'readerArrayFinal, which is used to create the Readers field for the self-eval.
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

    Set evaldoc = db.CreateDocument
    evaldoc.form = "selfEvalAssoc16"
    evaldoc.flagEval = "New"
    evaldoc.FirstName = uidoc.Document.FirstName
    evaldoc.LastName = uidoc.Document.LastName
    evaldoc.AssocNum = uidoc.Document.AssocNum
    evaldoc.depComp1 = uidoc.Document.depComp1
    evaldoc.depComp2 = uidoc.Document.depComp2
    evaldoc.depComp3 = uidoc.Document.depComp3
    evaldoc.depComp4 = uidoc.Document.depComp4
    evaldoc.depComp5 = uidoc.Document.depComp5
    Set ret = New NotesItem(evaldoc,"$$Return", "[http://dapps1.mhs.org/Databases/Human%20Resources/mwhc-pe.nsf/ReturnSubmit?OpenForm]")
    Set setauthor = New NotesItem(evaldoc, "Author", "[Web]", AUTHORS)
    Set publicacc = New Notesitem(evaldoc, "$PublicAccess", "1")
    Set reader = New NotesItem(evaldoc, "Readers1", readerArrayFinal, READERS)
    Set unid = New NotesItem(evaldoc, "UNID", evaldoc.UniversalID)
    Set saveoptions = New NotesItem(evaldoc, "SaveOptions","1")

    'The following code adds the Goals field data to the evaluation *********
    Dim ctrpopgoal, ctractgoal, ctrginfo, ctrpill, ctrthresh, ctrstr, ctract, ctrwt As Integer 'dim all counters
    Dim popgoal, actgoal, ginfo, pill, thresh, stretch, act, wt As String
    Dim factgoal, fpill, fthresh, fstretch, fact, fwt As Variant
    Dim fpopgoal, fginfo As NotesItem
    Dim strpill, strthresh, strstretch, stract, strwt As String
    Dim strginfo As String
    ctrpopgoal = 1 'populate goal text
    ctractgoal = 1 'actual goal text
    ctrginfo = 1 'populate goal info field
    ctrpill = 2 'actual pillar text
    ctrthresh = 3 'actual threshold text
    ctrstr = 4 'actual stretch text
    ctract = 5 'actual actual text
    ctrwt = 1 'actual weight text

    While ctractgoal < 72
        'build the field names as text
        popgoal = "lmGoal" & Cstr(ctrpopgoal)
        actgoal = "LM" & Cstr(ctractgoal)
        ginfo = "gInfo" & Cstr(ctrginfo)
        pill = "LM" & Cstr(ctrpill) 'this is now strategic priority
        thresh = "LM" & Cstr(ctrthresh) 'repurpose this <to good>
        stretch = "LM" & Cstr(ctrstr) 'repurpose this.<to better>
        act = "LM" & Cstr(ctract) 'repurpose this <to best>
        wt = "lmwt" & Cstr(ctrwt)
        'get values of fields on main document (these are arrays that must be converted to strings before they can be used)
        factgoal = doc.GetItemValue(actgoal)
        fpill = doc.GetItemValue(pill)
        fthresh = doc.GetItemValue(thresh)
        fstretch = doc.GetItemValue(stretch)
        fact = doc.GetItemValue(act)
        fwt = doc.GetItemValue(wt)
        Set fpopgoal = evaldoc.ReplaceItemValue(popgoal, factgoal)'set goal text in child doc with goal text of main doc.
        'convert field values to usable strings using join function
        strpill = Join(fpill, Chr(13))
        strthresh = Join(fthresh, Chr(13))
        stract = Join(fact, Chr(13))
        strwt = Join(fwt, Chr(13))
        'concatenate the strings together into a variable and set the field value on the child doc.
        strginfo = "Strategic Priority: " & strpill & Chr(13) & Chr(13) & "Good: " & strthresh & Chr(13) & Chr(13) & "Actual: " & stract & Chr(13) & Chr(13)  & "Weight: " & strwt 
        Set fginfo = evaldoc.ReplaceItemValue(ginfo, strginfo)
        'now increment each counter
        ctrpopgoal = ctrpopgoal + 1
        ctractgoal = ctractgoal + 5
        ctrginfo = ctrginfo + 1
        ctrpill = ctrpill + 5
        ctrthresh = ctrthresh + 5
        ctrstr = ctrstr + 5
        ctract = ctract + 5
        ctrwt = ctrwt + 1
    Wend

    '**********************************************************(End LTM field additions)****
    Call evaldoc.MakeResponse(uidoc.Document)

    Dim email As Variant
    email = doc.fldEmail
    evaldoc.fldEmail = email

    Call evaldoc.Save(False,False)

    doc.selfEvalUNID = evaldoc.universalID
    Call doc.ReplaceItemValue("fldSave", "Yes")
    Call uidoc.Save

    '_______

    evaldoc.docUrl = "http://inotes.medicorp.org/databases/human%20resources/mwhc-pe.nsf/0/" & evaldoc.universalID & "?EditDocument"
    evaldoc.docUrlint = "http://dapps1.mhs.org/databases/human%20resources/mwhc-pe.nsf/0/" & evaldoc.universalID & "?EditDocument"


    Dim url As Variant


    url = evaldoc.docUrl
    urlint = evaldoc.docUrlint

    Dim mdoc As NotesDocument 
    Dim body As NotesMIMEEntity 
    Dim header As NotesMIMEHeader 
    Dim stream As NotesStream 
    Set db = s.CurrentDatabase 
    Set stream = s.CreateStream 
    s.ConvertMIME = False ' Do not convert MIME to rich text 
    Dim subtext As String
    subtext = evaldoc.FirstName(0) + " " + evaldoc.LastName(0)
    Set mdoc = db.CreateDocument 
    mdoc.Form = "Memo" 
    mdoc.Subject = "Self Evaluation for " + subtext
    mdoc.SendTo = email(0)
    Set body = mdoc.CreateMIMEEntity 
    Call stream.writetext(|<html>|) 	
    Call stream.writetext(|<body style="font-family: Arial, Helvetica, sans-serif;"><strong> Please fill out the following Self-Evaluation for your Performance Appraisal by clicking on one of the links below:</strong>|) 
    Call stream.WriteText(|<h4><a href=|)
    Call stream.writetext(|"| & urlint(0) &|"|)
    Call stream.WriteText(| style="margin-right:5px;">Work</a> or|)
    Call stream.WriteText(|<a href=|)
    Call stream.writetext(|"| & url(0) &|"|)
    Call stream.WriteText(| style="margin-left:7px;">Home</a></h4><br>|)
    Call stream.writetext(|<h5>PLEASE DO NOT DELETE THIS EMAIL: </h5>If you need to access your Self Eval again, you will need the link in this email.</body>|) 
    Call stream.writetext(|</html>|) 
    Call body.SetContentFromText(stream, "text/HTML;charset=UTF-8", ENC_IDENTITY_7BIT) 
    Call mdoc.Send(False) 
    s.ConvertMIME = True ' Restore conversion - very important 


    mbend = Msgbox ("Thank you. Self Eval has been sent to "  + subtext +  " using the email address you provided." + Chr(13) + Chr(13) + "CLOSE THIS APPRAISAL?", 4+64+256+0, "Self Eval Send Successful!")

    If mbend = 7 Then
        'do nothing
    Else
        Call uidoc.Close
    End If

End Sub