Option Compare Database
Option Explicit
Public Const OIS_Title = "Avdeling LOG og ØS"
Public GL_URL As String         'url string
Public GL_GID As Long           'gruppe id
Public GL_SID As Long           'studie id
Public GL_EID As Long           'emne id
Public GL_FID As Long           'faglærer id
Public GL_YEAR As String        'studieår
Public GL_FACULTY As String     'avdelingsnavn
Public GL_SELECTION As String   'utvalgsvariabel
Public GL_URLSHB As String      'URL til studiehåndbok
Public GL_URLKPL As String      'URL til kursplan
Public GL_MODE As String        'Modus - editering eller ny
Public GL_ECODE As String       'Emnegruppekode i hovedbilde

Public Const clrBlack As Long = 0
Public Const clrWhite As Long = 16777215
Public Const clrRed As Long = 255
Public Const clrGreen As Long = 4259584
Public Const clrYellow As Long = 65535
Public Const clrBlue As Long = 8388608
Public Const clrGrey As Long = 14540253
' string constants
Public Const strInsert = "Legg til"
Public Const strUpdate = "Oppdater"



Public Function CryptPwd(Pwd As String) As String

    Dim PwdChar         As String * 1
    Dim CryptChar       As String * 1
    Dim i               As Integer
    Dim AsciiValue      As Integer
    
    
    CryptPwd = ""
    If Len(Pwd) = 0 Then
        CryptPwd = ""
        Exit Function
    End If
    For i = 1 To Len(Pwd)
        PwdChar = Mid(Pwd, i, 1)
        AsciiValue = Asc(PwdChar) + i
        If AsciiValue > 126 Then
            AsciiValue = AsciiValue - 94
        End If
        CryptChar = Chr(AsciiValue)
        CryptPwd = CryptPwd + CryptChar
    Next
End Function


Public Function DecryptPwd(CPwd As String) As String
    Dim PwdChar         As String * 1
    Dim CryptChar       As String * 1
    Dim i               As Integer
    Dim AsciiValue      As Integer
    
    
    DecryptPwd = ""
    If Len(CPwd) = 0 Then
        DecryptPwd = ""
        Exit Function
    End If

    For i = 1 To Len(CPwd)
        CryptChar = Mid(CPwd, i, 1)
        AsciiValue = Asc(CryptChar) - i
        If AsciiValue < 33 Then
            AsciiValue = AsciiValue + 94
        End If
        PwdChar = Chr(AsciiValue)
        DecryptPwd = DecryptPwd + PwdChar
    Next
End Function
Public Sub LockTxtboxes(myForm As Form)
On Error GoTo Err_LockTxtboxes
    Dim ctrl As Control
    For Each ctrl In myForm.Controls
        If Left(ctrl.Name, 3) = "txt" Then
            ctrl.Locked = True
            ctrl.BackColor = clrGrey
        End If
    Next
Exit_LockTxtboxes:
    Exit Sub

Err_LockTxtboxes:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_LockTxtboxes

End Sub
Public Function ExitUPlan()
On Error GoTo Err_ExitUPlan
    
    'DoCmd.Close
    DoCmd.Quit
    
Exit_ExitUPlan:
    Exit Function

Err_ExitUPlan:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_ExitUPlan
End Function

Public Function SelectReport(RepNo As Integer)
On Error GoTo Err_SelectReport
    
    Dim stDocName As String
    Dim mnu As CommandBar
    Dim mnuctl As CommandBarControl
    Set mnu = CommandBars("UPlan-Toolbar")
    Set mnuctl = mnu.Controls("Rapporter")
    
    Select Case RepNo
        Case 1
            mnuctl.Tag = "Emneoversikt"
            stDocName = "frmRapport"
        Case 2
            mnuctl.Tag = "Undervisning pr emne"
            stDocName = "frmRapport"
        Case 3
            mnuctl.Tag = "Undervisning pr lærer"
            stDocName = "frmRapport2"
        Case 4
            mnuctl.Tag = "Udekket undervisning"
            stDocName = "frmRapport"
        Case 5
            mnuctl.Tag = "Faglæreroversikt"
            stDocName = "frmRapportAnsatt"
    End Select
    
    DoCmd.OpenForm stDocName

Exit_SelectReport:
    Exit Function

Err_SelectReport:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_SelectReport

End Function

Public Function BrowseWeb(UrlNo As Integer)
On Error GoTo Err_BrowseWeb
Dim stDocType As String
Dim stDocName As String
Select Case UrlNo
    Case 1
        GL_URL = GL_URLKPL
        stDocType = "Kursplan"
    Case 2:
        GL_URL = GL_URLSHB
        stDocType = "Studiehåndbok"
End Select
    stDocName = "frmWeb"
    DoCmd.OpenForm stDocName, , , , , , stDocType

Exit_BrowseWeb:
    Exit Function

Err_BrowseWeb:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_BrowseWeb

End Function
Public Function SnuddNavn(strNavn As String) As String
    Dim Fnavn As String, Enavn As String
    Dim lenN As Integer, lenF As Integer, lenE As Integer
    lenN = Len(strNavn)     'lenght of name
    lenE = InStr(1, strNavn, ",") - 1
    If lenE = 0 Then ' comma separator not found
        SnuddNavn = Trim(strNavn)
        Exit Function
    End If
    lenF = lenN - lenE - 1
    Enavn = Mid(strNavn, 1, lenE)
    Fnavn = Mid(strNavn, lenE + 2)
    SnuddNavn = Trim(Fnavn & " " & Enavn)
End Function
Public Sub showStudyProgram(SID As Long)
On Error GoTo Err_btnLink_Click
    Dim myDb    As DAO.Database
    Dim MyRs1   As DAO.Recordset
    Dim MyRs2   As DAO.Recordset
    Dim msg1 As String, msg2 As String, strSpraak As String
    Dim stDocName As String, stDocType As String
    Dim stTopPage As String
    Dim sqlstr1 As String, sqlstr2 As String
    sqlstr2 = "SELECT StudieID, StudieURL FROM tblStudium " & _
            "WHERE StudieID = " & SID & ";"
    Set myDb = CurrentDb
    
    Set MyRs1 = myDb.OpenRecordset("tblParameter", dbOpenDynaset)
    Set MyRs2 = myDb.OpenRecordset(sqlstr2, dbOpenDynaset)
    strSpraak = FinnSpraak(MyRs2!StudieID, "Number")
    Select Case strSpraak
        Case Is = "N"
            If Not IsNull(MyRs2!StudieURL) And MyRs2!StudieURL <> "" Then
                GL_URL = MyRs1!Studiehandbok & "/" & MyRs2!StudieURL
            Else
                GL_URL = MyRs1!Studiehandbok & "/" & MyRs1!topURL
            End If
        Case Is = "E"
            If Not IsNull(MyRs2!StudieURL) And MyRs2!StudieURL <> "" Then
                GL_URL = MyRs1!Studiehandbok_eng & "/" & MyRs2!StudieURL
            Else
                GL_URL = MyRs1!Studiehandbok_eng & "/" & MyRs1!topURL
            End If
    End Select
    MyRs1.Close
    MyRs2.Close
    myDb.Close
    stDocType = "Studiehåndbok"
    stDocName = "frmWeb"
    DoCmd.OpenForm stDocName, , , , , , stDocType
Exit_btnLink_Click:
    Exit Sub

Err_btnLink_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnLink_Click

End Sub

Public Function FinnSpraak(Ekode As Variant, TypeKode As String) As String
    If TypeKode = "String" Then
        Select Case Left(Ekode, 3)
            Case Is = "LOG"
                If Mid(Ekode, 4, 1) >= "7" Then
                    FinnSpraak = "E"
                Else
                    FinnSpraak = "N"
                End If
            Case Is = "SCM"
                If Mid(Ekode, 4, 1) >= "7" Then
                    FinnSpraak = "E"
                Else
                    FinnSpraak = "N"
                End If
            Case Is = "TRA"
                If Mid(Ekode, 4, 1) >= "8" Then
                    FinnSpraak = "E"
                Else
                    FinnSpraak = "N"
                End If
            Case Is = "PHD"
                FinnSpraak = "E"
            Case Else
                FinnSpraak = "N"
        End Select
    ElseIf TypeKode = "Number" Then
        If (Ekode = 17) Or (Ekode = 27) Or (Ekode = 28) Then
            FinnSpraak = "E"
        Else
            FinnSpraak = "N"
        End If
    End If
End Function
Public Function ValgfriReport()
On Error GoTo Err_ValgfriReport
    
    Dim myDb As DAO.Database
    Dim rsEmne As DAO.Recordset
    Dim rsOblig As DAO.Recordset
    Dim rsReport As DAO.Recordset
    Dim rsParameter As DAO.Recordset
    Dim numStud As Long
    Dim strStudyYear As String
    Dim strEmne As String, strNavn As String
    Dim sqlOblig As String
    Dim sqlStudium As String
    Dim blnFound As Boolean
    Dim stDocName As String

    Set myDb = CurrentDb
    Set rsReport = myDb.OpenRecordset("tmpEmne", dbOpenDynaset)
    
    Do While Not rsReport.EOF
        rsReport.Delete
        rsReport.MoveNext
    Loop
    Set rsParameter = myDb.OpenRecordset("tblParameter", dbOpenDynaset)
    strStudyYear = rsParameter!studyYear
    rsParameter.Close
    Set rsEmne = myDb.OpenRecordset("qryEmne", dbOpenDynaset)
    Set rsOblig = myDb.OpenRecordset("qryOblig", dbOpenDynaset)
    
    Do While Not rsEmne.EOF
        rsOblig.MoveFirst
        blnFound = False
        Do While Not rsOblig.EOF
            If rsEmne!Emnekode = rsOblig!Emnekode Then 'obligatorisk emne funnet
               blnFound = True
               Exit Do
            Else
                rsOblig.MoveNext
            End If
        Loop
        If Not blnFound Then ' valgfritt emne
            rsReport.AddNew
            rsReport!Emnekode = rsEmne!Emnekode
            rsReport!Emnenavn = rsEmne!Emnenavn
            rsReport!Semester = rsEmne!Semester
            rsReport!studyYear = strStudyYear
            rsReport.Update
        End If
        rsEmne.MoveNext
    Loop
    rsOblig.Close
    rsEmne.Close
    rsReport.Close
    stDocName = "rptValgfri"
    GL_SELECTION = "Emner som ikke er obligatoriske på noe studium"
    DoCmd.OpenReport stDocName, acViewPreview 'WhereCondition:=stCriteria
    
Exit_ValgfriReport:
    Exit Function

Err_ValgfriReport:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_ValgfriReport

End Function

Public Sub CreateMenu(strMenu As String)
On Error GoTo Err_CreateMenu
    
    Dim stDocName As String
    Dim mnu As CommandBar
    Dim mnuctl As CommandBarControl
    
    'adding commandbar
    Set mnu = CommandBars.Add(strMenu, , msoBarTypeMenuBar)
    
    'adding menu-control
    Set mnuctl = mnu.Controls.Add(msoControlButton)
        mnuctl.Caption = "Oppsett"
        mnuctl.OnAction = MsgBox("oppsett")
   
    

Exit_CreateMenu:
    Exit Sub

Err_CreateMenu:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_CreateMenu

End Sub

Public Sub EnableMenu()
On Error GoTo Err_EnableMenu
    
    Dim stDocName As String, Num As Integer
    Dim mnu As CommandBar
    Dim mnuctl As CommandBarControl
    Dim mnuctl2 As CommandBarControl
    Dim blnEnabled As Boolean
    Dim blnVisible As Boolean
    Dim msg1 As String
    
    msg1 = "Nytt valg"
    Set mnu = CommandBars("UPlan-Toolbar")
    'Set mnuctl = CommandBars("UPlan-Toolbar").Controls("Vis").Controls(1)
    'mnuctl.Caption = "Emneoversikt"
    'Set mnuctl = CommandBars("Oppsett").Controls.Add(Type:=msoControlButton)
    'With mnuctl
    '    .BeginGroup = True
    '    .Caption = "Nytt valg"
    '    .OnAction = "=MsgBox(msg1)"
    'End With

    mnu.Enabled = True
    mnu.Visible = True
    
Exit_EnableMenu:
    Exit Sub

Err_EnableMenu:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_EnableMenu

End Sub
Public Sub SelectPrinter()
    'DoCmd.OpenForm "frmPrintSetup"
    DoCmd.RunCommand (acCmdPrintSelection)
    Printers.UserControl
End Sub
Public Sub SelectLocations()
    DoCmd.OpenForm "frmLocations"
End Sub
Public Sub EditStilling()
    DoCmd.OpenForm "frmStilling"
End Sub
Public Sub EditEmnekode()
    DoCmd.OpenForm "frmEmnekode"
End Sub
Public Sub EditFaggruppe()
    DoCmd.OpenForm "frmGrupper"
End Sub
Public Sub EditForeleser()
    DoCmd.OpenForm "frmForeleser"
End Sub
Public Sub EditKursogLarer()
    DoCmd.OpenForm "frmForeleserEmne"
End Sub
Public Sub EditEmne()
    DoCmd.OpenForm "frmMain"
End Sub
Public Sub EditObligemner()
    DoCmd.OpenForm "frmObligemner"
End Sub
Public Sub ShowEduPlanner()
    DoCmd.OpenForm "frmAboutEdu"
End Sub
Public Sub ModifyMenu()
On Error GoTo Err_ModifyMenu
    
    Dim stDocName As String, Num As Integer
    Dim mnu As CommandBar
    Dim mnuctl As CommandBarControl
    Dim mnuctl2 As CommandBarControl
    
    Set mnu = CommandBars("UPlan-Toolbar")
Exit_ModifyMenu:
    Exit Sub

Err_ModifyMenu:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_ModifyMenu

End Sub