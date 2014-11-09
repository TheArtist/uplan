Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6543
    DatasheetFontHeight =10
    ItemSuffix =31
    Left =4800
    Top =1050
    Right =11340
    Bottom =6945
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7354d1591821e340
    End
    RecordSource ="qryUndervisning"
    Caption ="Undervisningsrapport"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =5905
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =3945
                    Top =5362
                    Width =1140
                    FontSize =9
                    FontWeight =300
                    ForeColor =0
                    Name ="btnPreview"
                    Caption ="Vis"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3945
                    LayoutCachedTop =5362
                    LayoutCachedWidth =5085
                    LayoutCachedHeight =5722
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =425
                    Top =5362
                    Width =1128
                    FontSize =9
                    TabIndex =1
                    Name ="btnCancel"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =425
                    LayoutCachedTop =5362
                    LayoutCachedWidth =1553
                    LayoutCachedHeight =5722
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =401
                    Top =705
                    Width =5929
                    Height =4562
                    TabIndex =2
                    Name ="frSelect"
                    BeforeUpdate ="[Event Procedure]"

                    LayoutCachedLeft =401
                    LayoutCachedTop =705
                    LayoutCachedWidth =6330
                    LayoutCachedHeight =5267
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =510
                            Top =585
                            Width =1680
                            Height =225
                            FontSize =9
                            BackColor =-2147483633
                            Name ="Label8"
                            Caption ="Velg emnegrupper"
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =590
                            Top =1015
                            Width =268
                            Height =215
                            OptionValue =1
                            Name ="chkAlle"

                            LayoutCachedLeft =590
                            LayoutCachedTop =1015
                            LayoutCachedWidth =858
                            LayoutCachedHeight =1230
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =850
                                    Top =979
                                    Width =1620
                                    Height =225
                                    FontSize =9
                                    Name ="Label25"
                                    Caption ="Alle emnegrupper"
                                    LayoutCachedLeft =850
                                    LayoutCachedTop =979
                                    LayoutCachedWidth =2470
                                    LayoutCachedHeight =1204
                                End
                            End
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =567
                    Top =1393
                    Width =5551
                    Height =3642
                    FontSize =9
                    TabIndex =3
                    Name ="lstEmner"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1134;2010"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =567
                    LayoutCachedTop =1393
                    LayoutCachedWidth =6118
                    LayoutCachedHeight =5035
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5197
                    Top =5362
                    Width =1140
                    FontSize =9
                    FontWeight =300
                    TabIndex =4
                    ForeColor =0
                    Name ="btnPrint"
                    Caption ="Skriv ut"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5197
                    LayoutCachedTop =5362
                    LayoutCachedWidth =6337
                    LayoutCachedHeight =5722
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =4463
                    Top =967
                    Width =1678
                    Height =237
                    FontSize =9
                    TabIndex =5
                    Name ="cboSem"
                    RowSourceType ="Value List"
                    RowSource ="Høst;Vår;Høst og vår"

                    LayoutCachedLeft =4463
                    LayoutCachedTop =967
                    LayoutCachedWidth =6141
                    LayoutCachedHeight =1204
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3479
                            Top =964
                            Width =945
                            Height =240
                            FontSize =9
                            Name ="Label27"
                            Caption ="Semester:"
                            LayoutCachedLeft =3479
                            LayoutCachedTop =964
                            LayoutCachedWidth =4424
                            LayoutCachedHeight =1204
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2716
                    Top =5362
                    Width =1128
                    FontSize =9
                    TabIndex =6
                    Name ="btnExcel"
                    Caption ="Til excel"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2716
                    LayoutCachedTop =5362
                    LayoutCachedWidth =3844
                    LayoutCachedHeight =5722
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    OverlapFlags =87
                    Left =4464
                    Top =425
                    Width =313
                    Height =260
                    TabIndex =7
                    Name ="chkAktive"

                    LayoutCachedLeft =4464
                    LayoutCachedTop =425
                    LayoutCachedWidth =4777
                    LayoutCachedHeight =685
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4715
                            Top =401
                            Width =1620
                            Height =225
                            FontSize =9
                            Name ="Label30"
                            Caption ="Bare aktive emner"
                            LayoutCachedLeft =4715
                            LayoutCachedTop =401
                            LayoutCachedWidth =6335
                            LayoutCachedHeight =626
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
    
    
Sub Inintiaze()
    

End Sub

Public Sub GenerateReport(strView As String)
On Error GoTo Err_GenerateReport
    
    Dim mnu As CommandBar
    Dim mnuctl As CommandBarControl
    Set mnu = CommandBars("UPlan-Toolbar")
    Set mnuctl = mnu.Controls("Rapporter")

    Dim stDocName As String
    Dim stCriteria As String
    Dim sqlAnalysed As String
    Dim strSem As String
    Dim myList As ListBox
    Dim i As Integer, j As Integer
    Dim blnAktive As Boolean
    Dim Emne() As String * 3
    Const strH = "H"
    Const strV = "V"
    Const strHV = "H+V"
    Const strVH = "V+H"
    
    
    Select Case mnuctl.Tag
        Case "Emneoversikt"
            stDocName = "rptEmne"
        Case "Undervisning pr emne"
            stDocName = "rptUndervisning"
        Case "Undervisning pr lærer"
            stDocName = "rptBelastning"
        Case "Udekket undervisning"
            stDocName = "rptEmneIkkePlanlagt"
            'Call Backlog    'lager backlog-tabell
    End Select
    blnAktive = Me.chkAktive.Value
    
    'Semester
    Select Case Me.cboSem
        Case "Høst"
            strSem = "Semester = '" & strH & "' OR Semester = '" & strHV & "' OR Semester = '" & strVH & "'"
        Case "Vår"
            strSem = "Semester = '" & strV & "' OR Semester = '" & strHV & "' OR Semester = '" & strVH & "'"
        Case "Høst og vår"
            strSem = ""
        Case Else
            strSem = ""
    End Select
    If Me.frSelect.Value = 1 Then ' skriv ut alle emner
        ' bare aktive emner
        If blnAktive = True Then
            GL_SELECTION = "Utvalg = Alle aktive emner"
            If strSem <> "" Then
                stCriteria = "Aktiv = " & blnAktive & " AND (" & strSem & ")"
            Else
                stCriteria = "Aktiv = " & blnAktive
            End If
        Else
            GL_SELECTION = "Utvalg = Alle emner"
            If strSem <> "" Then
                stCriteria = strSem
            Else
                stCriteria = ""
            End If
        End If
            
        Select Case strView
            Case "Print"
                DoCmd.OpenReport stDocName, acViewNormal, WhereCondition:=stCriteria
            Case "Preview"
                DoCmd.OpenReport stDocName, acViewPreview, WhereCondition:=stCriteria
        End Select
    Else
        Set myList = Me!lstEmner
        'Find selected subjects
        j = 0
        ReDim Emne(myList.ListCount)
        For i = 0 To myList.ListCount - 1
            If myList.Selected(i) Then
                j = j + 1
                Emne(j) = myList.ItemData(i) ' = Valgt emne
                'MsgBox "index, og emne er: " & j & Emne(j)
            End If
        Next i
        If j > 0 Then       ' minst et emne er valgt
            stCriteria = "Left(Emnekode,3) = '" & Emne(1) & "'"
            GL_SELECTION = "Utvalg = " & Emne(1)
            For i = 2 To j
                stCriteria = stCriteria & " OR Left(Emnekode,3) = '" & Emne(i) & "'"
                GL_SELECTION = GL_SELECTION & ", " & Emne(i)
            Next i
            If blnAktive = True Then
                GL_SELECTION = GL_SELECTION & " (Aktive)"
                If strSem <> "" Then
                    stCriteria = "Aktiv = " & blnAktive & " AND (" & strSem & ") AND (" & stCriteria & ")"
                Else
                    stCriteria = "Aktiv = " & blnAktive & " AND (" & stCriteria & ")"
                End If
            Else
                If strSem <> "" Then
                    stCriteria = "(" & strSem & ") AND (" & stCriteria & ")"
                End If
            End If
            Select Case strView
                Case "Print"
                    DoCmd.OpenReport stDocName, acViewNormal, WhereCondition:=stCriteria
                Case "Preview"
                    DoCmd.RunCommand (acCmdPrintSelection)
                    DoCmd.OpenReport stDocName, acViewPreview, WhereCondition:=stCriteria
            End Select
        End If
    End If
Exit_GenerateReport:
    Exit Sub

Err_GenerateReport:
    MsgBox Err.Description
    Resume Exit_GenerateReport
End Sub

Private Sub btnCancel_Click()
On Error GoTo Err_btnCancel_Click


    DoCmd.Close

Exit_btnCancel_Click:
    Exit Sub

Err_btnCancel_Click:
    MsgBox Err.Description
    Resume Exit_btnCancel_Click
    
End Sub




Private Sub btnExcel_Click()
On Error GoTo Err_btnExcel_Click

    Me.frSelect = 1
    Me.cboSem = "Høst og vår"
    Call GenerateExcel
   
Exit_btnExcel_Click:
    Exit Sub

Err_btnExcel_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnExcel_Click
    

End Sub


Private Sub Form_Load()
On Error GoTo Err_Form_Load
    Dim mnu As CommandBar
    Dim mnuctl As CommandBarControl
    Set mnu = CommandBars("UPlan-Toolbar")
    Set mnuctl = mnu.Controls("Rapporter")
    Set Application.Printer = Nothing
    If mnuctl.Tag = "Undervisning pr emne" Then
        Me.btnExcel.Visible = True
    Else
        Me.btnExcel.Visible = False
    End If
           
    sqlSource = "SELECT Kode, Beskrivelse FROM tblEmnekode ORDER BY Kode;"
    Me.lstEmner.RowSource = sqlSource
    Me.frSelect.Value = 0
Exit_Form_Load:
    Exit Sub

Err_Form_Load:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_Form_Load
    

End Sub


Private Sub btnPrint_Click()
On Error GoTo Err_btnPrint_Click


   Call GenerateReport("Print")
   
Exit_btnPrint_Click:
    Exit Sub

Err_btnPrint_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnPrint_Click
    
End Sub
Private Sub btnPreview_Click()
On Error GoTo Err_btnPreview_Click


   Call GenerateReport("Preview")

Exit_btnPreview_Click:
    Exit Sub

Err_btnPreview_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnPreview_Click
    
End Sub

Private Sub frSelect_BeforeUpdate(Cancel As Integer)
    Dim sqlSource As String
    If Me.frSelect.Value > 0 Then
        sqlSource = "SELECT Kode, Beskrivelse FROM tblEmnekode ORDER BY Kode;"
        Me.lstEmner.RowSource = sqlSource
    End If

End Sub

Private Sub lstEmner_Click()
    Me.frSelect.Value = 0
End Sub

Public Sub Backlog()
On Error GoTo Err_Backlog
    Dim myDb As DAO.Database
    Dim emneRS As DAO.Recordset
    Dim leRS As DAO.Recordset
    Dim backlogRS As DAO.Recordset
    Dim lngEID As Long
    Dim sqlLarerEmne As String
    Dim sngStp As Single, sngRStp As Single
    Dim NoOfEmne As Integer
    Set myDb = CurrentDb()
    Set backlogRS = myDb.OpenRecordset("tblBacklog", dbOpenDynaset)
    sqlDelete = "DELETE * FROM tblBacklog;"
    myDb.Execute (sqlDelete)
    
    Set emneRS = myDb.OpenRecordset("tblEmne", dbOpenDynaset)
    Do While Not emneRS.EOF
        If emneRS!Aktiv And emneRS!Studiepoeng < 30 Then
            lngEID = emneRS!EmneID
            sqlLarerEmne = "SELECT * FROM tblLarerEmne WHERE EmneID=" & lngEID & ";"
            Set leRS = myDb.OpenRecordset(sqlLarerEmne, dbOpenSnapshot)
            If leRS.RecordCount = 0 Then        ' no maching records found
                sngRStp = emneRS!Studiepoeng
            Else
                sngStp = 0
                Do While Not leRS.EOF
                    sngStp = sngStp + leRS!Studiepoeng
                    leRS.MoveNext
                Loop
                sngRStp = emneRS!Studiepoeng - sngStp
            End If
            If sngRStp > 0 Then '
                backlogRS.AddNew
                    backlogRS!Emnekode = emneRS!Emnekode
                    backlogRS!Emnenavn = emneRS!Emnenavn
                    backlogRS!Sted = emneRS!Sted
                    backlogRS!Aktiv = emneRS!Aktiv
                    backlogRS!Semester = emneRS!Semester
                    If IsNumeric(sngRStp) Then
                        backlogRS!RStp = sngRStp
                    Else
                        backlogRS!RStp = 0
                    End If
                    backlogRS!Comment = emneRS!Comment
                backlogRS.Update
            End If
        End If
        emneRS.MoveNext
    Loop

Exit_Backlog:
    Exit Sub

Err_Backlog:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_Backlog

End Sub

Public Sub GenerateExcel()
On Error GoTo Err_GenerateExcel
    Dim ExApp As Object
    Dim ExBook As Excel.Workbook
    Dim ExSheet As Excel.Worksheet
    Dim ExRange As Range
    Dim MyCol As Range
    Dim myDb As DAO.Database
    Dim emneRS As DAO.Recordset
    Dim leRS As DAO.Recordset
    Dim fagRS As DAO.Recordset
    Dim backlogRS As DAO.Recordset
    Dim lngEID As Long
    Dim sqlLarerEmne As String, sqlGruppe As String
    Dim sngStp As Single, sngRStp As Single
    Dim i As Integer, j As Integer, k As Integer
    Dim strSemester As String, strEmnekode As String
    Dim blnAktiv As Boolean
    Dim forrigeKurs As String, forrigeSted As String
    Dim sheetNo As Integer
    Dim fagNavn(1 To 5) As String
    Dim C1 As String, C2 As String

    Set myDb = CurrentDb()

    sqlGruppe = "SELECT * FROM tblGruppe ORDER BY GruppeID;"
    Set fagRS = myDb.OpenRecordset(sqlGruppe, dbOpenSnapshot)
    Do While Not fagRS.EOF
        If fagRS!GruppeID <= 5 Then
            k = fagRS!GruppeID
            fagNavn(k) = fagRS!GruppeNavn
        End If
        fagRS.MoveNext
    Loop
    fagRS.Close
    
    Set ExBook = ExApp.Workbooks.Add
'Add two more worksheets
    ExBook.Sheets.Add   'sheet no 4
    ExBook.Sheets.Add   'sheet no 5
    sheetNo = 1
    Do While sheetNo <= 5
        Set ExSheet = ExBook.Sheets(sheetNo)
        ExSheet.PageSetup.CenterFooter = fagNavn(sheetNo)
        ExSheet.Name = fagNavn(sheetNo)
        ExSheet.Cells(1, 1).ColumnWidth = 10
        ExSheet.Cells(1, 1).RowHeight = 20
        ExSheet.Cells(1, 2).ColumnWidth = 35
        ExSheet.Cells(1, 3).ColumnWidth = 8
        ExSheet.Cells(1, 4).ColumnWidth = 20
        ExSheet.Cells(1, 5).ColumnWidth = 8
        ExSheet.Cells(1, 6).ColumnWidth = 8
        ExSheet.Cells(1, 7).ColumnWidth = 25
        ExSheet.Cells(1, 8).ColumnWidth = 8
        
    
        Set ExRange = ExSheet.Range("A1", "H1")
        ExRange.Merge
        ExRange.Font.Size = 16
        ExRange.Font.Bold = True
        ExSheet.Cells(1, 1) = "Undervisningsplan pr emne, Avd ØIS " & GL_YEAR
    
        Set ExRange = ExSheet.Range("A2", "H2")
        ExRange.Font.Bold = True
        ExRange.Borders.LineStyle = True
        ExSheet.Cells(2, 1) = "Kode"
        ExSheet.Cells(2, 2) = "Kursnavn"
        ExSheet.Cells(2, 3) = "Stp"
        ExSheet.Cells(2, 4) = "Sted"
        ExSheet.Cells(2, 5) = "Høst"
        ExSheet.Cells(2, 6) = "Vår"
        ExSheet.Cells(2, 7) = "Faglærer"
        ExSheet.Cells(2, 8) = "Ok"
        
        ExSheet.Columns(3).HorizontalAlignment = xlCenter
        ExSheet.Columns(5).HorizontalAlignment = xlCenter
        ExSheet.Columns(6).HorizontalAlignment = xlCenter
        ExSheet.Columns(8).HorizontalAlignment = xlCenter

        sqlLarerEmne = "SELECT * FROM qryExcel WHERE (Aktiv = True AND FagID = " & sheetNo & ") ORDER BY Emnekode, Sted, Semester;"
        Set leRS = myDb.OpenRecordset(sqlLarerEmne, dbOpenSnapshot)
        'Debug.Print leRS.RecordCount
        i = 3
        forrigeKurs = ""
        forrigeSted = ""
        Do
            If (leRS!Emnekode <> forrigeKurs) Then ' nytt kurs, setter inn blank linje
                C1 = "A" & i - 1
                C2 = "H" & i - 1
                Set ExRange = ExSheet.Range(C1, C2)
                ' HFN FIX: ExRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                ExSheet.Cells(i, 1) = leRS!Emnekode
                ExSheet.Cells(i, 2) = leRS!Emnenavn
                ExSheet.Cells(i, 3) = leRS![tblEmne.Studiepoeng]
                ExSheet.Cells(i, 4) = leRS!Sted
            ElseIf (leRS!Emnekode = forrigeKurs) And (leRS!Sted <> forrigeSted) Then
                ExSheet.Cells(i, 1) = ""
                ExSheet.Cells(i, 2) = ""
                ExSheet.Cells(i, 4) = leRS!Sted
                ExSheet.Cells(i, 3) = leRS![tblEmne.Studiepoeng]
            End If
            strSemester = leRS!Semester
            Select Case strSemester
                Case Is = "H"
                    ExSheet.Cells(i, 5) = leRS![tblLarerEmne.Studiepoeng]
                    ExSheet.Cells(i, 6) = ""
                Case Is = "V"
                    ExSheet.Cells(i, 5) = ""
                    ExSheet.Cells(i, 6) = leRS![tblLarerEmne.Studiepoeng]
                Case Is = "H+V"
                    ExSheet.Cells(i, 5) = leRS![tblLarerEmne.Studiepoeng] / 2
                    ExSheet.Cells(i, 6) = leRS![tblLarerEmne.Studiepoeng] / 2
                Case Is = "V+H"
                    ExSheet.Cells(i, 5) = leRS![tblLarerEmne.Studiepoeng] / 2
                    ExSheet.Cells(i, 6) = leRS![tblLarerEmne.Studiepoeng] / 2
            End Select
            ExSheet.Cells(i, 7) = leRS!Navn
            If leRS!Ferdig = True Then
                ExSheet.Cells(i, 8) = "X"
            Else
                ExSheet.Cells(i, 8) = ""
            End If
            If Not IsNull(leRS!Kommentar) And leRS!Kommentar <> "" Then
                ExSheet.Cells(i, 8).AddComment (leRS!Kommentar)
            End If
            forrigeKurs = leRS!Emnekode
            forrigeSted = leRS!Sted
            i = i + 1
            leRS.MoveNext
        Loop While Not leRS.EOF
Next_Worksheet:
        sheetNo = sheetNo + 1
    Loop
Show_Worksheet:
    ExApp.Visible = True
    ExApp.UserControl = True

Exit_GenerateExcel:
    Exit Sub

Err_GenerateExcel:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_GenerateExcel

End Sub
