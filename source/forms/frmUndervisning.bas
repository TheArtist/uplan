Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
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
    Width =6637
    DatasheetFontHeight =10
    ItemSuffix =24
    Left =2556
    Right =9528
    Bottom =6132
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb210ce4d139ae240
    End
    Caption ="Mainenance Report"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
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
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
        End
        Begin Section
            Height =5517
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =3921
                    Top =4700
                    Width =1140
                    FontWeight =300
                    ForeColor =0
                    Name ="btnPreview"
                    Caption ="Preview"
                    OnClick ="[Event Procedure]"
                    FontName ="Verdana"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =401
                    Top =4700
                    Width =1128
                    TabIndex =1
                    Name ="btnCancel"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =401
                    Top =705
                    Width =5929
                    Height =3827
                    TabIndex =2
                    Name ="frSelect"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =516
                            Top =588
                            Width =2496
                            Height =228
                            BackColor =-2147483633
                            Name ="Label8"
                            Caption ="Select system(s) and report criteria"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =614
                            Top =4204
                            Width =186
                            Height =213
                            OptionValue =1
                            Name ="optAll"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =827
                                    Top =4157
                                    Width =408
                                    Height =228
                                    Name ="Label19"
                                    Caption ="All"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1795
                            Top =4204
                            Width =210
                            Height =213
                            OptionValue =2
                            Name ="optNotAnalysed"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =2008
                                    Top =4157
                                    Width =1056
                                    Height =228
                                    Name ="Label12"
                                    Caption ="Not analysed"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =3354
                            Top =4204
                            Width =174
                            Height =213
                            OptionValue =3
                            Name ="optPartAnalysed"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =3568
                                    Top =4157
                                    Width =1227
                                    Height =228
                                    Name ="Label14"
                                    Caption ="Partly analysed"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =5007
                            Top =4204
                            Width =186
                            Height =213
                            OptionValue =4
                            Name ="optCompleted"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =5220
                                    Top =4157
                                    Width =936
                                    Height =228
                                    Name ="Label17"
                                    Caption ="Completed"
                                End
                            End
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =567
                    Top =985
                    Width =5551
                    Height =2904
                    TabIndex =3
                    Name ="lstSystem"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;567;1442"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5173
                    Top =4700
                    Width =1140
                    FontWeight =300
                    TabIndex =4
                    ForeColor =0
                    Name ="btnPrint"
                    Caption ="Print"
                    OnClick ="[Event Procedure]"
                    FontName ="Verdana"
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
    
Public Sub GenerateReport(strView As String)
On Error GoTo Err_GenerateReport
    Dim stDocName As String
    Dim stCriteria As String
    Dim mnu As CommandBar
    Dim mnuctl As CommandBarControl
    Dim sqlAnalysed As String
    Dim myList As ListBox
    Dim i As Integer, j As Integer
    Dim SysId() As Long
    Set myList = Me!lstSystem
    Set mnu = CommandBars("RCM-Toolbar")
    Set mnuctl = mnu.Controls("Reports")
    Select Case mnuctl.Tag
        Case "SystemTag"
            stDocName = "rptSystemTag"
        Case "FailureAnalysis"
            stDocName = "rptFailureAnalysis"
        Case "FailureEffect"
            stDocName = "rptFailureEffect"
        Case "FailureCriticality"
            stDocName = "rptFailureCriticality"
        Case "MaintenanceActivity"
            stDocName = "rptMaintenance"
        Case "FMECA"
            stDocName = "rptFMECAnalysis"
    End Select
    Select Case Me.frSelect     ' selection criteria
        Case 1                  'all
            sqlAnalyse = ""
        Case 2                  ' not analysed
            sqlAnalyse = "AnalysedFM = False AND AnalysedCE = False AND AnalysedMRD = False"
        Case 3                  ' partly analysed
            sqlAnalyse = "(AnalysedFM = True OR AnalysedCE = True OR AnalysedMRD = True)"
        Case 4                  ' completed
            sqlAnalyse = "AnalysedFM = True AND AnalysedCE = True AND AnalysedMRD = True"
    End Select
    'Find selected systems
    j = 0
    ReDim SysId(myList.ListCount)
    For i = 0 To myList.ListCount - 1
        If myList.Selected(i) Then
            j = j + 1
            SysId(j) = myList.ItemData(i) ' = SystemID
        End If
    Next i
    If j > 0 Then       ' at least one system is selected
        For i = 1 To j
            If sqlAnalyse <> "" Then
                stCriteria = "SystemID = " & SysId(i) & " AND " & sqlAnalyse
            Else
                stCriteria = "SystemID = " & SysId(i)
            End If
            Select Case strView
                Case "Print"
                    DoCmd.OpenReport stDocName, acViewNormal, WhereCondition:=stCriteria
                Case "Preview"
                    DoCmd.OpenReport stDocName, acViewPreview, WhereCondition:=stCriteria
            End Select
        Next i
    Else        'no systems selected, print them all in one report
        If sqlAnalyse <> "" Then
            Select Case strView
                Case "Print"
                    DoCmd.OpenReport stDocName, acViewNormal, WhereCondition:=sqlAnalyse
                Case "Preview"
                    DoCmd.OpenReport stDocName, acViewPreview, WhereCondition:=sqlAnalyse
            End Select
        Else
            Select Case strView
                Case "Print"
                    DoCmd.OpenReport stDocName, acViewNormal
                Case "Preview"
                    DoCmd.OpenReport stDocName, acViewPreview
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


Private Sub Form_Load()
On Error GoTo Err_Form_Load

    Dim sqlSource As String
    Dim mnu As CommandBar
    Dim mnuctl As CommandBarControl
        
    sqlSource = "SELECT SystemID, SID, Name FROM System ORDER BY SystemID;"
    Me.lstSystem.RowSource = sqlSource
    Set mnu = CommandBars("RCM-Toolbar")
    Set mnuctl = mnu.Controls("Reports")
    Select Case mnuctl.Tag
        Case "SystemTag"
            Me.Caption = "System and Tag Report"
        Case "FailureAnalysis"
            Me.Caption = "Failure Analysis Report"
        Case "FailureEffect"
            Me.Caption = "Failure Mode and Effect Report"
        Case "FailureEffectCriticality"
            Me.Caption = "Failure Mode and Criticality Report"
        Case "MaintenanceActivity"
            Me.Caption = "Maintenance Activity Report"
        Case "FMECA"
            Me.Caption = "FMECA Report"
    End Select
    Me.frSelect.Value = 1   'default is all equipment
Exit_Form_Load:
    Exit Sub

Err_Form_Load:
    MsgBox Err.Description
    Resume Exit_Form_Load
    

End Sub


Private Sub btnPrint_Click()
On Error GoTo Err_btnPrint_Click


   Call GenerateReport("Print")
   
Exit_btnPrint_Click:
    Exit Sub

Err_btnPrint_Click:
    MsgBox Err.Description
    Resume Exit_btnPrint_Click
    
End Sub
Private Sub btnPreview_Click()
On Error GoTo Err_btnPreview_Click


   Call GenerateReport("Preview")

Exit_btnPreview_Click:
    Exit Sub

Err_btnPreview_Click:
    MsgBox Err.Description
    Resume Exit_btnPreview_Click
    
End Sub
