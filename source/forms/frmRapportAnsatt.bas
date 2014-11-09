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
    Width =6519
    DatasheetFontHeight =10
    ItemSuffix =31
    Left =855
    Top =2205
    Right =7380
    Bottom =5520
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5f27304fdd97e340
    End
    RecordSource ="qryAnsatt"
    Caption ="Faglæreroversikt"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Section
            Height =3330
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =3118
                    Top =2574
                    Width =1140
                    FontSize =9
                    FontWeight =300
                    ForeColor =0
                    Name ="btnPreview"
                    Caption ="Vis"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3118
                    LayoutCachedTop =2574
                    LayoutCachedWidth =4258
                    LayoutCachedHeight =2934
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =425
                    Top =2598
                    Width =1128
                    FontSize =9
                    TabIndex =1
                    Name ="btnCancel"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =425
                    LayoutCachedTop =2598
                    LayoutCachedWidth =1553
                    LayoutCachedHeight =2958
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =401
                    Top =705
                    Width =5362
                    Height =1625
                    TabIndex =2
                    Name ="frSelect"

                    LayoutCachedLeft =401
                    LayoutCachedTop =705
                    LayoutCachedWidth =5763
                    LayoutCachedHeight =2330
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =510
                            Top =585
                            Width =2925
                            Height =225
                            FontSize =9
                            BackColor =-2147483633
                            Name ="Label8"
                            Caption =" Velg stillingstype eller stillingsandel"
                            LayoutCachedLeft =510
                            LayoutCachedTop =585
                            LayoutCachedWidth =3435
                            LayoutCachedHeight =810
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4370
                    Top =2574
                    Width =1140
                    FontSize =9
                    FontWeight =300
                    TabIndex =3
                    ForeColor =0
                    Name ="btnPrint"
                    Caption ="Skriv ut"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4370
                    LayoutCachedTop =2574
                    LayoutCachedWidth =5510
                    LayoutCachedHeight =2934
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2768
                    Top =1677
                    Width =2728
                    Height =297
                    FontSize =9
                    TabIndex =4
                    Name ="cboAndel"
                    RowSourceType ="Value List"
                    RowSource ="Alle;100;>=60;60;>20;20;<=20;10;0"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =2768
                    LayoutCachedTop =1677
                    LayoutCachedWidth =5496
                    LayoutCachedHeight =1974
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =685
                            Top =1717
                            Width =1941
                            Height =240
                            FontSize =9
                            Name ="Label27"
                            Caption ="Stillingsandel (%):"
                            LayoutCachedLeft =685
                            LayoutCachedTop =1717
                            LayoutCachedWidth =2626
                            LayoutCachedHeight =1957
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2768
                    Top =1110
                    Width =2743
                    Height =297
                    FontSize =9
                    TabIndex =5
                    Name ="cboStilling"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =2768
                    LayoutCachedTop =1110
                    LayoutCachedWidth =5511
                    LayoutCachedHeight =1407
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =685
                            Top =1150
                            Width =1941
                            Height =240
                            FontSize =9
                            Name ="Label30"
                            Caption ="Stillingstype:"
                            LayoutCachedLeft =685
                            LayoutCachedTop =1150
                            LayoutCachedWidth =2626
                            LayoutCachedHeight =1390
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
    
Public Sub TeacherReport(stC As String, stDoc As String)
On Error GoTo Err_TeacherReport
    
    'Dim mnu As CommandBar
    'Dim mnuctl As CommandBarControl
    'Set mnu = CommandBars("UPlan-Toolbar")
    'Set mnuctl = mnu.Controls("Rapporter")

    Dim stCriteria As String
    Dim sqlAnalysed As String
    'Stillingsandel
    If Not IsNull(Me.cboAndel) And Me.cboAndel <> "" Then
        Select Case Me.cboAndel
            Case "Alle"
                GL_SELECTION = "Utvalg: Alle stillinger"
                stCriteria = "Andel >= 0"
            Case "100"
                GL_SELECTION = "Utvalg: Stillingsandel = " & Me.cboAndel & "%"
                stCriteria = "Andel = 100"
            Case ">=60"
                GL_SELECTION = "Utvalg: Stillingsandel " & Me.cboAndel & "%"
                stCriteria = "Andel >= 60"
            Case "60"
                GL_SELECTION = "Utvalg: Stillingsandel = " & Me.cboAndel & "%"
                stCriteria = "Andel = 60"
            Case ">20"
                GL_SELECTION = "Utvalg: Stillingsandel " & Me.cboAndel & "%"
                stCriteria = "Andel > 20"
            Case "20"
                GL_SELECTION = "Utvalg: Stillingsandel = " & Me.cboAndel & "%"
                stCriteria = "Andel = 20"
            Case "<=20"
                GL_SELECTION = "Utvalg: Stillingsandel " & Me.cboAndel & "%"
                stCriteria = "Andel <= 20"
            Case "10"
                GL_SELECTION = "Utvalg: Stillingsandel = " & Me.cboAndel & "%"
                stCriteria = "Andel = 10"
            Case "0"
                GL_SELECTION = "Utvalg: Stillingsandel = " & Me.cboAndel & "%"
                stCriteria = "Andel = 0"
        End Select
    ElseIf Not IsNull(Me.cboStilling) And Me.cboStilling <> "" Then
        GL_SELECTION = "Utvalg: Stilling = " & Me.cboStilling
        stCriteria = "StNavn = '" & Me.cboStilling & "'"
    End If
    stC = stCriteria
    stDoc = "rptAnsatt"
Exit_TeacherReport:
    Exit Sub

Err_TeacherReport:
    MsgBox Err.Description
    Resume Exit_TeacherReport
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

Private Sub cboAndel_AfterUpdate()
    Me.cboStilling = ""
End Sub

Private Sub cboStilling_AfterUpdate()
    Me.cboAndel = ""
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load
    
    Dim sqlSource As String
    sqlSource = "SELECT StNavn FROM tblStilling ORDER BY StNavn;"
    Me.cboStilling.RowSource = sqlSource
    'Me.cboStilling.RowSource.valuelist = True
    'Me.cboStilling.AddItem ("Alle")

Exit_Form_Load:
    Exit Sub

Err_Form_Load:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_Form_Load
    

End Sub


Private Sub btnPrint_Click()
On Error GoTo Err_btnPrint_Click
    Dim Criteria As String, Docname As String
    Call TeacherReport(Criteria, Docname)
    DoCmd.OpenReport Docname, acViewNormal, WhereCondition:=Criteria
  
Exit_btnPrint_Click:
    Exit Sub

Err_btnPrint_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnPrint_Click
    
End Sub
Private Sub btnPreview_Click()
On Error GoTo Err_btnPreview_Click
    Dim Criteria As String, Docname As String
    Call TeacherReport(Criteria, Docname)
    DoCmd.OpenReport Docname, acViewPreview, WhereCondition:=Criteria

Exit_btnPreview_Click:
    Exit Sub

Err_btnPreview_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnPreview_Click
    
End Sub



'Private Sub btnPrnt_Click()
'On Error GoTo Err_btnPrnt_Click'

'    Dim stDocName As String
'
'    stDocName = "rptAnsatt"
'    DoCmd.OpenReport stDocName, acNormal'
'
'Exit_btnPrnt_Click:
'    Exit Sub
'
'Err_btnPrnt_Click:
'    MsgBox Err.Description
'    Resume Exit_btnPrnt_Click
'
'End Sub
