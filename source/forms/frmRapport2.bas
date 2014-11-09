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
    Width =6779
    DatasheetFontHeight =10
    ItemSuffix =35
    Left =840
    Top =615
    Right =7620
    Bottom =5280
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7c699e037f33e440
    End
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
            Height =4677
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =3874
                    Top =4062
                    Width =1140
                    FontSize =9
                    FontWeight =300
                    ForeColor =0
                    Name ="btnPreview"
                    Caption ="Vis"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3874
                    LayoutCachedTop =4062
                    LayoutCachedWidth =5014
                    LayoutCachedHeight =4422
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =354
                    Top =4062
                    Width =1128
                    FontSize =9
                    TabIndex =1
                    Name ="btnCancel"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =354
                    LayoutCachedTop =4062
                    LayoutCachedWidth =1482
                    LayoutCachedHeight =4422
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =354
                    Top =852
                    Width =5959
                    Height =3017
                    TabIndex =2
                    Name ="frSelect"

                    LayoutCachedLeft =354
                    LayoutCachedTop =852
                    LayoutCachedWidth =6313
                    LayoutCachedHeight =3869
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =685
                            Top =732
                            Width =2010
                            Height =225
                            FontSize =9
                            BackColor =-2147483633
                            Name ="Label8"
                            Caption =" eller velg fagtilhørighet"
                            LayoutCachedLeft =685
                            LayoutCachedTop =732
                            LayoutCachedWidth =2695
                            LayoutCachedHeight =957
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =543
                    Top =1273
                    Width =5611
                    Height =2397
                    FontSize =9
                    TabIndex =3
                    Name ="lstGrupper"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="567;2578"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =543
                    LayoutCachedTop =1273
                    LayoutCachedWidth =6154
                    LayoutCachedHeight =3670
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5126
                    Top =4062
                    Width =1140
                    FontSize =9
                    FontWeight =300
                    TabIndex =4
                    ForeColor =0
                    Name ="btnPrint"
                    Caption ="Skriv ut"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5126
                    LayoutCachedTop =4062
                    LayoutCachedWidth =6266
                    LayoutCachedHeight =4422
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4604
                    Top =236
                    Width =1648
                    Height =267
                    FontSize =9
                    TabIndex =5
                    Name ="cboSem"
                    RowSourceType ="Value List"
                    RowSource ="Høst;Vår;Høst og vår"

                    LayoutCachedLeft =4604
                    LayoutCachedTop =236
                    LayoutCachedWidth =6252
                    LayoutCachedHeight =503
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3590
                            Top =243
                            Width =945
                            Height =240
                            FontSize =9
                            Name ="Label27"
                            Caption ="Semester:"
                            LayoutCachedLeft =3590
                            LayoutCachedTop =243
                            LayoutCachedWidth =4535
                            LayoutCachedHeight =483
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =543
                    Top =288
                    Width =200
                    Height =215
                    TabIndex =6
                    Name ="chkAlle"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =543
                    LayoutCachedTop =288
                    LayoutCachedWidth =743
                    LayoutCachedHeight =503
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =796
                            Top =243
                            Width =690
                            Height =240
                            Name ="Label34"
                            Caption ="Velg alle"
                            LayoutCachedLeft =796
                            LayoutCachedTop =243
                            LayoutCachedWidth =1486
                            LayoutCachedHeight =483
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
    
Public Sub GenerateReport2(strView As String)
On Error GoTo Err_GenerateReport2
    
    Dim stDocName As String
    Dim stCriteria As String
    Dim sqlAnalysed As String
    Dim strSem As String
    Dim myList As ListBox
    Dim i As Integer, j As Integer
    Dim strGruppe As String, intGruppe As Integer
    Const strH = "H"
    Const strV = "V"
    Const strHV = "H+V"
    Const strVH = "V+H"
    
    stDocName = "rptBelastning"
    'Semester
    Select Case Me.cboSem
        Case "Høst"
            strSem = "Semester = '" & strH & "' OR Semester = '" & strHV & "' OR Semester = '" & strVH & "'"
        Case "Vår"
            strSem = "Semester = '" & strV & "' OR Semester = '" & strHV & "' OR Semester = '" & strVH & "'"
        Case "Høst og vår"
            strSem = "Semester = '" & strH & "' OR Semester = '" & strV & "' OR Semester = '" & strHV & "' OR Semester = '" & strVH & "'"
        Case Else
            strSem = ""
    End Select
    
    
    If chkAlle.Value = True Then 'alle faggrupper
        GL_SELECTION = "Utvalg = Alle"
        If strSem <> "" Then
            stCriteria = "(" & strSem & ")"
        End If
    Else
        Set myList = Me!lstGrupper
        strGruppe = myList.Column(1)
        intGruppe = myList.Column(0)
        stCriteria = "FagID = " & intGruppe
        GL_SELECTION = "Utvalg = " & strGruppe
        If strSem <> "" Then
            stCriteria = "(" & strSem & ") AND " & stCriteria
        End If
    End If
    Select Case strView
        Case "Print"
            DoCmd.OpenReport stDocName, acViewNormal, WhereCondition:=stCriteria
        Case "Preview"
            DoCmd.OpenReport stDocName, acViewPreview, WhereCondition:=stCriteria
    End Select
Exit_GenerateReport2:
    Exit Sub

Err_GenerateReport2:
    MsgBox Err.Description
    Resume Exit_GenerateReport2
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

Private Sub chkAlle_Click()
    Dim i As Integer
    For i = 1 To Me.lstGrupper.ListCount
        Me.lstGrupper.Selected(i) = True
    Next i
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load

    Dim sqlSource As String
    sqlSource = "SELECT * FROM tblGruppe WHERE GruppeID <= 5 ORDER BY GruppeID;"
    Me.lstGrupper.RowSource = sqlSource
    Me.chkAlle = False
Exit_Form_Load:
    Exit Sub

Err_Form_Load:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_Form_Load
    
End Sub

Private Sub btnPrint_Click()
On Error GoTo Err_btnPrint_Click


   Call GenerateReport2("Print")
   
Exit_btnPrint_Click:
    Exit Sub

Err_btnPrint_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnPrint_Click
    
End Sub
Private Sub btnPreview_Click()
On Error GoTo Err_btnPreview_Click


   Call GenerateReport2("Preview")

Exit_btnPreview_Click:
    Exit Sub

Err_btnPreview_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnPreview_Click
    
End Sub



Private Sub lstGrupper_Click()
    Me.chkAlle = False
End Sub
