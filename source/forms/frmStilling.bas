Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =255
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8333
    DatasheetFontHeight =10
    ItemSuffix =19
    Left =1830
    Top =285
    Right =10530
    Bottom =6765
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x84ace36f9a5de240
    End
    Caption ="Stillinger"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
            Width =1701
            LabelX =-1701
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
        Begin Section
            Height =5839
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    ColumnCount =2
                    Left =396
                    Top =623
                    Width =7542
                    Height =4074
                    FontSize =9
                    Name ="lstStilling"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1701;2835"
                    OnClick ="[Event Procedure]"
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =315
                    Top =225
                    Width =7728
                    Height =4980
                    TabIndex =1
                    Name ="Frame2"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =516
                            Top =108
                            Width =1704
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            Name ="lblPlants"
                            Caption ="Stillingskatergorier"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =2268
                    Top =4815
                    Width =5670
                    Height =285
                    FontSize =9
                    TabIndex =2
                    Name ="txtStNavn"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =283
                    Top =5329
                    Width =1875
                    Height =360
                    FontSize =9
                    TabIndex =3
                    Name ="btnClose"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2248
                    Top =5329
                    Width =1875
                    Height =360
                    FontSize =9
                    TabIndex =4
                    Name ="btnNew"
                    Caption ="Ny"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6178
                    Top =5329
                    Width =1875
                    Height =360
                    FontSize =9
                    TabIndex =5
                    Name ="btnUpdate"
                    Caption ="Oppdater"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4213
                    Top =5329
                    Width =1875
                    Height =360
                    FontSize =9
                    TabIndex =6
                    Name ="btnDelete"
                    Caption ="Fjern"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =396
                    Top =4815
                    Width =1758
                    Height =285
                    FontSize =9
                    TabIndex =7
                    Name ="txtStKode"
                    OnClick ="[Event Procedure]"
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

Private Sub btnDelete_Click()
On Error GoTo Err_btnDelete_Click
    Dim msg As String
    Dim myDb                As DAO.Database
    Dim sqlPlant            As String
    Dim strKode             As String

    Dim response As Integer
    If IsNull(lstStilling) Then
        msg = "Ingenting å slette. Du må velge en stilling."
        MsgBox msg, vbOKOnly + vbExclamation, OIS_Title
        Exit Sub
    End If
    msg = "Vil du fjerne '" & Me.lstStilling.Column(1) & "'?"
    response = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, OIS_Title)
    If response = vbNo Then
        Exit Sub
    End If
    
    strKode = Me.lstStilling.Column(0)
    Set myDb = CurrentDb
    sqlPlant = "DELETE * FROM tblStilling WHERE StKode = '" & strKode & "';"
    
    myDb.Execute (sqlPlant)
 
    Call LoadListbox
    Me.txtStKode = ""
    Me.txtStNavn = ""
Exit_btnDelete_Click:
    Exit Sub

Err_btnDelete_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnDelete_Click

End Sub


Private Sub btnNew_Click()
On Error GoTo Err_btnNew_Click

    Me.btnUpdate.Caption = strInsert
    Me.txtStKode.SetFocus
    Me.txtStKode = ""
    Me.txtStNavn = ""

Exit_btnNew_Click:
    Exit Sub

Err_btnNew_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnNew_Click
End Sub

Private Sub btnUpdate_Click()
On Error GoTo Err_btnUpdate_Click
    Dim myDb                As DAO.Database
    Dim sqlStilling         As String
    Dim sqlSource           As String
    Dim sqlInsert           As String
    Dim sqlUpdate           As String
    Dim msg As String
    
    If txtStKode = "" Then
        msg = "Ingenting å oppdatere. Du må velge en stilling. "
        MsgBox msg, vbOKOnly + vbExclamation, OIS_Title
        Exit Sub
    End If

    Set myDb = CurrentDb
    
    If Me.btnUpdate.Caption = "Oppdater" Then
        sqlUpdate = " SET StKode = '" & Me.txtStKode & "', StNavn = '" & Me.txtStNavn & _
                    "' WHERE StKode = '" & Me.txtStKode
        sqlStilling = "UPDATE tblStilling" & sqlUpdate & "';"
        myDb.Execute (sqlStilling)
    ElseIf Me.btnUpdate.Caption = strInsert Then
        sqlInsert = " (StKode, StNavn) VALUES ('" & Me.txtStKode & "', '" & Me.txtStNavn & "')"
        sqlStilling = "INSERT INTO tblStilling" & sqlInsert & ";"
        myDb.Execute (sqlStilling)
    End If
    
    Call LoadListbox
    Me.txtStKode = ""
    Me.txtStNavn = ""
Exit_btnUpdate_Click:
    Exit Sub

Err_btnUpdate_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnUpdate_Click
End Sub

Private Sub Form_Load()
'
    'GP_SysMgr = True
    Call LoadListbox
    'If Not GP_SysMgr Then
    '    Me.btnDelete.Enabled = False
    '    'Me.btnEdit.Enabled = False
    '    Me.btnNew.Enabled = False
    '    Me.btnUpdate.Enabled = False
    '    Me.txtStKode.Locked = True
    '    Me.txtStNavn.Locked = True
    'End If
End Sub
Private Sub btnClose_Click()
On Error GoTo Err_btnClose_Click

    DoCmd.Close , , acSaveNo
    'Call setParentTree
Exit_btnClose_Click:
    Exit Sub

Err_btnClose_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnClose_Click
    
End Sub


Private Sub lstStilling_Click()
On Error GoTo Err_lstStilling_Click

    'Me!txtStKode.SetFocus
    Me.txtStKode = Me.lstStilling.Column(0)
    Me.txtStNavn = Me.lstStilling.Column(1)
   
    Me.btnUpdate.Caption = "Oppdater"

Exit_lstStilling_Click:
    Exit Sub

Err_lstStilling_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_lstStilling_Click

End Sub


Public Sub LoadListbox()
    Dim sqlSource As String
    sqlSource = "SELECT StKode, StNavn FROM tblStilling ORDER BY StNavn;"
    Me.lstStilling.RowSource = sqlSource
End Sub
Private Sub txtStKode_Click()
On Error GoTo Err_txtStKode_Click

    Me.btnUpdate.Enabled = True
    If IsNull(Me.txtStKode) Or Me.txtStKode = "" Then
        Me.btnUpdate.Caption = strInsert
    Else
        Me.btnUpdate.Caption = "Oppdater"
    End If

Exit_txtStKode_Click:
    Exit Sub

Err_txtStKode_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_txtStKode_Click

End Sub
