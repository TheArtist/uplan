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
    Width =10231
    DatasheetFontHeight =10
    ItemSuffix =26
    Left =1530
    Top =225
    Right =11580
    Bottom =6705
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x84ace36f9a5de240
    End
    Caption ="Emnekoder og beskrivelse"
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
        Begin CheckBox
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
            Height =6169
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    ColumnCount =6
                    Left =396
                    Top =623
                    Width =8898
                    Height =4074
                    FontSize =9
                    Name ="lstStudier"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="859;851;4542;965;965;0"
                    OnClick ="[Event Procedure]"
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =315
                    Top =225
                    Width =9152
                    Height =4980
                    TabIndex =1
                    Name ="frStudier"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =510
                            Top =105
                            Width =750
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            Name ="lblPlants"
                            Caption ="Studier"
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =8560
                            Top =4848
                            Width =907
                            Height =170
                            OptionValue =1
                            Name ="chkAktiv"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =8790
                                    Top =4818
                                    Width =561
                                    Height =228
                                    FontSize =9
                                    Name ="Label24"
                                    Caption ="Aktiv"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =2152
                    Top =4815
                    Width =4473
                    Height =285
                    FontSize =9
                    TabIndex =2
                    Name ="txtStudienavn"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =283
                    Top =5329
                    Width =2205
                    Height =360
                    FontSize =9
                    TabIndex =3
                    Name ="btnClose"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2607
                    Top =5329
                    Width =2205
                    Height =360
                    FontSize =9
                    TabIndex =4
                    Name ="btnNew"
                    Caption ="Ny"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7255
                    Top =5329
                    Width =2205
                    Height =360
                    FontSize =9
                    TabIndex =5
                    Name ="btnUpdate"
                    Caption ="Oppdater"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4931
                    Top =5329
                    Width =2205
                    Height =360
                    FontSize =9
                    TabIndex =6
                    Name ="btnDelete"
                    Caption ="Slett"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    Left =396
                    Top =4815
                    Width =834
                    Height =285
                    FontSize =9
                    TabIndex =7
                    Name ="txtStudieID"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =6692
                    Top =4815
                    Width =894
                    Height =285
                    FontSize =9
                    TabIndex =8
                    Name ="txtGruppe"
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =7653
                    Top =4815
                    Width =849
                    Height =285
                    FontSize =9
                    TabIndex =9
                    Name ="txtStudieURL"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    Left =1247
                    Top =4818
                    Width =849
                    Height =285
                    FontSize =9
                    TabIndex =10
                    Name ="txtKode"
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
Public LP_Skode As String
Public LP_Grid As Long

Private Sub btnDelete_Click()
On Error GoTo Err_btnDelete_Click
    Dim msg As String
    Dim myDb                    As DAO.Database
    Dim sqlStudium              As String

    Dim response As Integer
    If IsNull(lstStudier) Then
        msg = "Ingenting å slette. Du må velge et studium."
        MsgBox msg, vbOKOnly + vbExclamation, OIS_Title
        Exit Sub
    End If
    
    msg = "Vil du slette '" & Me.lstStudier.Column(1) & "'?"
    response = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, OIS_Title)
    If response = vbNo Then
        Exit Sub
    End If
    
    Set myDb = CurrentDb
    sqlStudium = "DELETE * FROM tblStudium WHERE StudieID = '" & Me.txtStudieID & "';"
    
    myDb.Execute (sqlStudium)
 
    Call LoadListbox
Exit_btnDelete_Click:
    Exit Sub

Err_btnDelete_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnDelete_Click

End Sub


Private Sub btnNew_Click()
On Error GoTo Err_btnNew_Click

    Me.btnUpdate.Caption = strInsert
    Me.txtStudieID = ""
    Me.txtKode = ""
    Me.txtStudienavn.SetFocus
    Me.txtStudienavn = ""
    Me.txtGruppe = ""
    Me.txtStudieURL = ""
    Me.frStudier = 0

Exit_btnNew_Click:
    Exit Sub

Err_btnNew_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnNew_Click
End Sub

Private Sub btnOppdaterEmne_Click()
    Call UpdateGruppeID
End Sub

Private Sub btnUpdate_Click()
On Error GoTo Err_btnUpdate_Click
    Dim myDb                As DAO.Database
    Dim sqlStudium          As String
    Dim sqlSource           As String
    Dim sqlInsert           As String
    Dim sqlUpdate           As String
    Dim blnAktiv            As Boolean
    
    Dim msg As String
    If txtStudieID = "" And txtStudienavn = "" Then
        msg = "Ingenting å oppdatere. Dy må velge et studium "
        MsgBox msg, vbOKOnly + vbExclamation, OIS_Title
        Exit Sub
    End If

    Set myDb = CurrentDb
    LP_Skode = CLng(Me.txtStudieID)
    LP_Grid = CLng(Me.txtGruppe)
    If Me.frStudier = 1 Then  'aktiv
        blnAktiv = True
    Else
        blnAktiv = False
    End If
    If Me.btnUpdate.Caption = strUpdate Then
        sqlUpdate = " SET Studienavn = '" & Me.txtStudienavn & "', Studiekode = '" & Me.txtKode & "', StudieURL = '" & Me.txtStudieURL & "', GruppeID = " & LP_Grid & _
        ", Aktiv = " & blnAktiv & " WHERE StudieID = '" & LP_Skode & ";"
        sqlStudium = "UPDATE tblStudium" & sqlUpdate & "';"
        myDb.Execute (sqlStudium)
    ElseIf Me.btnUpdate.Caption = strInsert Then
        sqlInsert = " (Studienavn, Studiekode, GruppeID, StudieURL, Aktiv ) " & _
        " VALUES ('" & Me.txtStudienavn & "', '" & Me.txtKode & "', " & LP_Grid & ", '" & Me.txtStudieURL & "', " & blnAktiv & ");"
        sqlStudium = "INSERT INTO tblStudium" & sqlInsert & ";"
        myDb.Execute (sqlStudium)
    End If
    
    Call LoadListbox
Exit_btnUpdate_Click:
    Exit Sub

Err_btnUpdate_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnUpdate_Click
End Sub

Private Sub Form_Load()
'
    GP_SysMgr = True
    Call LoadListbox
    If Not GP_SysMgr Then
        Me.btnDelete.Enabled = False
        'Me.btnEdit.Enabled = False
        Me.btnNew.Enabled = False
        Me.btnUpdate.Enabled = False
        Me.txtStudieID.Locked = True
        Me.txtStudienavn.Locked = True
    End If
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


Private Sub lstStudier_Click()
On Error GoTo Err_lstStudier_Click

    LP_Skode = Me.lstStudier.Column(0)
    Me!txtStudienavn.SetFocus
    Me.txtStudieID = Me.lstStudier.Column(0)
    Me.txtKode = Me.lstStudier.Column(1)
    Me.txtStudienavn = Me.lstStudier.Column(2)
    Me.txtGruppe = Me.lstStudier.Column(3)
    Me.txtStudieURL = Me.lstStudier.Column(4)
    If Me.lstStudier.Column(5) = "-1" Then  'Aktiv
        Me.frStudier = 1
    Else
        Me.frStudier = 0
    End If
'If Not GP_SysMgr Then Exit Sub
    Me.btnUpdate.Caption = "Oppdater"

Exit_lstStudier_Click:
    Exit Sub

Err_lstStudier_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_lstStudier_Click

End Sub


Public Sub LoadListbox()
    Dim sqlSource As String
    sqlSource = "SELECT StudieID, Studiekode, Studienavn, GruppeID, StudieURL, Aktiv FROM tblStudium ORDER BY StudieID;"
    Me.lstStudier.RowSource = sqlSource
    Me.btnUpdate.Caption = "Oppdater"
    Me.txtStudienavn = ""
    Me.txtKode = ""
    Me.txtStudieID = ""
    Me.txtGruppe = ""
    Me.frStudier = 0
End Sub
Private Sub txtEmnekode_Click()
On Error GoTo Err_txtEmnekode_Click

    If Not GP_SysMgr Then Exit Sub
        Me.btnUpdate.Enabled = True
        If IsNull(Me.txtStudieID) Or Me.txtStudieID = "" Then
            Me.btnUpdate.Caption = strInsert
        Else
            Me.btnUpdate.Caption = "Oppdater"
        End If

Exit_txtEmnekode_Click:
    Exit Sub

Err_txtEmnekode_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_txtEmnekode_Click

End Sub
Public Sub UpdateGruppeID()
On Error GoTo Err_UpdateGruppeID
    Dim myDb As DAO.Database
    Dim MyEmneRs As DAO.Recordset
    Dim MyKodeRs As DAO.Recordset
    Dim msg1 As String, msg2 As String
    Dim Etype As String
    Dim Grid As Long
    Dim sqlStr As String
    Set myDb = CurrentDb
    
    Set MyKodeRs = myDb.OpenRecordset("tblEmnekode", dbOpenDynaset)
    Do While Not MyKodeRs.EOF
        Etype = Left(MyKodeRs!Kode, 3)
        Grid = MyKodeRs!GruppeID
        sqlStr = "UPDATE tblEmne SET tblEmne.FagID =" & Grid & _
                " WHERE Left(Emnekode,3) = '" & Etype & "';"
        myDb.Execute (sqlStr)
        MyKodeRs.MoveNext
    Loop
    msg = "Oppdatering av gruppekoder utført"
    MsgBox msg, vbInformation, OIS_Title
Exit_UpdateGruppeID:
    Exit Sub

Err_UpdateGruppeID:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_UpdateGruppeID


End Sub
