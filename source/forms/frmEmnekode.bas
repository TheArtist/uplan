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
    ItemSuffix =21
    Left =1530
    Top =225
    Right =10650
    Bottom =6900
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x84ace36f9a5de240
    End
    Caption ="Emnekoder og beskrivelse"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
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
            BorderLineStyle =0
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
            Width =1701
            LabelX =-1701
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
        Begin Section
            Height =6169
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    ColumnCount =3
                    Left =396
                    Top =623
                    Width =7638
                    Height =4074
                    FontSize =10
                    Name ="lstCode"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1701;4649;1134"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"

                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =315
                    Top =225
                    Width =7848
                    Height =4980
                    TabIndex =1
                    Name ="Frame2"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =510
                            Top =105
                            Width =951
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            Name ="lblPlants"
                            Caption ="Emnekoder"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =2089
                    Top =4815
                    Width =4602
                    Height =285
                    FontSize =9
                    TabIndex =2
                    Name ="txtEmne"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =283
                    Top =5329
                    Width =1305
                    Height =360
                    FontSize =9
                    TabIndex =3
                    Name ="btnClose"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =283
                    LayoutCachedTop =5329
                    LayoutCachedWidth =1588
                    LayoutCachedHeight =5689
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1690
                    Top =5329
                    Width =1305
                    Height =360
                    FontSize =9
                    TabIndex =4
                    Name ="btnNew"
                    Caption ="Ny"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1690
                    LayoutCachedTop =5329
                    LayoutCachedWidth =2995
                    LayoutCachedHeight =5689
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4504
                    Top =5329
                    Width =1305
                    Height =360
                    FontSize =9
                    TabIndex =5
                    Name ="btnUpdate"
                    Caption ="Oppdater"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4504
                    LayoutCachedTop =5329
                    LayoutCachedWidth =5809
                    LayoutCachedHeight =5689
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3097
                    Top =5329
                    Width =1305
                    Height =360
                    FontSize =9
                    TabIndex =6
                    Name ="btnDelete"
                    Caption ="Slett"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3097
                    LayoutCachedTop =5329
                    LayoutCachedWidth =4402
                    LayoutCachedHeight =5689
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =396
                    Top =4815
                    Width =1638
                    Height =285
                    FontSize =9
                    TabIndex =7
                    Name ="txtEmnekode"
                    OnClick ="[Event Procedure]"

                End
                Begin TextBox
                    OverlapFlags =215
                    Left =6746
                    Top =4815
                    Width =1278
                    Height =285
                    FontSize =9
                    TabIndex =8
                    Name ="txtGruppe"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5878
                    Top =5329
                    Width =2280
                    Height =360
                    FontSize =9
                    TabIndex =9
                    Name ="btnOppdaterEmne"
                    Caption ="Oppdater emnetabell"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5878
                    LayoutCachedTop =5329
                    LayoutCachedWidth =8158
                    LayoutCachedHeight =5689
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
Public LP_Ekode As String
Public LP_Grid As Long

Private Sub btnDelete_Click()
On Error GoTo Err_btnDelete_Click
    Dim msg As String
    Dim myDb                As DAO.Database
    Dim sqlCode            As String

    Dim response As Integer
    If IsNull(lstCodes) Then
        msg = "Ingenting å slette. Du må velge en type."
        MsgBox msg, vbOKOnly + vbExclamation, OIS_Title
        Exit Sub
    End If
    
    msg = "Slette '" & Me.lstCode.Column(1) & "'?"
    response = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, OIS_Title)
    If response = vbNo Then
        Exit Sub
    End If
    
    Set myDb = CurrentDb
    sqlCode = "DELETE * FROM tblEmnekode WHERE Kode = '" & Me.txtEmnekode & "';"
    
    myDb.Execute (sqlCode)
 
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
    Me.txtEmnekode.SetFocus
    Me.txtEmnekode = ""
    Me.txtEmne = ""
    Me.txtGruppe = ""

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
    Dim sqlCode             As String
    Dim sqlSource           As String
    Dim sqlInsert           As String
    Dim sqlUpdate           As String
    
    Dim msg As String
    If txtEmnekode = "" And txtEmne = "" Then
        msg = "Ingenting å oppdatere. Dy må velge en kode "
        MsgBox msg, vbOKOnly + vbExclamation, OIS_Title
        Exit Sub
    End If

    Set myDb = CurrentDb
    LP_Ekode = Me.txtEmnekode
    LP_Grid = CLng(Me.txtGruppe)
    If Me.btnUpdate.Caption = "Oppdater" Then
        sqlUpdate = " SET Kode = '" & LP_Ekode & "', Beskrivelse = '" & Me.txtEmne & "', GruppeID = " & LP_Grid & " WHERE Kode = '" & LP_Ekode
        sqlCode = "UPDATE tblEmnekode" & sqlUpdate & "';"
        myDb.Execute (sqlCode)
    ElseIf Me.btnUpdate.Caption = strInsert Then
        sqlInsert = " (Kode, Beskrivelse, GruppeID) VALUES ('" & LP_Ekode & "', '" & Me.txtEmne & "', " & LP_Grid & ")"
        sqlCode = "INSERT INTO tblEmnekode" & sqlInsert & ";"
        myDb.Execute (sqlCode)
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
        Me.txtEmnekode.Locked = True
        Me.txtEmne.Locked = True
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


Private Sub lstCode_Click()
On Error GoTo Err_lstCode_Click

    LP_Ekode = Me.lstCode.Column(0)
    'Me!txtEmnekode.SetFocus
    Me.txtEmnekode = Me.lstCode.Column(0)
    Me.txtEmne = Me.lstCode.Column(1)
    Me.txtGruppe = Me.lstCode.Column(2)
   
If Not GP_SysMgr Then Exit Sub
    Me.btnUpdate.Caption = "Oppdater"

Exit_lstCode_Click:
    Exit Sub

Err_lstCode_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_lstCode_Click

End Sub


Public Sub LoadListbox()
    Dim sqlSource As String
    sqlSource = "SELECT Kode, Beskrivelse, GruppeID FROM tblEmnekode ORDER BY Kode;"
    Me.lstCode.RowSource = sqlSource
    Me.btnUpdate.Caption = "Oppdater"
    Me.txtEmne = ""
    Me.txtEmnekode = ""
    Me.txtGruppe = ""
End Sub
Private Sub txtEmnekode_Click()
On Error GoTo Err_txtEmnekode_Click

    If Not GP_SysMgr Then Exit Sub
        Me.btnUpdate.Enabled = True
        If IsNull(Me.txtEmnekode) Or Me.txtEmnekode = "" Then
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
