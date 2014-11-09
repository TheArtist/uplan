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
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10204
    DatasheetFontHeight =10
    ItemSuffix =59
    Left =480
    Top =420
    Right =11190
    Bottom =6945
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2059e7521c5ce340
    End
    Caption ="Emnedata"
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
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
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Section
            Height =5732
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2844
                    Top =623
                    Width =6624
                    Height =341
                    FontSize =9
                    Name ="txtNavn"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =2844
                    LayoutCachedTop =623
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =964
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =570
                            Top =660
                            Width =1920
                            Height =240
                            FontSize =9
                            Name ="Label1"
                            Caption ="Etternavn, fornavn:"
                            LayoutCachedLeft =570
                            LayoutCachedTop =660
                            LayoutCachedWidth =2490
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7419
                    Top =2475
                    Width =2049
                    Height =341
                    FontSize =9
                    TabIndex =1
                    Name ="txtFag"

                    LayoutCachedLeft =7419
                    LayoutCachedTop =2475
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =2816
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =2492
                            Width =1575
                            Height =330
                            FontSize =9
                            Name ="Label11"
                            Caption ="Fagtilhørighet:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =2492
                            LayoutCachedWidth =2141
                            LayoutCachedHeight =2822
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8107
                    Top =4818
                    Width =1526
                    Height =397
                    FontSize =9
                    TabIndex =2
                    Name ="btnOk"
                    Caption ="Ok"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8107
                    LayoutCachedTop =4818
                    LayoutCachedWidth =9633
                    LayoutCachedHeight =5215
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =453
                    Top =4818
                    Width =1526
                    Height =397
                    FontSize =9
                    TabIndex =3
                    Name ="btnCancel"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =453
                    LayoutCachedTop =4818
                    LayoutCachedWidth =1979
                    LayoutCachedHeight =5215
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7419
                    Top =2012
                    Width =2049
                    Height =341
                    TabIndex =4
                    Name ="cboAlder"
                    RowSourceType ="Value List"
                    RowSource ="\"H\";\"V\";\"H+V\";\"V+H\";\"U\""

                    LayoutCachedLeft =7419
                    LayoutCachedTop =2012
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =2353
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =2035
                            Width =1350
                            Height =330
                            FontSize =9
                            Name ="Label18"
                            Caption ="Aldersgruppe:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =2035
                            LayoutCachedWidth =1916
                            LayoutCachedHeight =2365
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2244
                    Top =3401
                    Width =7224
                    Height =341
                    FontSize =9
                    TabIndex =5
                    Name ="txtMerknader"

                    LayoutCachedLeft =2244
                    LayoutCachedTop =3401
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =3742
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =3400
                            Width =1515
                            Height =330
                            FontSize =9
                            Name ="Label20"
                            Caption ="Merknader:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =3400
                            LayoutCachedWidth =2081
                            LayoutCachedHeight =3730
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7419
                    Top =2938
                    Width =2049
                    Height =341
                    FontSize =9
                    TabIndex =6
                    Name ="txtSprak"

                    LayoutCachedLeft =7419
                    LayoutCachedTop =2938
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =3279
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =2931
                            Width =2040
                            Height =330
                            FontSize =9
                            Name ="Label22"
                            Caption ="Språk arbeidsplan:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =2931
                            LayoutCachedWidth =2606
                            LayoutCachedHeight =3261
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =453
                    Top =396
                    Width =9128
                    Height =4129
                    Name ="Box42"
                    LayoutCachedLeft =453
                    LayoutCachedTop =396
                    LayoutCachedWidth =9581
                    LayoutCachedHeight =4525
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6462
                    Top =4818
                    Width =1526
                    Height =397
                    FontSize =9
                    TabIndex =7
                    Name ="btnNew"
                    Caption ="Ny foreleser"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6462
                    LayoutCachedTop =4818
                    LayoutCachedWidth =7988
                    LayoutCachedHeight =5215
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =2244
                    Top =3864
                    Width =7224
                    Height =341
                    FontSize =9
                    TabIndex =8
                    Name ="txtEpost"

                    LayoutCachedLeft =2244
                    LayoutCachedTop =3864
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =4205
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =566
                            Top =3863
                            Width =1515
                            Height =330
                            FontSize =9
                            Name ="Label51"
                            Caption ="Epost:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =3863
                            LayoutCachedWidth =2081
                            LayoutCachedHeight =4193
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2211
                    Top =4818
                    Width =1526
                    Height =397
                    FontSize =9
                    TabIndex =9
                    Name ="Command54"
                    Caption ="Lukk"

                    LayoutCachedLeft =2211
                    LayoutCachedTop =4818
                    LayoutCachedWidth =3737
                    LayoutCachedHeight =5215
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =7419
                    Top =1086
                    Width =2049
                    Height =341
                    TabIndex =10
                    Name ="cboStilling"
                    RowSourceType ="Value List"
                    RowSource ="\"H\";\"V\";\"H+V\";\"V+H\";\"U\""

                    LayoutCachedLeft =7419
                    LayoutCachedTop =1086
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =1427
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =566
                            Top =1109
                            Width =1350
                            Height =330
                            FontSize =9
                            Name ="Label56"
                            Caption ="Stilling:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =1109
                            LayoutCachedWidth =1916
                            LayoutCachedHeight =1439
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =7419
                    Top =1549
                    Width =2049
                    Height =341
                    TabIndex =11
                    Name ="cboAndel"
                    RowSourceType ="Value List"
                    RowSource ="\"H\";\"V\";\"H+V\";\"V+H\";\"U\""

                    LayoutCachedLeft =7419
                    LayoutCachedTop =1549
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =1890
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =566
                            Top =1572
                            Width =1350
                            Height =330
                            FontSize =9
                            Name ="Label58"
                            Caption ="Stillingsandel:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =1572
                            LayoutCachedWidth =1916
                            LayoutCachedHeight =1902
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

Private Sub btnNew_Click()
        Me.btnOk.Caption = "Legg til"
        Me.Caption = "Nytt emne"
        Me.txtEkode = ""
        Me.txtEnavn = ""
        Me.txtStp = ""
        Me.cboSemester = ""
        Me.txtSted = ""
        Me.txtMerknader = ""
        Me.txtURL = ""
        Me.chkAktiv.Value = False
        Me.txtTimer = 0
        Me.txtUker = 0
        Me.txtUndervisning = 0
        Me.txtSumTimer = 0
        Me.txtEkode.SetFocus
End Sub

Private Sub btnOk_Click()
On Error GoTo Err_btnOk_Click
    Dim myDb As DAO.Database
    Dim myRS As DAO.Recordset
    Dim sqlUpdate As String, sqlInsert As String
    Dim blnAktiv As Boolean
    Dim lngFagID As Long, lngEID As Long
    blnAktiv = Me.chkAktiv.Value
' Kontrollerer data, legger inn default data hvis noe mangler
' Emnekode
    If IsNull(Me.txtEkode) Or Me.txtEkode = "" Then
        MsgBox "Emnekode må oppgis", vbExclamation, OIS_Title
        GoTo Exit_btnOk_Click
    Else
        Me.txtEkode = UCase(Me.txtEkode) 'store bokstaver i emnekode
    End If
' Studiepoeng
    If IsNull(Me.txtStp) Or Me.txtStp = "" Then Me.txtStp = "7,5"
' Semester
    If IsNull(Me.cboSemester) Or Me.cboSemester = "" Then Me.cboSemester = "H"
' Sted
    If IsNull(Me.txtSted) Or Me.txtSted = "" Then Me.txtSted = "Molde"
' Undervisning
    If IsNull(Me.txtUndervisning) Or Me.txtUndervisning = "" Then Me.txtUndervisning = "0"
' Sum belastningstimer
    If IsNull(Me.txtSumTimer) Or Me.txtSumTimer = "" Then Me.txtSumTimer = "0"
' Fagkode
    lngFagID = HentFag(Left(Me.txtEkode, 3))
    If lngFagID = 0 Then
        MsgBox "Emnekoden er ugyldig. Kan ikke lagre dette emnet", vbExclamation, OIS_Title
        GoTo Exit_btnOk_Click
    End If
' Lagrer data
    Set myDb = CurrentDb()
    Select Case Me.btnOk.Caption
        Case Is = "Ok"
            sqlUpdate = "SELECT * FROM tblEmne WHERE EmneID = " & GL_EID & ";"
            Set myRS = myDb.OpenRecordset(sqlUpdate, dbOpenDynaset)
            If myRS.RecordCount > 0 Then
                myRS.Edit
                myRS!Emnenavn = Me.txtEnavn
                myRS!Studiepoeng = CSng(Me.txtStp)
                myRS!Semester = Me.cboSemester
                myRS!Sted = Me.txtSted
                myRS!Aktiv = blnAktiv
                myRS!Comment = Me.txtMerknader
                myRS!UTimer = CLng(Me.txtUndervisning)
                myRS!TotalTimer = CLng(Me.txtSumTimer)
                myRS!pageURL = Me.txtURL
                myRS!FagID = lngFagID
                myRS.Update
            Else
                msg1 = "Ukjent emne. Kan ikke oppdateres."
                MsgBox msg1, vbExclamation + vbOKOnly, OIS_Title
                Exit Sub
            End If
           
        Case Is = "Legg til"
            sqlInsert = "SELECT * FROM tblEmne ORDER BY EmneID;"
            Set myRS = myDb.OpenRecordset(sqlInsert, dbOpenDynaset)
            myRS.AddNew
            myRS!Emnekode = Me.txtEkode
            myRS!Emnenavn = Me.txtEnavn
            myRS!Studiepoeng = CSng(Me.txtStp)
            myRS!Semester = Me.cboSemester
            myRS!Sted = Me.txtSted
            myRS!Aktiv = blnAktiv
            myRS!Comment = Me.txtMerknader
            myRS!UTimer = CLng(Me.txtUndervisning)
            myRS!TotalTimer = CLng(Me.txtSumTimer)
            myRS!pageURL = Me.txtURL
            myRS!FagID = lngFagID
            myRS.Update
        End Select
Exit_btnOk_Click:
    Me.btnOk.Caption = "Ok"
    Me.btnNew.Enabled = True
    DoCmd.Close
    Call Form_frmMain.ListviewFill(GL_GID, GL_ECODE)
    Exit Sub

Err_btnOk_Click:
    MsgBox Err.Description
    Resume Exit_btnOk_Click
End Sub
Private Sub btnCancel_Click()
On Error GoTo Err_btnCancel_Click
    GL_MODE = ""
    DoCmd.Close

Exit_btnCancel_Click:
    Exit Sub

Err_btnCancel_Click:
    MsgBox Err.Description
    Resume Exit_btnCancel_Click
    
End Sub

Private Sub Form_Load()
    Dim myDb As DAO.Database
    Dim myRS As DAO.Recordset
    Dim sqlStr As String, strCaption As String
    Dim NoOfEmne As Integer
    Set myDb = CurrentDb()
    sqlStr = "SELECT * FROM tblEmne WHERE tblEmne.EmneID= " & GL_EID & ";"
    
    Set myRS = myDb.OpenRecordset(sqlStr, dbOpenDynaset)
    If GL_MODE = "Edit" Then
        If myRS.RecordCount > 0 Then 'record found
            strCaption = "Emnedata for " & myRS!Emnekode & " " & myRS!Emnenavn & " (" & myRS!Sted & ")"
            Me.Caption = strCaption
            Me.txtEkode = myRS!Emnekode
            Me.txtEnavn = myRS!Emnenavn
            Me.txtStp = myRS!Studiepoeng
            Me.cboSemester = myRS!Semester
            Me.txtSted = myRS!Sted
            Me.txtMerknader = myRS!Comment
            Me.txtURL = myRS!pageURL
            Me.chkAktiv.Value = myRS!Aktiv
            Me.txtTimer = 0
            Me.txtUker = 0
            If Not IsNull(myRS!UTimer) And myRS!UTimer <> "" Then
                Me.txtUndervisning = myRS!UTimer
            Else
                Me.txtUndervisning = 0
            End If
            If Not IsNull(myRS!TotalTimer) And myRS!TotalTimer <> "" Then
                Me.txtSumTimer = myRS!TotalTimer
            Else
                Me.txtSumTimer = 0
            End If
            Me.btnOk.Caption = "Ok"
        End If
    ElseIf GL_MODE = "New" Then
        Call btnNew_Click
        btnNew.Enabled = False
        'MsgBox "Du må velge et emne"
        'Exit Sub
    End If
End Sub


Private Sub txtTimer_AfterUpdate()
    Me.txtUndervisning = CInt(Me.txtTimer) * CInt(Me.txtUker)
End Sub


Private Sub txtUker_AfterUpdate()
    Me.txtUndervisning = CInt(Me.txtTimer) * CInt(Me.txtUker)
End Sub


Public Function HentFag(strE As String) As Long
    Dim lngK As Long
    Dim myDb As DAO.Database
    Dim rsFag As DAO.Recordset
    Dim i As Integer, Funnet As Boolean
    Set myDb = CurrentDb()
    Set rsFag = myDb.OpenRecordset("tblEmnekode", dbOpenDynaset)
    Funnet = False
    Do While Not rsFag.EOF
        If strE = rsFag!Kode Then
            lngK = rsFag!GruppeID
            Funnet = True
            GoTo HentFag_Exit
        End If
        rsFag.MoveNext
    Loop
HentFag_Exit:
    If Not Funnet Then lngK = 0
    HentFag = lngK
    rsFag.Close
End Function
