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
    Width =9864
    DatasheetFontHeight =10
    ItemSuffix =46
    Left =480
    Top =420
    Right =10950
    Bottom =7215
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
            Height =5725
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7415
                    Top =623
                    Width =2049
                    Height =341
                    FontSize =9
                    Name ="txtEkode"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =7415
                    LayoutCachedTop =623
                    LayoutCachedWidth =9464
                    LayoutCachedHeight =964
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =664
                            Width =1290
                            Height =330
                            FontSize =9
                            Name ="Label1"
                            Caption ="Emnekode:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =664
                            LayoutCachedWidth =1856
                            LayoutCachedHeight =994
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2244
                    Top =1059
                    Width =7224
                    Height =341
                    FontSize =9
                    TabIndex =1
                    Name ="txtEnavn"

                    LayoutCachedLeft =2244
                    LayoutCachedTop =1059
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =1400
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =1094
                            Width =1515
                            Height =330
                            FontSize =9
                            Name ="Label7"
                            Caption ="Emnenavn:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =1094
                            LayoutCachedWidth =2081
                            LayoutCachedHeight =1424
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7415
                    Top =1495
                    Width =2049
                    Height =341
                    FontSize =9
                    TabIndex =2
                    Name ="txtStp"

                    LayoutCachedLeft =7415
                    LayoutCachedTop =1495
                    LayoutCachedWidth =9464
                    LayoutCachedHeight =1836
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =1524
                            Width =1815
                            Height =330
                            FontSize =9
                            Name ="Label9"
                            Caption ="Studiepoeng:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =1524
                            LayoutCachedWidth =2381
                            LayoutCachedHeight =1854
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7415
                    Top =2367
                    Width =2049
                    Height =341
                    FontSize =9
                    TabIndex =3
                    Name ="txtSted"

                    LayoutCachedLeft =7415
                    LayoutCachedTop =2367
                    LayoutCachedWidth =9464
                    LayoutCachedHeight =2708
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =2384
                            Width =1575
                            Height =330
                            FontSize =9
                            Name ="Label11"
                            Caption ="Sted:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =2384
                            LayoutCachedWidth =2141
                            LayoutCachedHeight =2714
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8107
                    Top =5102
                    Width =1526
                    Height =397
                    FontSize =9
                    TabIndex =4
                    Name ="btnOk"
                    Caption ="Ok"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8107
                    LayoutCachedTop =5102
                    LayoutCachedWidth =9633
                    LayoutCachedHeight =5499
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =453
                    Top =5102
                    Width =1526
                    Height =397
                    FontSize =9
                    TabIndex =5
                    Name ="btnCancel"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =453
                    LayoutCachedTop =5102
                    LayoutCachedWidth =1979
                    LayoutCachedHeight =5499
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
                    Left =7415
                    Top =1931
                    Width =2049
                    Height =341
                    TabIndex =6
                    Name ="cboSemester"
                    RowSourceType ="Value List"
                    RowSource ="\"H\";\"V\";\"H+V\";\"V+H\";\"U\""

                    LayoutCachedLeft =7415
                    LayoutCachedTop =1931
                    LayoutCachedWidth =9464
                    LayoutCachedHeight =2272
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =1954
                            Width =1350
                            Height =330
                            FontSize =9
                            Name ="Label18"
                            Caption ="Semester:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =1954
                            LayoutCachedWidth =1916
                            LayoutCachedHeight =2284
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2244
                    Top =3675
                    Width =7224
                    Height =341
                    FontSize =9
                    TabIndex =7
                    Name ="txtMerknader"

                    LayoutCachedLeft =2244
                    LayoutCachedTop =3675
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =4016
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =3674
                            Width =1515
                            Height =330
                            FontSize =9
                            Name ="Label20"
                            Caption ="Merknader:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =3674
                            LayoutCachedWidth =2081
                            LayoutCachedHeight =4004
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7415
                    Top =4111
                    Width =2049
                    Height =341
                    FontSize =9
                    TabIndex =8
                    Name ="txtURL"

                    LayoutCachedLeft =7415
                    LayoutCachedTop =4111
                    LayoutCachedWidth =9464
                    LayoutCachedHeight =4452
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =570
                            Top =4104
                            Width =2040
                            Height =330
                            FontSize =9
                            Name ="Label22"
                            Caption ="URL til studiehåndbok:"
                            LayoutCachedLeft =570
                            LayoutCachedTop =4104
                            LayoutCachedWidth =2610
                            LayoutCachedHeight =4434
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7415
                    Top =2803
                    Width =2049
                    Height =341
                    FontSize =9
                    TabIndex =9
                    Name ="txtUndervisning"

                    LayoutCachedLeft =7415
                    LayoutCachedTop =2803
                    LayoutCachedWidth =9464
                    LayoutCachedHeight =3144
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =2814
                            Width =1575
                            Height =330
                            FontSize =9
                            Name ="Label26"
                            Caption ="Undervisning:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =2814
                            LayoutCachedWidth =2141
                            LayoutCachedHeight =3144
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7415
                    Top =3239
                    Width =2049
                    Height =341
                    FontSize =9
                    TabIndex =10
                    Name ="txtSumTimer"

                    LayoutCachedLeft =7415
                    LayoutCachedTop =3239
                    LayoutCachedWidth =9464
                    LayoutCachedHeight =3580
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =3244
                            Width =1950
                            Height =330
                            FontSize =9
                            Name ="Label28"
                            Caption ="Sum timetall for emne:"
                            LayoutCachedLeft =566
                            LayoutCachedTop =3244
                            LayoutCachedWidth =2516
                            LayoutCachedHeight =3574
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5952
                    Top =2803
                    Width =794
                    Height =341
                    TabIndex =11
                    Name ="txtUker"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Antall uker eller samlinger med undervisning"

                    LayoutCachedLeft =5952
                    LayoutCachedTop =2803
                    LayoutCachedWidth =6746
                    LayoutCachedHeight =3144
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4762
                    Top =2803
                    Width =794
                    Height =341
                    TabIndex =12
                    Name ="txtTimer"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Undervisinngstimer per uke eller per samling"

                    LayoutCachedLeft =4762
                    LayoutCachedTop =2803
                    LayoutCachedWidth =5556
                    LayoutCachedHeight =3144
                End
                Begin Label
                    OverlapFlags =93
                    Left =5672
                    Top =2834
                    Width =224
                    Height =284
                    Name ="Label32"
                    Caption ="X"
                    LayoutCachedLeft =5672
                    LayoutCachedTop =2834
                    LayoutCachedWidth =5896
                    LayoutCachedHeight =3118
                End
                Begin Label
                    OverlapFlags =93
                    Left =6859
                    Top =2834
                    Width =227
                    Height =284
                    Name ="Label33"
                    Caption ="="
                    LayoutCachedLeft =6859
                    LayoutCachedTop =2834
                    LayoutCachedWidth =7086
                    LayoutCachedHeight =3118
                End
                Begin Label
                    OverlapFlags =93
                    Left =4875
                    Top =3174
                    Width =572
                    Height =283
                    Name ="Label34"
                    Caption ="(timer)"
                    LayoutCachedLeft =4875
                    LayoutCachedTop =3174
                    LayoutCachedWidth =5447
                    LayoutCachedHeight =3457
                End
                Begin Label
                    OverlapFlags =93
                    Left =5782
                    Top =3174
                    Width =1245
                    Height =285
                    Name ="Label35"
                    Caption ="(uker/samlinger)"
                    LayoutCachedLeft =5782
                    LayoutCachedTop =3174
                    LayoutCachedWidth =7027
                    LayoutCachedHeight =3459
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =453
                    Top =396
                    Width =9128
                    Height =4579
                    Name ="Box42"
                    LayoutCachedLeft =453
                    LayoutCachedTop =396
                    LayoutCachedWidth =9581
                    LayoutCachedHeight =4975
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =7426
                    Top =4622
                    Width =1474
                    Height =284
                    TabIndex =13
                    Name ="chkAktiv"

                    LayoutCachedLeft =7426
                    LayoutCachedTop =4622
                    LayoutCachedWidth =8900
                    LayoutCachedHeight =4906
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7667
                            Top =4592
                            Width =690
                            Height =240
                            FontSize =9
                            Name ="Label44"
                            Caption ="Aktiv"
                            LayoutCachedLeft =7667
                            LayoutCachedTop =4592
                            LayoutCachedWidth =8357
                            LayoutCachedHeight =4832
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6462
                    Top =5102
                    Width =1526
                    Height =397
                    FontSize =9
                    TabIndex =14
                    Name ="btnNew"
                    Caption ="Nytt emne"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6462
                    LayoutCachedTop =5102
                    LayoutCachedWidth =7988
                    LayoutCachedHeight =5499
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
