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
    Width =10882
    DatasheetFontHeight =10
    ItemSuffix =41
    Left =420
    Top =345
    Right =11505
    Bottom =8085
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x84ace36f9a5de240
    End
    Caption ="Studier og obligatoriske emner"
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
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Section
            Height =7756
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ListBox
                    OverlapFlags =215
                    ColumnCount =2
                    Left =396
                    Top =1536
                    Width =4473
                    Height =4764
                    ColumnOrder =5
                    FontSize =9
                    Name ="lstEmne"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="851;3402"

                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =225
                    Top =225
                    Width =10368
                    Height =6645
                    ColumnOrder =4
                    TabIndex =1
                    Name ="frOblig"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =225
                    LayoutCachedTop =225
                    LayoutCachedWidth =10593
                    LayoutCachedHeight =6870
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =390
                            Top =105
                            Width =2370
                            Height =240
                            BackColor =-2147483633
                            Name ="lblPlants"
                            Caption ="Studier og obligatoriske emner"
                            LayoutCachedLeft =390
                            LayoutCachedTop =105
                            LayoutCachedWidth =2760
                            LayoutCachedHeight =345
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =1927
                            Top =1247
                            Width =850
                            Height =170
                            OptionValue =1
                            Name ="chkAlle"

                            LayoutCachedLeft =1927
                            LayoutCachedTop =1247
                            LayoutCachedWidth =2777
                            LayoutCachedHeight =1417
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2157
                                    Top =1190
                                    Width =690
                                    Height =240
                                    Name ="Label33"
                                    Caption ="Vis alle"
                                    LayoutCachedLeft =2157
                                    LayoutCachedTop =1190
                                    LayoutCachedWidth =2847
                                    LayoutCachedHeight =1430
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8502
                    Top =6916
                    Width =2040
                    Height =360
                    TabIndex =2
                    Name ="btnClose"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8502
                    LayoutCachedTop =6916
                    LayoutCachedWidth =10542
                    LayoutCachedHeight =7276
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =396
                    Top =510
                    Width =4529
                    Height =340
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =3
                    Name ="cboStudium"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;5670"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =396
                    LayoutCachedTop =510
                    LayoutCachedWidth =4925
                    LayoutCachedHeight =850
                End
                Begin ListBox
                    OverlapFlags =215
                    ColumnCount =3
                    Left =5892
                    Top =1531
                    Width =4479
                    Height =4764
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =4
                    Name ="lstObligemne"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;851;3402"

                    LayoutCachedLeft =5892
                    LayoutCachedTop =1531
                    LayoutCachedWidth =10371
                    LayoutCachedHeight =6295
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5045
                    Top =3231
                    Width =681
                    Height =396
                    TabIndex =5
                    Name ="btnAdd"
                    Caption =">>"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Legg til obligatorisk emne"

                    LayoutCachedLeft =5045
                    LayoutCachedTop =3231
                    LayoutCachedWidth =5726
                    LayoutCachedHeight =3627
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3741
                    Top =1193
                    Width =1131
                    Height =280
                    ColumnOrder =0
                    TabIndex =6
                    Name ="cboEmnekode"
                    RowSourceType ="Table/Query"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="cboEmnekode"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3004
                            Top =1190
                            Width =690
                            Height =240
                            Name ="Label30"
                            Caption ="Vis bare:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =396
                    Top =1190
                    Width =1293
                    Height =240
                    Name ="Label31"
                    Caption ="Alle emner"
                End
                Begin Label
                    OverlapFlags =215
                    Left =5896
                    Top =1190
                    Width =1500
                    Height =240
                    Name ="Label34"
                    Caption ="Obligatoriske emner"
                    LayoutCachedLeft =5896
                    LayoutCachedTop =1190
                    LayoutCachedWidth =7396
                    LayoutCachedHeight =1430
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7532
                    Top =514
                    Width =2841
                    Height =336
                    TabIndex =7
                    Name ="btnStudiehandbok"
                    Caption ="Vis beskrivelse i studiehandbok"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Slå opp i studiehandboka"

                    LayoutCachedLeft =7532
                    LayoutCachedTop =514
                    LayoutCachedWidth =10373
                    LayoutCachedHeight =850
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CustomControl
                    Enabled = NotDefault
                    TabStop = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =396
                    Top =6406
                    Width =9975
                    Height =375
                    AutoActivate =1
                    TabIndex =8
                    Name ="stbStatusLine"
                    OleData = Begin
                        0x000e0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff01000000a367388e8685d111b16a00c0 ,
                        0xf02836280000000000000000000000000032672e5842c9010700000080010000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000dc000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdfffffffffffffffffffffffffffffffffffffffeffffff ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff01000000a367388e8685d111b16a00c0 ,
                        0xf028362800000000000000000000000080dc7fdc78bacb010500000080010000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000dc000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffefffffffeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000000feffffff030000000400000005000000feffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000002143341208000000bb44000095020000887ee1e60000060080000000 ,
                        0x00001d00ffff130001efcdab00000500e8a01c0006007200ffffffffffffffff ,
                        0x000000000000000003000000a0046600021f0000021f000000000000a0140000 ,
                        0x8c0600008c060000030000000900000034002000730074007500640069006500 ,
                        0x720009000000340020007300740075006400690065007200a0040000a91e0000 ,
                        0xa91e000000000000000000001fdeecbd01000500010000000352e30b918fce11 ,
                        0x9de300aa004bb851010000009001444201000d4d532053616e73205365726966 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OLEClass ="SBarCtrl"
                    Class ="MSComctlLib.SBarCtrl.2"

                    LayoutCachedLeft =396
                    LayoutCachedTop =6406
                    LayoutCachedWidth =10371
                    LayoutCachedHeight =6781
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5045
                    Top =3741
                    Width =681
                    Height =396
                    TabIndex =9
                    Name ="btnRemove"
                    Caption ="<<"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Fjern obligatorisk emne"

                    LayoutCachedLeft =5045
                    LayoutCachedTop =3741
                    LayoutCachedWidth =5726
                    LayoutCachedHeight =4137
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6433
                    Top =6916
                    Width =2040
                    Height =360
                    TabIndex =10
                    Name ="btnObligrapport"
                    Caption ="Vis obligatorske emner"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6433
                    LayoutCachedTop =6916
                    LayoutCachedWidth =8473
                    LayoutCachedHeight =7276
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4364
                    Top =6916
                    Width =2040
                    Height =360
                    TabIndex =11
                    Name ="btnRelasjoner"
                    Caption ="Vis studierelasjoner"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4364
                    LayoutCachedTop =6916
                    LayoutCachedWidth =6404
                    LayoutCachedHeight =7276
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2295
                    Top =6916
                    Width =2040
                    Height =360
                    TabIndex =12
                    Name ="btnValgfri"
                    Caption ="Vis valgfrie emner"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2295
                    LayoutCachedTop =6916
                    LayoutCachedWidth =4335
                    LayoutCachedHeight =7276
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =226
                    Top =6916
                    Width =2040
                    Height =360
                    TabIndex =13
                    Name ="btnExcel"
                    Caption ="Vis oversikt i Excel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =226
                    LayoutCachedTop =6916
                    LayoutCachedWidth =2266
                    LayoutCachedHeight =7276
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


Private Sub btnAdd_Click()
    Dim myDb As DAO.Database
    Dim numStud As Long
    Dim strEmne As String, strTitle As String
    Dim sqlOblig As String
    Dim sqlSource As String
    If Not IsNull(Me.cboStudium.Column(0)) And Me.cboStudium.Column(0) <> "" Then
        numStud = CLng(Me.cboStudium.Column(0))
    Else
        MsgBox "Du må velge et studium", vbExclamation, OIS_Title
        Exit Sub
    End If
    strEmne = Me.lstEmne.Column(0)
    strTitle = Me.lstEmne.Column(1)
    
    Set myDb = CurrentDb
    sqlOblig = "INSERT INTO tblObligemne (StudieID, Emnekode, Emnenavn) VALUES (" & numStud & ", '" & strEmne & "', '" & strTitle & "');"
    myDb.Execute (sqlOblig)
    Call LoadObligEmner(numStud)
End Sub


Private Sub btnExcel_Click()
On Error GoTo Err_btnExcel_Click
    Dim myDb As DAO.Database
    Dim numStud As Long
    Dim rsStudium As DAO.Recordset
    Dim rsEmne As DAO.Recordset
    Dim rsOblig As DAO.Recordset
    Dim ExApp As Object
    Dim ExBook As Excel.Workbook
    Dim ExSheet As Excel.Worksheet
    Const txtTitle = "Studier og obligatoriske emner"
    Const txtOblig = "O"
    Dim cmax As Integer, cant As Integer
    Dim rmax As Integer, rant As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim arrStudnavn() As String
    Dim sqlStudium As String, sqlEmne As String, sqlOblig As String
    Dim stDocName As String, Lastemne As String
    Dim arrStudium() As String, arrStudID() As Long
    Dim arrEmne() As String
    
    Set myDb = CurrentDb

' read study programs into array
    sqlStudium = "SELECT * FROM tblStudium WHERE Aktiv = " & True & " ORDER BY Studiekode;"
    Set rsStudium = myDb.OpenRecordset(sqlStudium, dbOpenDynaset)
    rsStudium.MoveLast
    cant = rsStudium.RecordCount
    ReDim arrStudium(cant + 2)
    ReDim arrStudnavn(cant + 2)
    ReDim arrStudID(cant + 2)
    rsStudium.MoveFirst
    j = 1
    Do While Not rsStudium.EOF
        j = j + 1
        arrStudID(j) = rsStudium!StudieID
        arrStudium(j) = rsStudium!Studiekode
        arrStudnavn(j) = rsStudium!StudieFSkode & " - " & rsStudium!Studienavn
        rsStudium.MoveNext
    Loop
    cmax = j
' read courses into array
    sqlEmne = "SELECT * FROM tblEmne WHERE Aktiv = " & True & " ORDER BY Emnekode;"
    Set rsEmne = myDb.OpenRecordset(sqlEmne, dbOpenDynaset)
    rsEmne.MoveLast
    rant = rsEmne.RecordCount
    ReDim arrEmne(rant + 2)
    rsEmne.MoveFirst
    i = 1
    Lastemne = ""
    Do While Not rsEmne.EOF
        If Trim(rsEmne!Emnekode) <> Lastemne Then
            i = i + 1
            arrEmne(i) = Trim(rsEmne!Emnekode) & " " & Trim(rsEmne!Emnenavn)
            Lastemne = Trim(rsEmne!Emnekode)
        End If
        rsEmne.MoveNext
    Loop
    rmax = i
    
' create Excel worksheet and fill in
    Set ExApp = CreateObject("Excel.Application")
    Set ExBook = ExApp.Workbooks.Add
    Set ExSheet = ExBook.Sheets(1)
    For j = 2 To cmax
        ExSheet.Cells(1, j) = arrStudium(j)
        ExSheet.Cells(1, j).Orientation = 90
        ExSheet.Cells(1, j).HorizontalAlignment = 3 'Center
        ExSheet.Cells(1, j).AddComment (arrStudnavn(j))
    Next j
    
    For i = 2 To rmax
        ExSheet.Cells(i, 1) = arrEmne(i)
    Next i

    ExSheet.Rows(1).RowHeight = 40
    ExSheet.Cells(1, 1) = txtTitle

    ExSheet.Columns(1).ColumnWidth = 50
    For j = 2 To cmax
        ExSheet.Columns(j).ColumnWidth = 2.1
    Next j
'read obligemner and fill into worksheet
    sqlOblig = "SELECT * FROM tblObligemne ORDER BY StudieID;"
    Set rsOblig = myDb.OpenRecordset(sqlOblig, dbOpenDynaset)
    Do While Not rsOblig.EOF
        For k = 2 To cmax
            If rsOblig!StudieID = arrStudID(k) Then j = k
        Next k
        
        For k = 2 To rmax
            If Left(rsOblig!Emnekode, 6) = Left(arrEmne(k), 6) Then i = k
        Next k
        ExSheet.Cells(i, j) = txtOblig
        ExSheet.Cells(i, j).HorizontalAlignment = 3 'Center
        rsOblig.MoveNext
    Loop
    ExApp.Visible = True
    ExApp.UserControl = True

Exit_btnExcel_Click:
    Exit Sub

Err_btnExcel_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnExcel_Click

End Sub

Private Sub btnObligrapport_Click()
On Error GoTo Err_btnObligrapport_Click
    Dim myDb As DAO.Database
    Dim numStud As Long
    Dim rsReport As DAO.Recordset
    Dim rsAlle As DAO.Recordset
    Dim sqlOblig As String, sqlDelete As String
    Dim stDocName As String
    numStud = CLng(Me.cboStudium.Column(0))
    GL_SELECTION = Me.cboStudium.Column(1)
    
    Set myDb = CurrentDb
    Set rsReport = myDb.OpenRecordset("tmpObligEmne", dbOpenDynaset)
    
    Do While Not rsReport.EOF
        rsReport.Delete
        rsReport.MoveNext
    Loop
    
    Set rsAlle = myDb.OpenRecordset("qryAlleEmner", dbOpenDynaset)
    sqlOblig = "SELECT * FROM tblObligemne WHERE StudieID = " & numStud & ";"
    Set rsOblig = myDb.OpenRecordset(sqlOblig, dbOpenDynaset)
    Do While Not rsOblig.EOF
        rsAlle.MoveFirst
        Do While Not rsAlle.EOF
            If rsOblig!Emnekode = rsAlle!Emnekode Then
                rsReport.AddNew
                rsReport!Emnekode = rsAlle!Emnekode
                rsReport!Emnenavn = rsAlle!Emnenavn
                rsReport!Studiepoeng = rsAlle!Studiepoeng
                rsReport!Semester = rsAlle!Semester
                rsReport!studyYear = rsAlle!studyYear
                rsReport!Comment = rsAlle!Comment
                rsReport!Aktiv = rsAlle!Aktiv
                rsReport.Update
                Exit Do
            Else
                rsAlle.MoveNext
            End If
        Loop
        rsOblig.MoveNext
    Loop
    rsOblig.Close
    rsAlle.Close
    rsReport.Close
    stDocName = "rptObligemne"
    DoCmd.OpenReport stDocName, acViewPreview 'WhereCondition:=stCriteria

Exit_btnObligrapport_Click:
    Exit Sub

Err_btnObligrapport_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnObligrapport_Click
End Sub

Private Sub btnRelasjoner_Click()
    Dim myDb As DAO.Database
    Dim rsStudium As DAO.Recordset
    Dim rsOblig As DAO.Recordset
    Dim rsReport As DAO.Recordset
    Dim numStud As Long
    Dim strEmne As String, strNavn As String
    Dim sqlOblig As String
    Dim sqlStudium As String
    If Not IsNull(Me.lstEmne.Column(0)) And Me.lstEmne.Column(0) <> "" Then
        strEmne = Me.lstEmne.Column(0)
        strNavn = Me.lstEmne.Column(1)
    Else
        MsgBox "Du må velge et emne", vbExclamation, OIS_Title
        Exit Sub
    End If

    Set myDb = CurrentDb
    Set rsReport = myDb.OpenRecordset("tmpStudium", dbOpenDynaset)
    
    Do While Not rsReport.EOF
        rsReport.Delete
        rsReport.MoveNext
    Loop
    
    Set rsStudium = myDb.OpenRecordset("tblStudium", dbOpenDynaset)
    sqlOblig = "SELECT * FROM tblObligemne WHERE Emnekode = '" & strEmne & "' ORDER BY StudieID;"
    Set rsOblig = myDb.OpenRecordset(sqlOblig, dbOpenDynaset)
    If rsOblig.RecordCount = 0 Then
        MsgBox "Emnet " & strEmne & " er ikke obligatorisk på noe studium", vbInformation, OIS_Title
        Exit Sub
    End If
    
    Do While Not rsOblig.EOF
        rsStudium.MoveFirst
        Do While Not rsStudium.EOF
            If (rsOblig!StudieID = rsStudium!StudieID) And (rsStudium!Aktiv = True) Then
                rsReport.AddNew
                rsReport!StudieID = rsStudium!StudieID
                rsReport!Studienavn = rsStudium!Studienavn
                rsReport!Oblig = True
                rsReport.Update
                Exit Do
            Else
                rsStudium.MoveNext
            End If
        Loop
        rsOblig.MoveNext
    Loop
    rsOblig.Close
    rsStudium.Close
    rsReport.Close
    stDocName = "rptRelasjon"
    GL_SELECTION = strEmne & " " & strNavn
    DoCmd.OpenReport stDocName, acViewPreview 'WhereCondition:=stCriteria

End Sub

Private Sub btnRemove_Click()
    Dim myDb As DAO.Database
    Dim numStud As Long
    Dim strEmne As String
    Dim sqlDelete As String
    Dim sqlSource As String
    If Not IsNull(Me.lstObligemne) And Me.lstObligemne <> "" Then
        numStud = CLng(Me.lstObligemne.Column(0))
        strEmne = Me.lstObligemne.Column(1)
    Else
        MsgBox "Du må velge et emne", vbExclamation, OIS_Title
        Exit Sub
    End If
    
    Set myDb = CurrentDb
    sqlDelete = "DELETE * FROM tblObligemne WHERE StudieID = " & numStud & " AND Emnekode = '" & strEmne & "';"
    myDb.Execute (sqlDelete)
    Call LoadObligEmner(numStud)

End Sub

Private Sub btnStudiehandbok_Click()
    Dim numStud As Long
    numStud = CLng(Me.cboStudium.Column(0))
    Call showStudyProgram(numStud)
End Sub

Private Sub btnValgfri_Click()
    Dim myDb As DAO.Database
    Dim rsEmne As DAO.Recordset
    Dim rsOblig As DAO.Recordset
    Dim rsReport As DAO.Recordset
    Dim rsParameter As DAO.Recordset
    Dim numStud As Long
    Dim strStudyYear As String
    Dim strEmne As String, strNavn As String
    Dim str
    Dim sqlOblig As String
    Dim sqlStudium As String
    Dim blnFound As Boolean

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
            rsReport!Semester = rsEmne!FirstOfSemester
            rsReport!Aktiv = rsEmne!Aktiv
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

End Sub

Private Sub cboEmnekode_Click()
    Dim ValgtFag As String
    ValgtFag = Me.cboEmnekode.Column(0)
    Call LoadListbox(ValgtFag)
    Me.frOblig = 0
End Sub


Private Sub cboStudium_Click()
    Dim numStud As Long
    Dim strEmne As String, sqlEmne As String
    numStud = CLng(Me.cboStudium.Column(0))
    Call LoadObligEmner(numStud)
End Sub

Private Sub Form_Load()
'
   'GP_SysMgr = True
    Call LoadListbox("")
    Call LoadComboBox
    Me.cboEmnekode.RowSource = "SELECT Kode from tblEmnekode ORDER BY Kode"
    Me.frOblig = 1
    'Me.stbStatusLine.Panels(2) = ""

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


Public Sub LoadListbox(strEmne As String)
    Dim sqlStr As String
    If strEmne = "" Then
        'sqlStr = "SELECT * FROM qryEmne " & _
                "WHERE Aktiv = True ORDER BY Emnekode;"
        sqlStr = "SELECT * FROM qryEmne ORDER BY Emnekode;"

    Else
        'sqlStr = "SELECT * FROM qryEmne " & _
                "WHERE Aktiv = True AND Left(Emnekode,3) = '" & strEmne & "' ORDER BY Emnekode;"
        sqlStr = "SELECT * FROM qryEmne " & _
                "WHERE Left(Emnekode,3) = '" & strEmne & "' ORDER BY Emnekode;"
    End If
    
    Me.lstEmne.RowSource = sqlStr
    Me.stbStatusLine.Panels(1) = Me.lstEmne.ListCount & " emner"
End Sub



Public Sub LoadComboBox()
    Dim sqlSource As String
    sqlSource = "SELECT StudieID, Studienavn FROM tblStudium WHERE Aktiv = True ORDER BY Studienavn;"
    Me.cboStudium.RowSource = sqlSource
    Me.stbStatusLine.Panels(2) = Me.cboStudium.ListCount & " studier"
End Sub


Private Sub frOblig_Click()
    Call LoadListbox("")
    Me.cboEmnekode = ""
End Sub

Public Sub LoadObligEmner(lngStud As Long)
    sqlEmne = "SELECT StudieID, Emnekode, Emnenavn FROM tblObligemne WHERE StudieID = " & lngStud & " ORDER BY Emnekode;"
    Me.lstObligemne.RowSource = sqlEmne
    Me.stbStatusLine.Panels(3) = Me.lstObligemne.ListCount & " obl. emner"
End Sub
