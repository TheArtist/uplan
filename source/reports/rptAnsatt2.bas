Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =127
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9694
    DatasheetFontHeight =10
    ItemSuffix =28
    Left =3630
    Top =90
    Right =15165
    Bottom =10500
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x371335253c3be340
    End
    RecordSource ="qryAnsatt"
    Caption ="rptEmne"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            Width =850
            Height =850
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            BorderLineStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            ControlSource ="Navn"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =623
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontFamily =34
                    Top =60
                    Width =4104
                    Height =504
                    FontSize =20
                    FontWeight =400
                    Name ="Label12"
                    Caption ="Faglæreroversikt"
                    FontName ="Arial"
                End
            End
        End
        Begin PageHeader
            Height =850
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Top =397
                    Width =3291
                    Height =288
                    FontSize =10
                    Name ="Navn_Label"
                    Caption ="Navn"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =3456
                    Top =401
                    Width =1704
                    Height =288
                    FontSize =10
                    Name ="Stilling_Label"
                    Caption ="Stilling"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =5555
                    Top =396
                    Width =858
                    Height =288
                    FontSize =10
                    Name ="Andel_Label"
                    Caption ="St.andel"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =6636
                    Top =396
                    Width =2888
                    Height =288
                    FontSize =10
                    Name ="Merknad_Label"
                    Caption ="Merknad"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =736
                    Width =9527
                    Name ="Line15"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2666
                    Width =1308
                    Height =288
                    FontSize =11
                    ForeColor =10040115
                    Name ="Text19"
                    ControlSource ="studyYear"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =0
                            TextFontFamily =34
                            Width =2544
                            Height =288
                            FontWeight =400
                            Name ="Label20"
                            Caption ="Avdeling ØIS, studieåret"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =0
                    TextFontFamily =34
                    Left =5555
                    Width =3975
                    Height =288
                    FontWeight =400
                    Name ="lblUtvalg"
                    Caption ="%"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =340
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Top =56
                    Width =3291
                    ColumnWidth =6216
                    Name ="Navn"
                    ControlSource ="Navn"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3452
                    Top =56
                    Width =1878
                    ColumnWidth =960
                    TabIndex =1
                    Name ="Stilling"
                    ControlSource ="StNavn"
                    StatusBarText ="Totalt antall studiepoeng som dette emnet gir"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5612
                    Top =56
                    Width =534
                    ColumnWidth =1416
                    TabIndex =2
                    Name ="Andel"
                    ControlSource ="Andel"
                    StatusBarText ="Sted hvor kurset gjennomføres"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =6636
                    Top =56
                    Width =2888
                    ColumnWidth =2244
                    TabIndex =3
                    Name ="Merknad"
                    ControlSource ="Merk"

                End
            End
        End
        Begin PageFooter
            Height =504
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Top =228
                    Width =4387
                    Height =276
                    FontSize =9
                    ForeColor =8388608
                    Name ="Text13"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5839
                    Top =226
                    Width =3694
                    Height =276
                    FontSize =9
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text14"
                    ControlSource ="=\"Side \" & [Page] & \" av \" & [Pages]"

                End
                Begin Line
                    BorderWidth =3
                    Left =57
                    Top =228
                    Width =9476
                    BorderColor =12632256
                    Name ="Line16"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =566
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3469
                    Top =113
                    Width =1459
                    Height =283
                    FontSize =10
                    Name ="Text23"
                    ControlSource ="=Count([Navn])"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =0
                            TextFontFamily =34
                            Top =113
                            Width =2430
                            Height =285
                            FontSize =10
                            FontWeight =400
                            Name ="Label24"
                            Caption ="Antall faglærere i utvalg:"
                            FontName ="Arial"
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
Private Sub Report_Open(Cancel As Integer)
    Me.lblUtvalg.Caption = GL_SELECTION
End Sub
