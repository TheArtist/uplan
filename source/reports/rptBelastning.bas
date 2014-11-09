Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8957
    DatasheetFontHeight =10
    ItemSuffix =35
    Left =60
    Top =90
    Right =9900
    Bottom =6000
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x087c303cbe33e440
    End
    RecordSource ="qryBelastning01"
    Caption ="Undervisningsrapport"
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
            ShowDatePicker =0
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="Navn"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =566
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontFamily =34
                    Left =60
                    Top =60
                    Width =8508
                    Height =504
                    FontSize =20
                    Name ="Label12"
                    Caption ="Undervisningsplan pr faglærer"
                    FontName ="Arial"
                End
            End
        End
        Begin PageHeader
            Height =1036
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =1699
                    Top =340
                    Width =3975
                    Height =288
                    FontSize =10
                    Name ="Emnenavn_Label"
                    Caption ="Emne"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1699
                    LayoutCachedTop =340
                    LayoutCachedWidth =5674
                    LayoutCachedHeight =628
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =6519
                    Top =340
                    Width =537
                    Height =288
                    FontSize =10
                    Name ="Semester_Label"
                    Caption ="Sem"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =7113
                    Top =340
                    Width =1167
                    Height =288
                    FontSize =10
                    Name ="Sted_Label"
                    Caption ="Sted"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7113
                    LayoutCachedTop =340
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =628
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =5725
                    Top =340
                    Width =630
                    Height =288
                    FontSize =10
                    Name ="Studiepoeng_Label"
                    Caption ="Stp"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =113
                    Top =340
                    Width =1535
                    Height =288
                    FontSize =10
                    Name ="Navn_Label"
                    Caption ="Faglærer"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =113
                    LayoutCachedTop =340
                    LayoutCachedWidth =1648
                    LayoutCachedHeight =628
                End
                Begin Line
                    BorderWidth =2
                    Left =56
                    Top =623
                    Width =8888
                    Name ="Line15"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =4535
                    Width =4260
                    Height =288
                    FontSize =10
                    Name ="lblUtvalg"
                    Caption ="xxx"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4535
                    LayoutCachedWidth =8795
                    LayoutCachedHeight =288
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =8333
                    Top =340
                    Width =522
                    Height =288
                    FontSize =10
                    Name ="Label30"
                    Caption ="Ok"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8333
                    LayoutCachedTop =340
                    LayoutCachedWidth =8855
                    LayoutCachedHeight =628
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =113
                    Width =4372
                    Height =288
                    FontSize =12
                    FontWeight =400
                    Name ="lblHeading"
                    FontName ="Arial"
                    LayoutCachedLeft =113
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =288
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =283
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =56
                    Width =2549
                    ColumnWidth =2970
                    Name ="Navn"
                    ControlSource ="Navn"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =383
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =1700
                    Top =57
                    Width =783
                    ColumnWidth =1152
                    Name ="Emnekode"
                    ControlSource ="Emnekode"

                    LayoutCachedLeft =1700
                    LayoutCachedTop =57
                    LayoutCachedWidth =2483
                    LayoutCachedHeight =297
                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =2554
                    Top =57
                    Width =3057
                    ColumnWidth =6216
                    TabIndex =1
                    Name ="Emnenavn"
                    ControlSource ="Emnenavn"

                    LayoutCachedLeft =2554
                    LayoutCachedTop =57
                    LayoutCachedWidth =5611
                    LayoutCachedHeight =297
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6576
                    Top =56
                    Width =417
                    ColumnWidth =1020
                    TabIndex =2
                    Name ="Semester"
                    ControlSource ="Semester"
                    StatusBarText ="\"H\", \"V\" eller \"H+V\" (eventuelt også \"S\")"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7167
                    Top =57
                    Width =1344
                    ColumnWidth =1416
                    TabIndex =3
                    Name ="Sted"
                    ControlSource ="Sted"
                    StatusBarText ="Sted hvor kurset gjennomføres"

                    LayoutCachedLeft =7167
                    LayoutCachedTop =57
                    LayoutCachedWidth =8511
                    LayoutCachedHeight =297
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5725
                    Top =56
                    Width =666
                    TabIndex =4
                    Name ="Studiepoeng"
                    ControlSource ="tblLarerEmne.Studiepoeng"
                    StatusBarText ="Antall studiepoeng av emnet som vedkommnede lærer skal ha"

                End
                Begin CheckBox
                    Left =8674
                    Top =70
                    Width =226
                    Height =227
                    TabIndex =5
                    Name ="Check31"
                    ControlSource ="Ferdig"

                    LayoutCachedLeft =8674
                    LayoutCachedTop =70
                    LayoutCachedWidth =8900
                    LayoutCachedHeight =297
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =640
            Name ="GroupFooter1"
            Begin
                Begin Line
                    BorderWidth =1
                    Left =56
                    Top =339
                    Width =8888
                    Name ="Line18"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5662
                    Top =56
                    Width =681
                    Height =230
                    Name ="Text22"
                    ControlSource ="=Sum([tblLarerEmne.Studiepoeng])"

                    LayoutCachedLeft =5662
                    LayoutCachedTop =56
                    LayoutCachedWidth =6343
                    LayoutCachedHeight =286
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =0
                            TextFontFamily =34
                            Left =2551
                            Top =60
                            Width =1926
                            Height =228
                            FontSize =8
                            FontWeight =400
                            Name ="Label23"
                            Caption ="Sum studiepoeng:"
                            FontName ="Arial"
                            LayoutCachedLeft =2551
                            LayoutCachedTop =60
                            LayoutCachedWidth =4477
                            LayoutCachedHeight =288
                        End
                    End
                End
                Begin Line
                    Left =2544
                    Top =56
                    Width =6401
                    Name ="Line24"
                    LayoutCachedLeft =2544
                    LayoutCachedTop =56
                    LayoutCachedWidth =8945
                    LayoutCachedHeight =56
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8100
                    Top =56
                    Width =756
                    Height =230
                    TabIndex =1
                    Name ="Text33"
                    ControlSource ="=Sum([Stp])"

                    LayoutCachedLeft =8100
                    LayoutCachedTop =56
                    LayoutCachedWidth =8856
                    LayoutCachedHeight =286
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
                    Left =57
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
                    Left =4483
                    Top =228
                    Width =4252
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
                    Width =8888
                    BorderColor =12632256
                    Name ="Line16"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
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
    Me.lblHeading.Caption = GL_FACULTY & ", studieåret " & GL_YEAR
End Sub
