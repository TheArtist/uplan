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
    Width =9977
    DatasheetFontHeight =10
    ItemSuffix =26
    Left =570
    Top =45
    Right =12285
    Bottom =7755
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xdac745677ecae340
    End
    RecordSource ="qryAktiveEmner"
    Caption ="Emnerapport"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    DisplayOnSharePointSite =0
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
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
            BorderColor =8388608
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
        Begin Line
            Width =1701
            BorderColor =8388608
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
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
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
        Begin CheckBox
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
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
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
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
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
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial"
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
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
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
        Begin Subform
            OldBorderStyle =0
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
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
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
        Begin BreakLevel
            ControlSource ="Emnekode"
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
                    Left =60
                    Top =60
                    Width =5235
                    Height =510
                    FontSize =20
                    FontWeight =400
                    Name ="Label12"
                    Caption ="Emneoversikt - aktive emner"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =5295
                    LayoutCachedHeight =570
                End
            End
        End
        Begin PageHeader
            Height =737
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =56
                    Top =283
                    Width =3291
                    Height =288
                    FontSize =10
                    Name ="Emnenavn_Label"
                    Caption ="Emne"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =3741
                    Top =283
                    Width =606
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
                    Left =4362
                    Top =283
                    Width =510
                    Height =288
                    FontSize =10
                    Name ="Semester_Label"
                    Caption ="Sem"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4362
                    LayoutCachedTop =283
                    LayoutCachedWidth =4872
                    LayoutCachedHeight =571
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =4979
                    Top =283
                    Width =1434
                    Height =288
                    FontSize =10
                    Name ="Sted_Label"
                    Caption ="Sted"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4979
                    LayoutCachedTop =283
                    LayoutCachedWidth =6413
                    LayoutCachedHeight =571
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =6528
                    Top =283
                    Width =2429
                    Height =288
                    FontSize =10
                    Name ="Comment_Label"
                    Caption ="Merknad"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6528
                    LayoutCachedTop =283
                    LayoutCachedWidth =8957
                    LayoutCachedHeight =571
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =56
                    Top =623
                    Width =9476
                    Name ="Line15"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2666
                    Top =1
                    Width =1308
                    Height =287
                    FontSize =11
                    ForeColor =10040115
                    Name ="Text19"
                    ControlSource ="studyYear"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =0
                            TextFontFamily =34
                            Left =60
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
                    TextAlign =3
                    TextFontFamily =34
                    Left =4197
                    Width =5482
                    Height =288
                    FontSize =12
                    FontWeight =400
                    Name ="lblUtvalg"
                    FontName ="Arial"
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
                    Left =57
                    Top =57
                    Width =729
                    ColumnWidth =1152
                    Name ="Emnekode"
                    ControlSource ="Emnekode"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =793
                    Top =56
                    Width =2823
                    ColumnWidth =6216
                    TabIndex =1
                    Name ="Emnenavn"
                    ControlSource ="Emnenavn"

                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3741
                    Top =56
                    Width =486
                    ColumnWidth =960
                    TabIndex =2
                    Name ="Studiepoeng"
                    ControlSource ="Studiepoeng"
                    StatusBarText ="Totalt antall studiepoeng som dette emnet gir"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4422
                    Top =56
                    Width =387
                    ColumnWidth =1020
                    TabIndex =3
                    Name ="Semester"
                    ControlSource ="Semester"
                    StatusBarText ="\"H\", \"V\" eller \"H+V\" (eventuelt også \"S\")"

                    LayoutCachedLeft =4422
                    LayoutCachedTop =56
                    LayoutCachedWidth =4809
                    LayoutCachedHeight =296
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4975
                    Top =56
                    Width =1425
                    ColumnWidth =1416
                    TabIndex =4
                    Name ="Sted"
                    ControlSource ="Sted"
                    StatusBarText ="Sted hvor kurset gjennomføres"

                    LayoutCachedLeft =4975
                    LayoutCachedTop =56
                    LayoutCachedWidth =6400
                    LayoutCachedHeight =296
                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =6528
                    Top =56
                    Width =2999
                    ColumnWidth =2244
                    TabIndex =5
                    Name ="Comment"
                    ControlSource ="Comment"

                    LayoutCachedLeft =6528
                    LayoutCachedTop =56
                    LayoutCachedWidth =9527
                    LayoutCachedHeight =296
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
                    Left =5952
                    Top =226
                    Width =3439
                    Height =276
                    FontSize =9
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text14"
                    ControlSource ="=\"Side \" & [Page] & \" of \" & [Pages]"

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
            Height =1077
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3174
                    Top =283
                    Width =1083
                    Height =275
                    FontSize =10
                    Name ="Text17"
                    ControlSource ="=Count([Emnekode])"

                    LayoutCachedLeft =3174
                    LayoutCachedTop =283
                    LayoutCachedWidth =4257
                    LayoutCachedHeight =558
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =0
                            TextFontFamily =34
                            Left =53
                            Top =283
                            Width =2280
                            Height =288
                            FontSize =10
                            FontWeight =400
                            Name ="Label18"
                            Caption ="Totalt antall emner: "
                            FontName ="Arial"
                            LayoutCachedLeft =53
                            LayoutCachedTop =283
                            LayoutCachedWidth =2333
                            LayoutCachedHeight =571
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2828
                    Top =691
                    Width =1428
                    Height =275
                    FontSize =10
                    TabIndex =1
                    Name ="Text21"
                    ControlSource ="=Sum([Studiepoeng])"

                    LayoutCachedLeft =2828
                    LayoutCachedTop =691
                    LayoutCachedWidth =4256
                    LayoutCachedHeight =966
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =0
                            TextFontFamily =34
                            Left =63
                            Top =696
                            Width =2301
                            Height =288
                            FontSize =10
                            FontWeight =400
                            Name ="Label22"
                            Caption ="Totalt antall studiepoeng: "
                            FontName ="Arial"
                            LayoutCachedLeft =63
                            LayoutCachedTop =696
                            LayoutCachedWidth =2364
                            LayoutCachedHeight =984
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
