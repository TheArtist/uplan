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
    ItemSuffix =37
    Left =60
    Top =90
    Right =10590
    Bottom =8190
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x78fb72a99b7ce340
    End
    RecordSource ="qryUndervisning"
    Caption ="Undervisningsrapport"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="Emnekode"
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
                    Caption ="Undervisningsplan pr emne"
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
                    Left =113
                    Top =340
                    Width =3690
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
                    Left =3860
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
                    Left =4454
                    Top =340
                    Width =1482
                    Height =288
                    FontSize =10
                    Name ="Sted_Label"
                    Caption ="Sted"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =6293
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
                    Left =7148
                    Top =340
                    Width =1355
                    Height =285
                    FontSize =10
                    Name ="Navn_Label"
                    Caption ="Faglærer"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
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
                    Left =4325
                    Width =4462
                    Height =288
                    FontSize =12
                    FontWeight =400
                    Name ="lblUtvalg"
                    FontName ="Arial"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =8385
                    Top =340
                    Width =525
                    Height =285
                    FontSize =10
                    Name ="Label35"
                    Caption ="Ok"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =56
                    Width =4372
                    Height =288
                    FontSize =12
                    FontWeight =400
                    Name ="lblHeading"
                    FontName ="Arial"
                    LayoutCachedLeft =56
                    LayoutCachedWidth =4428
                    LayoutCachedHeight =288
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =283
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =56
                    Width =783
                    ColumnWidth =1152
                    FontWeight =700
                    Name ="Emnekode"
                    ControlSource ="Emnekode"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =896
                    Width =3018
                    ColumnWidth =6216
                    FontWeight =700
                    TabIndex =1
                    Name ="Emnenavn"
                    ControlSource ="Emnenavn"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6346
                    Width =513
                    Height =230
                    FontWeight =700
                    TabIndex =2
                    Name ="txtStp"
                    ControlSource ="tblEmne.Studiepoeng"

                End
                Begin CheckBox
                    Left =8730
                    Width =226
                    Height =230
                    TabIndex =3
                    Name ="chkFerdig"
                    ControlSource ="Ferdig"

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
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3972
                    Top =57
                    Width =417
                    ColumnWidth =1020
                    Name ="Semester"
                    ControlSource ="Semester"
                    StatusBarText ="\"H\", \"V\" eller \"H+V\" (eventuelt også \"S\")"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4446
                    Top =57
                    Width =1674
                    ColumnWidth =1416
                    TabIndex =1
                    Name ="Sted"
                    ControlSource ="Sted"
                    StatusBarText ="Sted hvor kurset gjennomføres"

                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6292
                    Top =56
                    Width =666
                    TabIndex =2
                    Name ="Studiepoeng"
                    ControlSource ="tblLarerEmne.Studiepoeng"
                    StatusBarText ="Antall studiepoeng av emnet som vedkommnede lærer skal ha"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7140
                    Top =57
                    Width =1805
                    ColumnWidth =2970
                    TabIndex =3
                    Name ="Navn"
                    ControlSource ="Navn"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =113
            Name ="GroupFooter1"
            Begin
                Begin Line
                    BorderWidth =1
                    Left =56
                    Top =56
                    Width =8888
                    Name ="Line18"
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
                    Left =4478
                    Top =226
                    Width =4207
                    Height =276
                    FontSize =9
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text14"
                    ControlSource ="=\"Side \" & [Page] & \" av \" & [Pages]"

                    LayoutCachedLeft =4478
                    LayoutCachedTop =226
                    LayoutCachedWidth =8685
                    LayoutCachedHeight =502
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
