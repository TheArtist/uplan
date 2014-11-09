Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9921
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =270
    Top =210
    Right =11430
    Bottom =5745
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x134323cefc6ae340
    End
    RecordSource ="tmpEmne"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Line
            Width =1701
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =680
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =1
                    TextFontFamily =34
                    Left =-4
                    Width =5550
                    Height =510
                    FontSize =20
                    ForeColor =8388608
                    Name ="Label12"
                    Caption ="Emnerapport - valgfrie emner "
                End
            End
        End
        Begin PageHeader
            Height =1303
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2606
                    Top =1
                    Width =1308
                    Height =287
                    FontSize =11
                    ForeColor =10040115
                    Name ="Text19"
                    ControlSource ="studyYear"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Width =2544
                            Height =288
                            FontSize =11
                            ForeColor =8388608
                            Name ="Label20"
                            Caption ="Avdeling ØIS, studieåret"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =4
                    TextAlign =1
                    TextFontFamily =34
                    Top =283
                    Width =7162
                    Height =288
                    FontSize =12
                    ForeColor =8388608
                    Name ="lblUtvalg"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =793
                    Width =4761
                    Height =288
                    FontSize =10
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Emnenavn_Label"
                    Caption ="Emne"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =1190
                    Width =9296
                    BorderColor =8388608
                    Name ="Line15"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =8277
                    Top =793
                    Width =1026
                    Height =288
                    FontSize =10
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label7"
                    Caption ="Semester"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =7200
                    Top =793
                    Width =1026
                    Height =288
                    FontSize =10
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label9"
                    Caption ="Aktiv"
                    Tag ="DetachedLabel"
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
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =1083
                    FontSize =10
                    Name ="Studienavn"
                    ControlSource ="Emnekode"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1065
                    Width =6753
                    FontSize =10
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="Emnenavn"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8277
                    Width =1038
                    FontSize =10
                    TabIndex =2
                    Name ="Text8"
                    ControlSource ="Semester"
                End
                Begin CheckBox
                    Left =7937
                    Top =56
                    Width =396
                    Height =212
                    TabIndex =3
                    Name ="Check10"
                    ControlSource ="Aktiv"
                End
            End
        End
        Begin PageFooter
            Height =623
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =170
                    Width =4387
                    Height =276
                    FontSize =9
                    ForeColor =8388608
                    Name ="Text13"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5215
                    Top =170
                    Width =4084
                    Height =276
                    FontSize =9
                    TabIndex =1
                    RightMargin =57
                    ForeColor =8388608
                    Name ="Text14"
                    ControlSource ="=\"Side \" & [Page] & \" av \" & [Pages]"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =53
                    Width =9296
                    BorderColor =12632256
                    Name ="Line16"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =453
            Name ="ReportFooter"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Width =2901
                    Height =288
                    FontSize =10
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label5"
                    Caption ="Antall valgfrie emner:"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3288
                    Top =56
                    Width =1083
                    FontSize =10
                    Name ="Text6"
                    ControlSource ="=Count([Emnekode])"
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
