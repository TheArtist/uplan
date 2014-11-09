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
    Width =9694
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =270
    Top =210
    Right =12195
    Bottom =6390
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xebaa05cc7d6ae340
    End
    RecordSource ="qryRelasjon"
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
                    Left =56
                    Width =5490
                    Height =510
                    FontSize =20
                    ForeColor =8388608
                    Name ="Label12"
                    Caption ="Studierelasjoner for emne "
                End
            End
        End
        Begin PageHeader
            Height =963
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
                    TextAlign =3
                    TextFontFamily =34
                    Left =4242
                    Width =4837
                    Height =288
                    FontSize =12
                    ForeColor =8388608
                    Name ="lblUtvalg"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =453
                    Width =4761
                    Height =288
                    FontSize =10
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Emnenavn_Label"
                    Caption ="Studium"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =7535
                    Top =453
                    Width =1560
                    Height =285
                    FontSize =10
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label23"
                    Caption ="Obligatorisk"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =850
                    Width =9131
                    BorderColor =8388608
                    Name ="Line15"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =396
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =7368
                    FontSize =10
                    Name ="Studienavn"
                    ControlSource ="Studienavn"
                End
                Begin CheckBox
                    Left =8107
                    Top =56
                    Width =397
                    Height =230
                    TabIndex =1
                    Name ="chkOblig"
                    ControlSource ="Oblig"
                End
            End
        End
        Begin PageFooter
            Height =453
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =2
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
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5669
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
                    Top =2
                    Width =9476
                    BorderColor =12632256
                    Name ="Line16"
                End
                Begin Line
                    BorderWidth =3
                    Width =9641
                    BorderColor =12632256
                    Name ="Line1"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =226
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
End Sub
