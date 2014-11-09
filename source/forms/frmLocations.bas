Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =119
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12245
    DatasheetFontHeight =10
    ItemSuffix =33
    Left =1560
    Top =255
    Right =14460
    Bottom =7350
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa8680edd5b3ae340
    End
    RecordSource ="tblParameter"
    Caption ="Navn og lokasjoner"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
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
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            Width =5103
            Height =3402
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =6122
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3864
                    Top =850
                    Width =7638
                    Height =255
                    ColumnWidth =2310
                    FontSize =9
                    Name ="txtHolder"
                    ControlSource ="Holder"
                    StatusBarText ="Institusjonens navn"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3864
                    LayoutCachedTop =850
                    LayoutCachedWidth =11502
                    LayoutCachedHeight =1105
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1010
                            Top =850
                            Width =1935
                            Height =255
                            FontSize =9
                            Name ="LicensePlant_Label"
                            Caption ="Institusjon:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1010
                            LayoutCachedTop =850
                            LayoutCachedWidth =2945
                            LayoutCachedHeight =1105
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3864
                    Top =3929
                    Width =7653
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =1
                    Name ="txtAvdPath"
                    ControlSource ="AvdPath"
                    StatusBarText ="The path for storing activity descriptions"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3864
                    LayoutCachedTop =3929
                    LayoutCachedWidth =11517
                    LayoutCachedHeight =4184
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1010
                            Top =3929
                            Width =2391
                            Height =255
                            FontSize =9
                            Name ="LicenseDocPath_Label"
                            Caption ="Sti til arbeidsplan:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1010
                            LayoutCachedTop =3929
                            LayoutCachedWidth =3401
                            LayoutCachedHeight =4184
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3852
                    Top =4302
                    Width =7656
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =2
                    Name ="txtArbPlan"
                    ControlSource ="ArbPlan"
                    StatusBarText ="Mal for arbeidsplan"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3852
                    LayoutCachedTop =4302
                    LayoutCachedWidth =11508
                    LayoutCachedHeight =4557
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1010
                            Top =4298
                            Width =2634
                            Height =264
                            FontSize =9
                            Name ="LicenseSparePath_Label"
                            Caption ="Mal for arbeidsplan, Norsk:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1010
                            LayoutCachedTop =4298
                            LayoutCachedWidth =3644
                            LayoutCachedHeight =4562
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =255
                    Left =396
                    Top =2101
                    Width =11338
                    Height =3339
                    TabIndex =3
                    Name ="Frame6"

                    LayoutCachedLeft =396
                    LayoutCachedTop =2101
                    LayoutCachedWidth =11734
                    LayoutCachedHeight =5440
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =684
                            Top =1980
                            Width =2034
                            Height =240
                            FontSize =9
                            Name ="Label7"
                            Caption ="Lokasjoner og studieår"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =255
                    Left =396
                    Top =564
                    Width =11289
                    Height =1138
                    TabIndex =4
                    Name ="Frame8"

                    LayoutCachedLeft =396
                    LayoutCachedTop =564
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =1702
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =684
                            Top =456
                            Width =1638
                            Height =240
                            FontSize =9
                            Name ="Label9"
                            Caption ="Navn og adresse"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10487
                    Top =5612
                    Width =1005
                    Height =345
                    FontSize =9
                    TabIndex =5
                    Name ="btnClose"
                    Caption ="Ok"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =10487
                    LayoutCachedTop =5612
                    LayoutCachedWidth =11492
                    LayoutCachedHeight =5957
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9240
                    Top =5612
                    Width =1095
                    Height =345
                    FontSize =9
                    TabIndex =6
                    Name ="btnCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =5612
                    LayoutCachedWidth =10335
                    LayoutCachedHeight =5957
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3864
                    Top =1225
                    Width =7638
                    Height =255
                    FontSize =9
                    TabIndex =7
                    Name ="txtAddress"
                    ControlSource ="Address"
                    StatusBarText ="Adresse"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3864
                    LayoutCachedTop =1225
                    LayoutCachedWidth =11502
                    LayoutCachedHeight =1480
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1025
                            Top =1225
                            Width =1920
                            Height =255
                            FontSize =9
                            Name ="Label20"
                            Caption ="Adresse:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1025
                            LayoutCachedTop =1225
                            LayoutCachedWidth =2945
                            LayoutCachedHeight =1480
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3864
                    Top =2437
                    Width =7653
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="txtKursinfo"
                    ControlSource ="Kursinfo"
                    StatusBarText ="URL for Kursplan"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3864
                    LayoutCachedTop =2437
                    LayoutCachedWidth =11517
                    LayoutCachedHeight =2692
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1010
                            Top =2437
                            Width =2391
                            Height =255
                            FontSize =9
                            Name ="Label22"
                            Caption ="Kursplan URL:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1010
                            LayoutCachedTop =2437
                            LayoutCachedWidth =3401
                            LayoutCachedHeight =2692
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3864
                    Top =2810
                    Width =7653
                    Height =255
                    FontSize =9
                    TabIndex =9
                    Name ="txtStudiehandbok"
                    ControlSource ="Studiehandbok"
                    StatusBarText ="URL for studiehåndbok"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3864
                    LayoutCachedTop =2810
                    LayoutCachedWidth =11517
                    LayoutCachedHeight =3065
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1010
                            Top =2810
                            Width =2391
                            Height =255
                            FontSize =9
                            Name ="Label24"
                            Caption ="Studiehåndbok URL:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1010
                            LayoutCachedTop =2810
                            LayoutCachedWidth =3401
                            LayoutCachedHeight =3065
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3864
                    Top =3556
                    Width =7653
                    Height =255
                    FontSize =9
                    TabIndex =10
                    Name ="txttopURL"
                    ControlSource ="topURL"
                    StatusBarText ="URL for studiehåndbok"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3864
                    LayoutCachedTop =3556
                    LayoutCachedWidth =11517
                    LayoutCachedHeight =3811
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1010
                            Top =3565
                            Width =2706
                            Height =252
                            FontSize =9
                            Name ="Label26"
                            Caption ="Topp page studiehåndbok:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1010
                            LayoutCachedTop =3565
                            LayoutCachedWidth =3716
                            LayoutCachedHeight =3817
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3843
                    Top =5048
                    Width =7656
                    Height =255
                    FontSize =9
                    TabIndex =11
                    Name ="txtstudyYear"
                    ControlSource ="studyYear"
                    StatusBarText ="Studieår"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3843
                    LayoutCachedTop =5048
                    LayoutCachedWidth =11499
                    LayoutCachedHeight =5303
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1010
                            Top =5044
                            Width =2619
                            Height =264
                            FontSize =9
                            Name ="Label28"
                            Caption ="Studieår:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1010
                            LayoutCachedTop =5044
                            LayoutCachedWidth =3629
                            LayoutCachedHeight =5308
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3864
                    Top =3183
                    Width =7653
                    Height =255
                    FontSize =9
                    TabIndex =12
                    Name ="txtStudiehandbok_eng"
                    ControlSource ="Studiehandbok_eng"
                    StatusBarText ="URL for studiehåndbok"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3864
                    LayoutCachedTop =3183
                    LayoutCachedWidth =11517
                    LayoutCachedHeight =3438
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1010
                            Top =3181
                            Width =2670
                            Height =255
                            FontSize =9
                            Name ="Label30"
                            Caption ="Studiehåndbok, eng. URL:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1010
                            LayoutCachedTop =3181
                            LayoutCachedWidth =3680
                            LayoutCachedHeight =3436
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3852
                    Top =4675
                    Width =7656
                    Height =255
                    FontSize =9
                    TabIndex =13
                    Name ="txtArbPlanE"
                    ControlSource ="ArbPlanE"
                    StatusBarText ="Mal for arbeidsplan"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3852
                    LayoutCachedTop =4675
                    LayoutCachedWidth =11508
                    LayoutCachedHeight =4930
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1010
                            Top =4671
                            Width =2634
                            Height =264
                            FontSize =9
                            Name ="Label32"
                            Caption ="Mal for arbeidsplan, Engelsk:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1010
                            LayoutCachedTop =4671
                            LayoutCachedWidth =3644
                            LayoutCachedHeight =4935
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnClose_Click()
On Error GoTo Err_btnClose_Click


    DoCmd.Close

Exit_btnClose_Click:
    Exit Sub

Err_btnClose_Click:
    MsgBox Err.Description
    Resume Exit_btnClose_Click
    
End Sub
Private Sub btnCancel_Click()
On Error GoTo Err_btnCancel_Click


    DoCmd.Close , , acSaveNo

Exit_btnCancel_Click:
    Exit Sub

Err_btnCancel_Click:
    MsgBox Err.Description
    Resume Exit_btnCancel_Click
    
End Sub
