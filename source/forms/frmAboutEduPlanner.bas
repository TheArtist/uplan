Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridX =20
    GridY =24
    Width =6387
    ItemSuffix =39
    Left =4020
    Top =3510
    Right =10755
    Bottom =7275
    HelpContextId =26
    RecSrcDt = Begin
        0x7a9683ca721be340
    End
    RecordSource ="tblParameter"
    Caption ="About Kursplanlegger"
    HelpFile ="MCAHLP.hlp"
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    Begin
        Begin Rectangle
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
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
            TextFontFamily =2
            BorderLineStyle =0
        End
        Begin TextBox
            BorderLineStyle =0
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
        Begin Section
            SpecialEffect =1
            Height =3614
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =1672
                    Top =1251
                    Width =4380
                    FontSize =9
                    TabIndex =1
                    BackColor =-2147483633
                    Name ="RegName"
                    ControlSource ="LicenseHolder"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    Left =1672
                    Top =1464
                    Width =4380
                    FontSize =9
                    BackColor =-2147483633
                    Name ="RegCompanyName"
                    ControlSource ="LicenseAddress"
                    FontName ="Tahoma"

                End
                Begin Label
                    OverlapFlags =85
                    Left =1587
                    Top =897
                    Width =2400
                    Height =240
                    FontSize =9
                    BackColor =-2147483633
                    Name ="Text8"
                    Caption ="This product is licensed to:"
                    FontName ="Tahoma"
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =0
                    OverlapFlags =255
                    Left =1615
                    Top =1228
                    Width =4437
                    Height =806
                    BackColor =12632256
                    BorderColor =8421504
                    Name ="Box11"
                End
                Begin Label
                    OverlapFlags =93
                    Left =1590
                    Top =90
                    Width =2445
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483633
                    Name ="Text5"
                    Caption ="U-Plan"
                End
                Begin Label
                    OverlapFlags =223
                    Left =1587
                    Top =354
                    Width =4500
                    Height =240
                    FontSize =9
                    BackColor =-2147483633
                    Name ="Text7"
                    Caption ="Copyright © 2006-2014 Molde University College"
                    FontName ="Tahoma"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =85
                    Left =147
                    Top =2100
                    Width =5898
                    BorderColor =8421504
                    Name ="Line13"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =85
                    Left =147
                    Top =2119
                    Width =5898
                    BorderColor =16777215
                    Name ="Line14"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =79
                    TextFontFamily =34
                    Left =4430
                    Top =2338
                    Width =1622
                    Height =347
                    FontSize =9
                    FontWeight =400
                    TabIndex =3
                    Name ="btnOk"
                    Caption ="&Ok"
                    StatusBarText ="Close this form."
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =170
                    Top =2338
                    Width =4170
                    Height =1176
                    BackColor =-2147483633
                    Name ="Text17"
                    Caption ="Warning: This computer program is protected by copyright law and international t"
                        "reaties. Unauthorized reproduction or distribution of this program, or any porti"
                        "on of it, may result in severe civil and criminal penalties, and will be prosecu"
                        "ted to the maximum extent possible under law. "
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    Left =1672
                    Top =1700
                    Width =1065
                    Height =240
                    FontSize =9
                    BackColor =12632256
                    Name ="Text18"
                    Caption ="Serial Number:"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    Left =2891
                    Top =1700
                    Width =2880
                    FontSize =9
                    TabIndex =2
                    BackColor =-2147483633
                    Name ="SerialNumber"
                    ControlSource ="LicenseNo"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    Left =4030
                    Top =90
                    Width =1728
                    Height =273
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BackColor =-2147483633
                    Name ="Version"
                    ControlSource ="LicenseVersion"

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
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub LastControl_Enter()
On Error Resume Next
DoCmd.GoToControl "btnOk"
End Sub

Private Sub btnOk_Click()
On Error Resume Next
    DoCmd.Close 'A_FORM, Me.RegName
End Sub
