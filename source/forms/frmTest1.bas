Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6994
    DatasheetFontHeight =11
    ItemSuffix =1
    Top =600
    Right =8775
    Bottom =8610
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xc13c78d5d6efe340
    End
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    DatasheetBackColor12 =-2147483643
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
        End
        Begin Section
            Height =5952
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =793
                    Top =1133
                    Width =1725
                    Height =405
                    ForeColor =4138256
                    Name ="Rediger"
                    Caption ="Rediger faglærer"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =793
                    LayoutCachedTop =1133
                    LayoutCachedWidth =2518
                    LayoutCachedHeight =1538
                    BorderColor =12835293
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

Private Sub Rediger_Click()
On Error GoTo Err_Rediger_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frmForeleser"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Rediger_Click:
    Exit Sub

Err_Rediger_Click:
    MsgBox Err.Description
    Resume Exit_Rediger_Click
    
End Sub
