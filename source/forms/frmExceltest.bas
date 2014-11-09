Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    DatasheetFontHeight =10
    ItemSuffix =1
    Left =270
    Top =225
    Right =6225
    Bottom =1935
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x152d7abc598ce340
    End
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin Section
            Height =1701
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1190
                    Top =566
                    Width =1814
                    Height =347
                    Name ="btnExcel"
                    Caption ="Åpne Excel"
                    OnClick ="[Event Procedure]"

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

Private Sub btnExcel_Click()
    Dim ExApp As Object
    Dim ExBook As Excel.Workbook
    Dim ExSheet As Excel.Worksheet

    Set ExApp = CreateObject("Excel.Application")
    'ExApp.Workbooks.Open (AFile)
    Set ExBook = ExApp.Workbooks.Add
    Set ExSheet = ExBook.Sheets(1)
    ExSheet.Name = "NNN"
    ExSheet.Cells(1, 1) = "Emne matrise"
    ExSheet.Columns(1).ColumnWidth = 10
    ExApp.Visible = True
    ExApp.UserControl = True

End Sub
