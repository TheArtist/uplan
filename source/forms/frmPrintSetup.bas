Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =55
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5115
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =2835
    Top =60
    Right =8295
    Bottom =1965
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7c26acabdc53e240
    End
    Caption ="Printer setup"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Section
            Height =1230
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CustomControl
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =60
                    Top =75
                    Width =480
                    Height =480
                    AutoActivate =1
                    Name ="dlgPrinter"
                    OnExit ="[Event Procedure]"
                    OLEClass ="CommonDialog"
                    Class ="MSCOMDLG.COMMONDIALOG"

                End
                Begin Label
                    OverlapFlags =85
                    Left =705
                    Top =225
                    Width =2040
                    Height =600
                    Name ="Label1"
                    Caption ="Printer setup is completed Press Ok to continue"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3435
                    Top =390
                    Width =1095
                    Height =405
                    TabIndex =1
                    Name ="btnOk"
                    Caption ="Ok"
                    OnClick ="[Event Procedure]"

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

Private Sub dlgPrinter_Exit(Cancel As Integer)
    Forms!PrintSetupFrm.Close
End Sub

Private Sub Form_Load()
On Error Resume Next
    'dlgPrinter.Flags = &H40  'Printer Setup dialog only
    'dlgPrinter.ShowPrinter
    DoCmd.RunCommand (acCmdPrintSelection)
    Printers.UserControl
End Sub

Private Sub btnOk_Click()
On Error GoTo Err_btnOk_Click

    DoCmd.Close

Exit_btnOk_Click:
    Exit Sub

Err_btnOk_Click:
    MsgBox Err.Description, , RCM_Title
    Resume Exit_btnOk_Click
End Sub
