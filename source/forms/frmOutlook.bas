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
    ItemSuffix =5
    Left =270
    Top =225
    Right =6900
    Bottom =3525
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2059e7521c5ce340
    End
    DatasheetFontName ="Arial"
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
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =2494
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1585
                    Top =1190
                    Width =2949
                    Height =341
                    FontSize =9
                    Name ="txtRecipient"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =285
                            Top =1185
                            Width =1200
                            Height =240
                            FontSize =9
                            Name ="Label1"
                            Caption ="Epostadresse:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2953
                    Top =1927
                    Width =1526
                    Height =397
                    FontSize =9
                    TabIndex =1
                    Name ="btnOutlook"
                    Caption ="Hent Adresse"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1583
                    Top =566
                    Width =2949
                    Height =341
                    FontSize =9
                    TabIndex =2
                    Name ="txtNavn"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =283
                            Top =567
                            Width =1050
                            Height =240
                            FontSize =9
                            Name ="Label4"
                            Caption ="Navn:"
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


Private Sub btnOutlook_Click()
Dim oOutlookApp As Outlook.Application
Set oOutlookApp = New Outlook.Application
    Dim oRecipient As Recipient
    Dim oNameSpace As Namespace
    Dim oContact As ContactItem
    Dim strName As String
    Set oNameSpace = GetNamespace("MAPI")
    Me.txtRecipient = ""
    If Me.txtNavn <> "" Then
        strName = Trim(Me.txtNavn)
        Set oRecipient = oNameSpace.CreateRecipient(strName)
        oRecipient.Resolve
        If oRecipient.Resolved Then
            Me.txtRecipient = oRecipient
        Else
            Me.txtRecipient = "Email address not found"
        End If
    End If
End Sub
