Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =55
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4875
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =4872
    Top =96
    Right =10008
    Bottom =3552
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x50c0d897fb5ee240
    End
    RecordSource ="UserTab"
    Caption ="User administration"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnActivate ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            SpecialEffect =3
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
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
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            BackStyle =0
            Width =5103
            Height =3402
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =2664
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    Left =2130
                    Top =566
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    Name ="UserName"
                    ControlSource ="UserName"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =510
                            Top =566
                            Width =1560
                            Height =255
                            Name ="UserName_Label"
                            Caption ="User name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =2130
                    Top =908
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="UserPwd"
                    ControlSource ="UserPwd"
                    InputMask ="Password"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =510
                            Top =908
                            Width =1560
                            Height =255
                            Name ="UserPwd_Label"
                            Caption ="Password"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =2130
                    Top =1250
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="UserType"
                    ControlSource ="UserType"
                    StatusBarText ="\"Adm\" or \"User\""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =510
                            Top =1250
                            Width =1560
                            Height =255
                            Name ="UserType_Label"
                            Caption ="Permission"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4180
                    Top =1984
                    Width =441
                    Height =405
                    TabIndex =3
                    Name ="btnNext"
                    Caption ="Command7"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x28000000110000000f0000000100040000000000b40000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888880000000 ,
                        0x8888800888888888800000008888800088888888840000008888800008888888 ,
                        0x8004322688888000008888888040e26288888000000888888440366688888000 ,
                        0x000088888111a3338888800000088888899deaee8888800000888888888caaaa ,
                        0x8888800008888888811133228888800088888888811909108888800888888888 ,
                        0x8ff865f088888888888888888f8765f088888888888888888f0865f000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Next Record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3685
                    Top =1984
                    Width =441
                    Height =405
                    TabIndex =4
                    Name ="btnPrevious"
                    Caption ="Command8"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x28000000110000000f0000000100040000000000b40000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888880000000 ,
                        0x8888888800888888800000008888888000888888840000008888880000888888 ,
                        0x8004322688888000008888888040e26288880000008888888440366688800000 ,
                        0x008888888111a3338888000000888888899deaee8888800000888888888caaaa ,
                        0x8888880000888888811133228888888000888888811909108888888800888888 ,
                        0x8ff865f088888888888888888f8765f088888888888888888f0865f000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Previous Record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1977
                    Top =1984
                    Width =780
                    Height =405
                    TabIndex =5
                    Name ="btnDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1130
                    Top =1984
                    Width =780
                    Height =405
                    TabIndex =6
                    Name ="btnNew"
                    Caption ="New"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =283
                    Top =1984
                    Width =780
                    Height =405
                    TabIndex =7
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                End
                Begin OptionGroup
                    OverlapFlags =247
                    Left =283
                    Top =283
                    Width =4365
                    Height =1587
                    TabIndex =8
                    Name ="Frame12"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =408
                            Top =168
                            Width =2328
                            Height =228
                            Name ="Label13"
                            Caption ="User, password and permission"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =15
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


Private Sub btnNext_Click()
On Error GoTo Err_btnNext_Click


    DoCmd.GoToRecord , , acNext

Exit_btnNext_Click:
    Exit Sub

Err_btnNext_Click:
    MsgBox NextMsg, vbExclamation, RCM_Title
    Resume Exit_btnNext_Click
    
End Sub
Private Sub btnPrevious_Click()
On Error GoTo Err_btnPrevious_Click


    DoCmd.GoToRecord , , acPrevious

Exit_btnPrevious_Click:
    Exit Sub

Err_btnPrevious_Click:
    MsgBox PrevMsg, vbExclamation, RCM_Title
    Resume Exit_btnPrevious_Click
    
End Sub
Private Sub btnDelete_Click()
On Error GoTo Err_btnDelete_Click
    
    Dim msg As String, response As Integer
    msg = "Do you want to remove user " & Me!UserName

    response = MsgBox(msg, vbExclamation + vbYesNo, RCM_Title)
    If response = vbNo Then
        Exit Sub
    End If
    DoCmd.SetWarnings False
    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
    Me.Requery
    DoCmd.SetWarnings True

Exit_btnDelete_Click:
    Exit Sub

Err_btnDelete_Click:
    MsgBox Err.Description, , RCM_Title
    Resume Exit_btnDelete_Click
    
End Sub
Private Sub btnNew_Click()
On Error GoTo Err_btnNew_Click

    Dim stDocName As String
    Dim stLinkCriteria As String '

    stDocName = "frmNewUser"
    DoCmd.OpenForm stDocName, , , stLinkCriteria


Exit_btnNew_Click:
    Exit Sub

Err_btnNew_Click:
    MsgBox Err.Description, , RCM_Title
    Resume Exit_btnNew_Click
    
End Sub
Private Sub btnClose_Click()
On Error GoTo Err_btnClose_Click


    DoCmd.Close

Exit_btnClose_Click:
    Exit Sub

Err_btnClose_Click:
    MsgBox Err.Description, , RCM_Title
    Resume Exit_btnClose_Click
    
End Sub


Private Sub Form_Activate()
    Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim msg1 As String, msg2 As String
    
    If Not GP_SysMgr Then
        msg1 = "You are not authorized to administer users."
        msg2 = "Please contact system administrator."
        MsgBox msg1 & vbNewLine & msg2, vbExclamation, RCM_Title
        DoCmd.Close
    End If

End Sub
