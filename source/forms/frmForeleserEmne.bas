Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15760
    DatasheetFontHeight =10
    ItemSuffix =108
    Left =-3795
    Right =11970
    Bottom =9000
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x00529b40d4c5e240
    End
    Caption ="Forelesere og emner"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
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
            Height =8730
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin OptionGroup
                    OverlapFlags =93
                    Left =226
                    Top =505
                    Width =7525
                    Height =7378
                    ColumnOrder =9
                    Name ="frPersoner"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =340
                            Top =397
                            Width =2550
                            Height =228
                            FontSize =9
                            Name ="Label17"
                            Caption ="Faglærer og emnetilknytning"
                            FontName ="Tahoma"
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =453
                            Top =907
                            Width =1021
                            Height =170
                            OptionValue =1
                            Name ="chkAllePersoner"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =680
                                    Top =850
                                    Width =684
                                    Height =228
                                    FontSize =9
                                    Name ="Label90"
                                    Caption ="Vis alle"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10015
                    Top =7427
                    Width =1710
                    Height =330
                    FontSize =9
                    TabIndex =1
                    Name ="btnDelete"
                    Caption ="Fjern "
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13719
                    Top =7427
                    Width =1710
                    Height =330
                    FontSize =9
                    TabIndex =2
                    Name ="btnOk"
                    Caption ="Lukk"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8163
                    Top =7427
                    Width =1710
                    Height =330
                    FontSize =9
                    TabIndex =3
                    Name ="btnUpdate"
                    Caption ="Legg til "
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =8050
                    Top =512
                    Width =7538
                    Height =4867
                    ColumnOrder =6
                    TabIndex =4
                    Name ="frEmner"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =8163
                            Top =397
                            Width =1305
                            Height =228
                            FontSize =9
                            Name ="Label56"
                            Caption ="Aktive emner"
                            FontName ="Tahoma"
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =8268
                            Top =907
                            Width =907
                            Height =170
                            OptionValue =1
                            Name ="chkAlleEmner"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =8498
                                    Top =877
                                    Width =684
                                    Height =228
                                    FontSize =9
                                    Name ="Label92"
                                    Caption ="Vis alle"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =13325
                    Top =850
                    Width =2105
                    Height =284
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =5
                    Name ="cboEmnegruppe"
                    RowSourceType ="Table/Query"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =12472
                            Top =850
                            Width =789
                            Height =228
                            FontSize =9
                            Name ="Label68"
                            Caption ="Vis bare:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CustomControl
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =340
                    Top =1252
                    Width =7260
                    Height =3960
                    AutoActivate =1
                    TabIndex =6
                    Name ="lvwPersonale"
                    OleData = Begin
                        0x00120000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000000000000000000000100000 ,
                        0x0600000001000000feffffff0000000001000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000004bf0d1bd8b85d111b16a00c0 ,
                        0xf0283628000000000000000000000000e08c02be7aaac9010700000000020000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000200000058010000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdfffffffffffffffffffffffffffffffffffffffeffffff ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000004bf0d1bd8b85d111b16a00c0 ,
                        0xf0283628000000000000000000000000005c94837aaac9010500000000020000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000200000058010000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffefffffffeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000000feffffff0300000004000000050000000600000007000000 ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000004bf0d1bd8b85d111b16a00c0f0283628214334120800000006320000 ,
                        0x491b00004e087deb010006001c00000000000000000000000086010006320000 ,
                        0x01efcdab00000500c009d00b07002a000800008005000080d89fd80b00000000 ,
                        0x00000000000000001fdeecbd01000500519fd80b0352e30b918fce119de300aa ,
                        0x004bb851010000009001905f0100065461686f6d61050020000000000000003a ,
                        0x11000005000000000000000000000000000000050000004e61766e0020000100 ,
                        0x00000000060b000009000000000000000000000000000000090000005374696c ,
                        0x6c696e6700200002000200000034060000060000000000000000000000000000 ,
                        0x0006000000416e64656c002000030000000000b70b0000080000000000000000 ,
                        0x00000000000000080000004d65726b6e61640020000400000000000000000009 ,
                        0x00000000000000000000000000000009000000506572736f6e49440000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000001000000feffffff0300000004000000050000000600000007000000 ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000004bf0d1bd8b85d111b16a00c0f0283628214334120800000006320000 ,
                        0x491b00004e087deb010006001c00000000000000000000000086010006320000 ,
                        0x01efcdab00000500c009d00b07002a000800008005000080d89fd80b00000000 ,
                        0x00000000000000001fdeecbd01000500519fd80b0352e30b918fce119de300aa ,
                        0x004bb851010000009001905f0100065461686f6d61050020000000000000003a ,
                        0x11000005000000000000000000000000000000050000004e61766e0020000100 ,
                        0x00000000060b000009000000000000000000000000000000090000005374696c ,
                        0x6c696e6700200002000200000034060000060000000000000000000000000000 ,
                        0x0006000000416e64656c002000030000000000170d0000080000000000000000 ,
                        0x00000000000000080000004d65726b6e61640020000400000000000000000009 ,
                        0x00000000000000000000000000000009000000506572736f6e49440000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OLEClass ="ListViewCtrl"
                    Class ="MSComctlLib.ListViewCtrl.2"

                End
                Begin CustomControl
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =340
                    Top =5891
                    Width =7260
                    Height =1365
                    AutoActivate =1
                    TabIndex =7
                    Name ="lvwKurs"
                    OleData = Begin
                        0x00140000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000000000000000000000100000 ,
                        0x0600000001000000feffffff0000000001000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000004bf0d1bd8b85d111b16a00c0 ,
                        0xf0283628000000000000000000000000305a1cf99d33ce010700000040020000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000aa010000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdfffffffffffffffffffffffffffffffffffffffeffffff ,
                        0x08000000feffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000004bf0d1bd8b85d111b16a00c0 ,
                        0xf028362800000000000000000000000070533f058b33ce010500000040020000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000020000008e010000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffeffffff09000000ffffffff ,
                        0xfffffffffffffffffeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000000feffffff0300000004000000050000000600000007000000 ,
                        0x08000000feffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0xe022860902002000000000000000ac1400000500000000000000000000000000 ,
                        0x00000500214334120800000006320000680900004e087deb010006001c000000 ,
                        0x0000000000000000008601000632000001efcdab0000050020f6420c07002a00 ,
                        0x0800008005000080e023890a0000000000000000000000001fdeecbd01000500 ,
                        0x5923890a0352e30b918fce119de300aa004bb851010000009001905f01000654 ,
                        0x61686f6d61070020000000000000004508000005000000000000000000000000 ,
                        0x000000050000004b6f64650020000100000000000c1600000700000000000000 ,
                        0x00000000000000000700000054697474656c0020000200020000003406000004 ,
                        0x0000000000000000000000000000000400000053747000200003000200000034 ,
                        0x060000090000000000000000000000000000000900000053656d657374657200 ,
                        0x2000040000000000940700000500000000000000000000000000000005000000 ,
                        0x5374656400200005000000000000000000030000000000000000000000000000 ,
                        0x00030000004944002000060000000000000000000a0000000000000000000000 ,
                        0x0000000001000000feffffff0300000004000000050000000600000007000000 ,
                        0x08000000feffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0xe022860902002000000000000000ac1400000500000000000000000000000000 ,
                        0x00000500214334120800000006320000680900004e087deb010006001c000000 ,
                        0x0000000000000000008601000632000001efcdab00000500f017860c07002a00 ,
                        0x0800008005000080f847c20a0e00000069006d0067006c006900730074000e00 ,
                        0x000069006d0067006c00690073007400000000001fdeecbd010005007147c20a ,
                        0x0352e30b918fce119de300aa004bb851010000009001905f0100065461686f6d ,
                        0x6107002000000000000000450800000500000000000000000000000000000005 ,
                        0x0000004b6f64650020000100000000000c160000070000000000000000000000 ,
                        0x000000000700000054697474656c002000020002000000340600000400000000 ,
                        0x0000000000000000000000040000005374700020000300020000003406000009 ,
                        0x0000000000000000000000000000000900000053656d65737465720020000400 ,
                        0x0000000094070000050000000000000000000000000000000500000053746564 ,
                        0x0020000500000000000000000003000000000000000000000000000000030000 ,
                        0x004944002000060000000000000000000a000000000000000000000000000000 ,
                        0x0a0000004d65726b6e6164657200540000004500640077006100720064006900 ,
                        0x61006e0020005300630072006900700074002000490054004300000045006e00 ,
                        0x670072006100760065007200730020004d005400000045007200610073002000 ,
                        0x440065006d00690020004900540043000000450072006100730020004c006900 ,
                        0x670068007400200049005400430000004500750072006f007300740069006c00 ,
                        0x65000000460065006c006900780020005400690074006c0069006e0067000000 ,
                        0x4600720061006e006b006c0069006e00200047006f0074006800690063002000 ,
                        0x42006f006f006b0000004600720061006e006b006c0069006e00200047006f00 ,
                        0x74006800690063002000440065006d00690000004600720061006e006b006c00 ,
                        0x69006e00200047006f0074006800690063002000480065006100760079000000 ,
                        0x4600720061006e006b006c0069006e00200047006f0074006800690063002000 ,
                        0x440065006d006900200043006f006e00640000004600720065006e0063006800 ,
                        0x200053006300720069007000740020004d0054000000430065006e0074007500 ,
                        0x72007900200047006f00740068006900630000004b0072006900730074006500 ,
                        0x6e00200049005400430000004c00750063006900640061002000530061006e00 ,
                        0x7300000000000000
                    End
                    OLEClass ="ListViewCtrl"
                    Class ="MSComctlLib.ListViewCtrl.2"

                End
                Begin CustomControl
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =8164
                    Top =1249
                    Width =7266
                    Height =3954
                    AutoActivate =1
                    TabIndex =8
                    Name ="lvwAlleEmner"
                    OleData = Begin
                        0x00140000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000004bf0d1bd8b85d111b16a00c0 ,
                        0xf0283628000000000000000000000000700126dd79aac9010700000080020000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000d5010000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdfffffffffffffffffffffffffffffffffffffffeffffff ,
                        0x09000000fffffffffeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000004bf0d1bd8b85d111b16a00c0 ,
                        0xf028362800000000000000000000000080c51fac87aac9010500000080020000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000d5010000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffeffffff08000000ffffffff ,
                        0xfffffffffeffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000000feffffff0300000004000000050000000600000007000000 ,
                        0x0800000009000000feffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0xe022860902002000000000000000ac1400000500000000000000000000000000 ,
                        0x00000500214334120800000006320000491b00004e087deb010006001c000000 ,
                        0x0000000000000000008601000632000001efcdab000005005846f00a07002a00 ,
                        0x0800008005000080a8607a0e0e00000069006d0067006c006900730074000e00 ,
                        0x000069006d0067006c00690073007400000000001fdeecbd0100050021607a0e ,
                        0x0352e30b918fce119de300aa004bb851010000009001905f0100065461686f6d ,
                        0x6108002000000000000000450800000500000000000000000000000000000005 ,
                        0x0000004b6f64650020000100000000004b130000070000000000000000000000 ,
                        0x000000000700000054697474656c002000020002000000340600000400000000 ,
                        0x0000000000000000000000040000005374700020000300020000003406000009 ,
                        0x0000000000000000000000000000000900000053656d65737465720020000400 ,
                        0x0000000094070000050000000000000000000000000000000500000053746564 ,
                        0x0020000500000000000000000007000000000000000000000000000000070000 ,
                        0x00456d6e01000000feffffff0300000004000000050000000600000007000000 ,
                        0x0800000009000000feffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0xe022860902002000000000000000ac1400000500000000000000000000000000 ,
                        0x00000500214334120800000006320000491b00004e087deb010006001c000000 ,
                        0x0000000000000000008601000632000001efcdab00000500d842dd0b07002a00 ,
                        0x0800008005000080e090150e0e00000069006d0067006c006900730074000e00 ,
                        0x000069006d0067006c00690073007400000000001fdeecbd010005005990150e ,
                        0x0352e30b918fce119de300aa004bb851010000009001905f0100065461686f6d ,
                        0x6108002000000000000000450800000500000000000000000000000000000005 ,
                        0x0000004b6f64650020000100000000004b130000070000000000000000000000 ,
                        0x000000000700000054697474656c002000020002000000340600000400000000 ,
                        0x0000000000000000000000040000005374700020000300020000003406000009 ,
                        0x0000000000000000000000000000000900000053656d65737465720020000400 ,
                        0x000000008c060000050000000000000000000000000000000500000053746564 ,
                        0x0020000500000000000000000007000000000000000000000000000000070000 ,
                        0x00456d6e654944002000060000000000000000000a0000000000000000000000 ,
                        0x000000000a0000004d65726b6e61646572002000070000000000000000000700 ,
                        0x0000000000000000000000000000070000004665726469670000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000050006900 ,
                        0x6300740075007200650041006c00690067006e006d0065006e00740000000000 ,
                        0x000000000000000000001d000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000054006500780074004200610063006b0067007200 ,
                        0x6f0075006e0064000000000000000000000000000000000000001d0000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000084021b010f007b021623170000000000 ,
                        0x0000000100000000000001000100000000000001000100010000010001000000 ,
                        0x0000000000000000
                    End
                    OLEClass ="ListViewCtrl"
                    Class ="MSComctlLib.ListViewCtrl.2"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =340
                    Top =5552
                    Width =6466
                    Height =290
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =9
                    Name ="txtLarer"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5099
                    Top =850
                    Width =2501
                    Height =284
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =10
                    Name ="cboStilling"
                    RowSourceType ="Table/Query"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4138
                            Top =850
                            Width =792
                            Height =228
                            FontSize =9
                            Name ="Label79"
                            Caption ="Vis bare:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9417
                    Top =6236
                    Width =5998
                    Height =283
                    ColumnOrder =11
                    FontSize =9
                    TabIndex =11
                    Name ="txtEmne"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8163
                            Top =6236
                            Width =576
                            Height =228
                            FontSize =9
                            Name ="Label81"
                            Caption ="Emne:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9417
                    Top =6576
                    Width =859
                    Height =283
                    ColumnOrder =16
                    FontSize =9
                    TabIndex =12
                    Name ="txtBelastning"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8163
                            Top =6582
                            Width =900
                            Height =228
                            FontSize =9
                            Name ="Label83"
                            Caption ="Stp:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12979
                    Top =6576
                    Width =1705
                    Height =283
                    ColumnOrder =14
                    FontSize =9
                    TabIndex =13
                    Name ="txtSted"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12302
                            Top =6576
                            Width =624
                            Height =228
                            FontSize =9
                            Name ="Label85"
                            Caption ="Sted:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11403
                    Top =6576
                    Width =847
                    Height =283
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =14
                    Name ="txtTid"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10488
                            Top =6582
                            Width =864
                            Height =228
                            FontSize =9
                            Name ="Label87"
                            Caption ="Semester:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11867
                    Top =7427
                    Width =1710
                    Height =330
                    FontSize =9
                    TabIndex =15
                    Name ="btnAddForeleser"
                    Caption ="Faglærer(e)"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9417
                    Top =6916
                    Width =5998
                    Height =283
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =16
                    Name ="txtMerknader"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8163
                            Top =6921
                            Width =1086
                            Height =228
                            FontSize =9
                            Name ="Label94"
                            Caption ="Merknader:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CustomControl
                    Enabled = NotDefault
                    TabStop = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =226
                    Top =7993
                    Width =15240
                    Height =375
                    AutoActivate =1
                    TabIndex =17
                    Name ="stbStatusLine"
                    OleData = Begin
                        0x00160000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff01000000a367388e8685d111b16a00c0 ,
                        0xf0283628000000000000000000000000c0f5bb573eaac9010b00000000090000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000200000059080000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffeffffff0c0000000d0000000e0000000f000000 ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff01000000a367388e8685d111b16a00c0 ,
                        0xf028362800000000000000000000000090758e1d79aac9010500000000090000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000200000059080000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffeffffff0600000007000000 ,
                        0x0800000009000000feffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000000feffffff0300000004000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x20000000210000002200000023000000feffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000021433412080000000269000095020000887ee1e600000600c2000000 ,
                        0x00001f00ffff130001efcdab000005000000000006007200ffffffffffffffff ,
                        0x000000000000000004000000a00466002f1a00002f1a000000000000a0140000 ,
                        0x2f1a00002f1a0000030000000900000034002000730074007500640069006500 ,
                        0x720009000000340020007300740075006400690065007200a0046200b82a0000 ,
                        0xb82a000003000000070000003400200065006d006e0065007200070000003400 ,
                        0x200065006d006e0065007200b5040000780e0000780e00000200000005000000 ,
                        0x300030003a003000320001000000040000006c74000036070000000001000300 ,
                        0x202002000000000030010000360000002020080000000000e802000066010000 ,
                        0x2020100000000000e80200004e04000028000000200000004000000001000100 ,
                        0x00000000000100000000000000000000000000000000000000000000ffffff00 ,
                        0x00000000000000003ffffff83ffffff83000001837ffffd837feffd837ffffd8 ,
                        0x37ffffd8377fffd837bfffd837dfffd837efffd837f6ffd837faffd837fcffd8 ,
                        0x35f8035837fe7fd837febfd837feffd837feffd837feffd837feffd837feffd8 ,
                        0x37ffffd837feffd837ffffd8300000183ffffff83ffffff80000000000000000 ,
                        0xffffffffc0000007800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x800000038000000380000003800000038000000380000003c0000007ffffffff ,
                        0x2800000020000000400000000100040000000000800200000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000009999999999999999999999999990000099999999999999 ,
                        0x99999999999990000099000000000000000000000009900000990fffffffffff ,
                        0xffffffffff09900000990ffffffffff0ffffffffff09900000990fffffffffff ,
                        0xffffffffff09900000990fffffffffffffffffffff09900000990fff0fffffff ,
                        0xffffffffff09900000990ffff0ffffffffffffffff09900000990fffff0fffff ,
                        0xffffffffff09900000990ffffff0ffffffffffffff09900000990fffffff0ff0 ,
                        0xffffffffff09900000990ffffffff0f0ffffffffff09900000990fffffffff00 ,
                        0xffffffffff09900000990f0ffffff000000000ff0f09900000990ffffffffff0 ,
                        0x0fffffffff09900000990ffffffffff0f0ffffffff09900000990ffffffffff0 ,
                        0xffffffffff09900000990ffffffffff0ffffffffff09900000990ffffffffff0 ,
                        0xffffffffff09900000990ffffffffff0ffffffffff09900000990ffffffffff0 ,
                        0xffffffffff09900000990fffffffffffffffffffff09900000990ffffffffff0 ,
                        0xffffffffff09900000990fffffffffffffffffffff0990000099000000000000 ,
                        0x0000000000099000009999999999999999999999999990000099999999999999 ,
                        0x9999999999999000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffc000000780000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0xc0000007ffffffff280000002000000040000000010004000000000080020000 ,
                        0x0000000000000000000000000000000000000000000080000080000000808000 ,
                        0x80000000800080008080000080808000c0c0c0000000ff0000ff000000ffff00 ,
                        0xff000000ff00ff00ffff0000ffffff0000007777777777777777777777777770 ,
                        0x0007777777777777777777777777777700000000000000000000000000000777 ,
                        0x0011111111111111111111111111107700999999999999999999999999991077 ,
                        0x0091000000000000000000000009107700910fffffffffffffffffffff091077 ,
                        0x00910ffffffffff0ffffffffff09107700910fffffffffffffffffffff091077 ,
                        0x00910fffffffffffffffffffff09107700910fff0fffffffffffffffff091077 ,
                        0x00910ffff0ffffffffffffffff09107700910fffff0fffffffffffffff091077 ,
                        0x00910ffffff0ffffffffffffff09107700910fffffff0ff0ffffffffff091077 ,
                        0x00910ffffffff0f0ffffffffff09107700910fffffffff00ffffffffff091077 ,
                        0x00910f0ffffff000000000ff0f09107700910ffffffffff00fffffffff091077 ,
                        0x00910ffffffffff0f0ffffffff09107700910ffffffffff0ffffffffff091077 ,
                        0x00910ffffffffff0ffffffffff09107700910ffffffffff0ffffffffff091077 ,
                        0x00910ffffffffff0ffffffffff09107700910ffffffffff0ffffffffff091077 ,
                        0x00910fffffffffffffffffffff09107700910ffffffffff0ffffffffff091077 ,
                        0x00910fffffffffffffffffffff09107700910000000000000000000000091077 ,
                        0x0091111111111111111111111119107000999999999999999999999999991000 ,
                        0x00000000000000000000000000000000f0000001e0000000c000000080000000 ,
                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                        0x800000008000000180000003c00000071fdeecbd01000500010000000352e30b ,
                        0x918fce119de300aa004bb851010000009001905f0100065461686f6d61732053 ,
                        0x6572696680000003c00000071fdeecbd01000500010000000352e30b918fce11 ,
                        0x9de300aa004bb851010000009001b03001000d4d532053616e73205365726966 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff00000000
                    End
                    OLEClass ="SBarCtrl"
                    Class ="MSComctlLib.SBarCtrl.2"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11338
                    Top =113
                    Width =1078
                    Height =227
                    ColumnOrder =15
                    TabIndex =18
                    Name ="txtKursID"

                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =2195
                    Top =7427
                    Width =2148
                    Height =330
                    FontSize =9
                    TabIndex =19
                    Name ="btnArbeidsplan"
                    Caption ="Overfør til arbeidsplan "
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6121
                    Top =7426
                    Width =1482
                    Height =330
                    FontSize =9
                    TabIndex =20
                    Name ="btnNullstill"
                    Caption ="Nullstill"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =8163
                    Top =5836
                    Width =6870
                    Height =343
                    FontSize =9
                    Name ="Label101"
                    Caption ="Legg til , endre eller fjerne dette emnet fra valgt faglærers kursportefølje"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =14853
                    Top =6633
                    Width =277
                    Height =215
                    ColumnOrder =5
                    TabIndex =21
                    Name ="chkFerdig"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =15083
                            Top =6571
                            Width =399
                            Height =288
                            FontSize =9
                            Name ="Label105"
                            Caption ="Ok"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =340
                    Top =7427
                    Width =1707
                    Height =330
                    FontSize =9
                    TabIndex =22
                    Name ="btnSendArbeidsplan"
                    Caption ="Send arbeidsplan  "
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CustomControl
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =8560
                    Top =3004
                    Width =570
                    Height =570
                    AutoActivate =1
                    TabIndex =23
                    Name ="imglist"
                    OleData = Begin
                        0x00520000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff01000000237f242c9185d111b16a00c0 ,
                        0xf028362800000000000000000000000070299bd2429bc8011500000080000000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000160000001b1f0000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffefffffffeffffff17000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x200000002100000022000000230000002400000025000000feffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff01000000237f242c9185d111b16a00c0 ,
                        0xf028362800000000000000000000000040c5c44b439bc8010500000080000000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000060000001b1f0000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffefffffffeffffff07000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000026000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff27000000 ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000000feffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000002143341208000000ed030000ed030000807ee1e60000060028000000 ,
                        0x10000d00c0c0c000ffffffff01efcdab00000500801f23000680ffffffffffff ,
                        0x050000800000000008000000000000000000000008000000010000006c740000 ,
                        0xed0000004749463839610e000e00c4ff000000002929295252527b7b7b8c0000 ,
                        0x946b00ff005a844a4a84734ace9400ff9400ffce00848484a5a58cb5b5b5ffff ,
                        0x8cffff94c6c6c6cec6cecececeeff7eff7f7f7c0c0c000000000000000000000 ,
                        0x000000000000000000000000000000000021f90401000016002c000000000e00 ,
                        0x0e0000056aa0258e88201e84398a48b39887a1920d00bcf16cb58034dd0019ad ,
                        0x47a148802ade845279547eaf5d4d527c509c919b20107054ae14085822300dbe ,
                        0x8f821a10268b06e0c263011047548c6a61a16043dc167062050a09754f2a030e ,
                        0x1211050905001112805c3696972621003b020000006c74000036040000000001 ,
                        0x0002002020100000000000e8020000260000001010100000000000280100000e ,
                        0x0300002800000020000000400000000100040000000000800200000000000000 ,
                        0x0000000000000000000000000000000000800000800000008080008000000080 ,
                        0x0080008080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff ,
                        0x00ff00ffff0000ffffff00000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000777777777 ,
                        0x777777777777777777730007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007ffffffff ,
                        0xfffffffffffffffffff00007888888888888888777777777777000007fb8b8b8 ,
                        0xb8b8b870000000000000000007fb8b8b8b8b87000000000000000000007fffff ,
                        0xffff700000000000000000000007777777770000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000ffffffffffffffffffffffffffffffffc000000380 ,
                        0x0000018000000180000001800000018000000180000001800000018000000180 ,
                        0x0000018000000180000001800000018000000180000001800000018000000180 ,
                        0x000001800000018000000380000007c0007fffe000fffff001fffff803ffffff ,
                        0xffffffffffffffffffffff2800000010000000200000000100040000000000c0 ,
                        0x0000000000000000000000000000000000000000000000000080000080000000 ,
                        0x80800080000000800080008080000080808000c0c0c0000000ff0000ff000000 ,
                        0xffff00ff000000ff00ff00ffff0000ffffff0000000000000000000000000000 ,
                        0x000000000000000000000077777777777777007fb8b8b8b8b8b7007f8b8b8b8b ,
                        0x8b87007fb8b8b8b8b8b7007f8b8b8b8b8b87007fb8b8b8b8b8b7007f8b8b8b8b ,
                        0x8b87007fb8b8b8b8b8b7007ffffffffffff70078b8b8b877777700078b8b8700 ,
                        0x00000000777770000000000000000000000000ffff0000ffff00008001000000 ,
                        0x0100000001000000010000000100000001000000010000000100000001000000 ,
                        0x0100000003000080ff0000c1ff0000ffff0000030000006c7400003604000000 ,
                        0x00010002002020100000000000e8020000260000001010100000000000280100 ,
                        0x000e030000280000002000000040000000010004000000000080020000000000 ,
                        0x0000000000000000000000000000000000000080000080000000808000800000 ,
                        0x00800080008080000080808000c0c0c0000000ff0000ff000000ffff00ff0000 ,
                        0x00ff00ff00ffff0000ffffff0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000733333333333333333333333333300007b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b8b83000078b8b8b8b8b8b8b8b8b8b8b8b8b300078b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b8b800007b8b8b8b8b8b8b8b8b8b8b8b8b88000078b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b8b700007b8b8b8b8b8b8b8b8b8b8b8b8b870007b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b8b030078b8b8b8b8b8b8b8b8b8b8b8b8b803007b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b87030078b8b8b8b8b8b8b8b8b8b8b8b8b703078b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b807307b8b8b8b8b8b8b8b8b8b8b8b8b8b073078b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b70b307b8b8b8b8b8b8b8b8b8b8b8b8b8708307fffff ,
                        0xfffffffffffffffffffff07b300778888888888888888888888888b8300007fb ,
                        0x8b8b8b8b8b8b8b8b8b8b8b8b300007f8b8b8b8b8b8b8b8b8b8b8b8b8300007fb ,
                        0x8b8b8b8b8b8b8b8b8b8b8b8b300007f8b8b8b8b8b8b8bfffffffffff000007fb ,
                        0x8b8b8b8b8b8b8777777777770000007fb8b8b8b8b8b870000000000000000007 ,
                        0xfb8b8b8b8b87000000000000000000007fffffffff7000000000000000000000 ,
                        0x0777777777000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000fffffffffffffffffffffffff0000001e00000 ,
                        0x00e0000000e0000000c0000000c0000000c0000000c000000080000000800000 ,
                        0x0080000000800000000000000000000000000000000000000000000000800000 ,
                        0x00e0000000e0000000e0000000e0000001e0000003f0001ffff8003ffffc007f ,
                        0xfffe00ffffffffffffffffffff28000000100000002000000001000400000000 ,
                        0x00c0000000000000000000000000000000000000000000000000008000008000 ,
                        0x000080800080000000800080008080000080808000c0c0c0000000ff0000ff00 ,
                        0x0000ffff00ff000000ff00ff00ffff0000ffffff000000000000000000000000 ,
                        0x000000000000000000000000000077777777777700007fb8b8b8b8b70007fb8b ,
                        0x8b8b8b807007f8b8b8b8b870707f8b8b8b8b8b07707ffffffffff70870777777 ,
                        0x7777777b7007f8b8b8b8b8b87007fb8b8b8fffff7007f8b8b8f7777770007fff ,
                        0xff7000000000077777000000000000000000000000ffff0000ffff0000e00000 ,
                        0x00c0000000c00000008000000080000000000000000000000000000000800000 ,
                        0x008000000080010000c07f0000e0ff0000ffff0000040000006c740000360400 ,
                        0x000000010002002020100000000000e802000026000000101010000000000028 ,
                        0x0100000e03000028000000200000004000000001000400000000008002000000 ,
                        0x0000000000000000000000000000000000000000008000008000000080800080 ,
                        0x000000800080008080000080808000c0c0c0000000ff0000ff000000ffff00ff ,
                        0x000000ff00ff00ffff0000ffffff000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000880000000000000 ,
                        0x0000000000000009077880000000000000000000000000990777788000000000 ,
                        0x0000000000000999800777788000000000000000000099980990077778800000 ,
                        0x0000000000099980999990077770000000000000009998099999999007700000 ,
                        0x0000000009998099999999999000000000000000999809999999999999900000 ,
                        0x0000000999809999999999999990000000000009980999999999999999000000 ,
                        0x0000000980999999999999999000000000000008099999999999999900000000 ,
                        0x0000000009999999999999900000000000000000000999999999990000000000 ,
                        0x0000000000000999999990000000000000000000000000099999000000000000 ,
                        0x0000000000000000099000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffc7ffffff81ffffff007ffffe001ffffc0007ff ,
                        0xf80001fff00003ffe00003ffc00003ff800001ff000003ff000007ff00000fff ,
                        0x00001fff80003fffe0007ffff800fffffe01ffffff83ffffffe7ffffffffffff ,
                        0xffffffffffffffffffffffffffffff2800000010000000200000000100040000 ,
                        0x000000c000000000000000000000000000000000000000000000000000800000 ,
                        0x8000000080800080000000800080008080000080808000c0c0c0000000ff0000 ,
                        0xff000000ffff00ff000000ff00ff00ffff0000ffffff00000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000907000000 ,
                        0x0000009807880000000009800077800000009809990070000009809999990000 ,
                        0x0008099999990000000099999990000000000999990000000000000990000000 ,
                        0x0000000000000000000000000000000000000000000000ffff0000ffff0000ff ,
                        0xff0000ff8f0000ff030000fe000000fc010000f8010000f0000000f0010000f0 ,
                        0x030000f8070000fe0f0000ff9f0000ffff0000ffff0000050000006c74000036 ,
                        0x0400000000010002002020100000000000e80200002600000010101000000000 ,
                        0x00280100000e0300002800000020000000400000000100040000000000800200 ,
                        0x0000000000000000000000000000000000000000000000800000800000008080 ,
                        0x0080000000800080008080000080808000c0c0c0000000ff0000ff000000ffff ,
                        0x00ff000000ff00ff00ffff0000ffffff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000088000000000 ,
                        0x0000000000000000000907788000000000000000000000000000000778800000 ,
                        0x000000000000700077770880077880000000000000070fff000008fff0077880 ,
                        0x000000000070fffff8808ffffff0077000000000070fffff8808ff77fffff000 ,
                        0x0000000070fffff8808fffff77fffff0000000070fffff8808ff77ffff77ff00 ,
                        0x00000070fffff8808fffff77fffff0000000070fffff8808ff77ffff77ff0000 ,
                        0x000070fffff8808fffff77fffff0000000000fffff8808ff77ffff77ff000000 ,
                        0x0000000ff880000fff77fffff000000000000000000000000fff77ff00000000 ,
                        0x0000000000000000000ffff000000000000000000000000000000f0000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffc7ffffff81ffffff007fffc0001fff8000 ,
                        0x07ff000001fe000003fc000003f8000003f0000007e000000fc000001f800000 ,
                        0x3f0000007fe02000fff87801fffffe03ffffff87ffffffefffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffff280000001000000020000000010004 ,
                        0x0000000000c00000000000000000000000000000000000000000000000000080 ,
                        0x00008000000080800080000000800080008080000080808000c0c0c0000000ff ,
                        0x0000ff000000ffff00ff000000ff00ff00ffff0000ffffff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000088000 ,
                        0x0000000000000880000000fff880f00880000fff880ffff00000fff880ffffff ,
                        0x0008ff880ffffff000008880ffffff000000000000fff0000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000ffff0000ffff00 ,
                        0x00ffff0000ff9f0000ff070000f0010000e0000000c001000080010000000300 ,
                        0x0000070000c40f0000ff1f0000ffff0000ffff0000ffff0000060000006c7400 ,
                        0x00360400000000010002002020100001000400e8020000260000001010100001 ,
                        0x000400280100000e030000280000002000000040000000010004000000000080 ,
                        0x0200000000000000000000100000000000000000000000000080000080000000 ,
                        0x80800080000000800080008080000080808000c0c0c0000000ff0000ff000000 ,
                        0xffff00ff000000ff00ff00ffff0000ffffff0000000077777777777777777777 ,
                        0x77770000000000000000000000000000000700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbb0000000000000bbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbb0000000000000bbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbb0000000000000bbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbb0000000000000bbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb070000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000fc000003f8000003f8000003f8 ,
                        0x000003f8000003f8000003f8000003f8000003f8000003f8000003f8000003f8 ,
                        0x000003f8000003f8000003f8000003f8000003f8000003f8000003f8000007ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffff28000000100000002000000001 ,
                        0x00040000000000c0000000000000000000000000000000000000000000000000 ,
                        0x008000008000000080800080000000800080008080000080808000c0c0c00000 ,
                        0x00ff0000ff000000ffff00ff000000ff00ff00ffff0000ffffff000000000000 ,
                        0x00000000bfbfbfbfbfbf0000fb000000fbfb0000bfbfbfbfbfbf0000fb00000b ,
                        0xfbfb0000bfbfbfbfbfbf0000fb0000000bfb0000bfbfbfbfb55f0000fbfbfbfb ,
                        0xf55b0000bfbfbfbfbfbf00000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000008001000080 ,
                        0x0100008001000080010000800100008001000080010000800100008001000080 ,
                        0x01000080010000ffff0000ffff0000ffff0000ffff0000ffff0000070000006c ,
                        0x740000360400000000010002002020100001000400e802000026000000101010 ,
                        0x0001000400280100000e03000028000000200000004000000001000400000000 ,
                        0x0080020000000000000000000010000000000000000000000000008000008000 ,
                        0x000080800080000000800080008080000080808000c0c0c0000000ff0000ff00 ,
                        0x0000ffff00ff000000ff00ff00ffff0000ffffff000000007777777777777777 ,
                        0x777777770000000000000000000000000000000700000000ffffffffffffffff ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000fff0000000000000 ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000fff0000000000000ffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000fff0000000000000 ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000fff0000000000000ffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000ffffffffffffffffffffff07000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000fc000003f8000003f80000 ,
                        0x03f8000003f8000003f8000003f8000003f8000003f8000003f8000003f80000 ,
                        0x03f8000003f8000003f8000003f8000003f8000003f8000003f8000003f80000 ,
                        0x07ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffff2800000010000000200000 ,
                        0x000100040000000000c000000000000000000000000000000000000000000000 ,
                        0x0000008000008000000080800080000000800080008080000080808000c0c0c0 ,
                        0x000000ff0000ff000000ffff00ff000000ff00ff00ffff0000ffffff00000000 ,
                        0x000000000000bfbfbfbfbfbf0000fb000000fbfb0000bfbfbfbfbfbf0000fb00 ,
                        0x000bfbfb0000bfbfbfbfbfbf0000fb0000000bfb0000bfbfbfbfb55f0000fbfb ,
                        0xfbfbf55b0000bfbfbfbfbfbf0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000800100 ,
                        0x0080010000800100008001000080010000800100008001000080010000800100 ,
                        0x008001000080010000ffff0000ffff0000ffff0000ffff0000ffff0000080000 ,
                        0x006c740000360400000000010002002020100001000400e80200002600000010 ,
                        0x10100001000400280100000e0300002800000020000000400000000100040000 ,
                        0x0000008002000000000000000000001000000000000000000000000000800000 ,
                        0x8000000080800080000000800080008080000080808000c0c0c0000000ff0000 ,
                        0xff000000ffff00ff000000ff00ff00ffff0000ffffff00000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000c000000000000000000000000000000cc0 ,
                        0x0000000000000000000000000000cccc000000000000000000000000000ccccc ,
                        0x0000000001000000feffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000002143341208000000ed030000ed030000807ee1e60000060028000000 ,
                        0x10000d00c0c0c000ffffffff01efcdab00000500e82b1e000680ffffffffffff ,
                        0x050000800000000008000000000000000000000008000000010000006c740000 ,
                        0xed0000004749463839610e000e00c4ff000000002929295252527b7b7b8c0000 ,
                        0x946b00ff005a844a4a84734ace9400ff9400ffce00848484a5a58cb5b5b5ffff ,
                        0x8cffff94c6c6c6cec6cecececeeff7eff7f7f7c0c0c000000000000000000000 ,
                        0x000000000000000000000000000000000021f90401000016002c000000000e00 ,
                        0x0e0000056aa0258e88201e84398a48b39887a1920d00bcf16cb58034dd0019ad ,
                        0x47a148802ade845279547eaf5d4d527c509c919b20107054ae14085822300dbe ,
                        0x8f821a10268b06e0c263011047548c6a61a16043dc167062050a09754f2a030e ,
                        0x1211050905001112805c3696972621003b020000006c74000036040000000001 ,
                        0x0002002020100000000000e8020000260000001010100000000000280100000e ,
                        0x0300002800000020000000400000000100040000000000800200000000000000 ,
                        0x0000000000000000000000000000000000800000800000008080008000000080 ,
                        0x0080008080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff ,
                        0x00ff00ffff0000ffffff00000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000777777777 ,
                        0x777777777777777777730007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007f8b8b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b30007fb8b8b8b8b8b8b8b8b8b8b8b8b830007ffffffff ,
                        0xfffffffffffffffffff00007888888888888888777777777777000007fb8b8b8 ,
                        0xb8b8b870000000000000000007fb8b8b8b8b87000000000000000000007fffff ,
                        0xffff700000000000000000000007777777770000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000ffffffffffffffffffffffffffffffffc000000380 ,
                        0x0000018000000180000001800000018000000180000001800000018000000180 ,
                        0x0000018000000180000001800000018000000180000001800000018000000180 ,
                        0x000001800000018000000380000007c0007fffe000fffff001fffff803ffffff ,
                        0xffffffffffffffffffffff2800000010000000200000000100040000000000c0 ,
                        0x0000000000000000000000000000000000000000000000000080000080000000 ,
                        0x80800080000000800080008080000080808000c0c0c0000000ff0000ff000000 ,
                        0xffff00ff000000ff00ff00ffff0000ffffff0000000000000000000000000000 ,
                        0x000000000000000000000077777777777777007fb8b8b8b8b8b7007f8b8b8b8b ,
                        0x8b87007fb8b8b8b8b8b7007f8b8b8b8b8b87007fb8b8b8b8b8b7007f8b8b8b8b ,
                        0x8b87007fb8b8b8b8b8b7007ffffffffffff70078b8b8b877777700078b8b8700 ,
                        0x00000000777770000000000000000000000000ffff0000ffff00008001000000 ,
                        0x0100000001000000010000000100000001000000010000000100000001000000 ,
                        0x0100000003000080ff0000c1ff0000ffff0000030000006c7400003604000000 ,
                        0x00010002002020100000000000e8020000260000001010100000000000280100 ,
                        0x000e030000280000002000000040000000010004000000000080020000000000 ,
                        0x0000000000000000000000000000000000000080000080000000808000800000 ,
                        0x00800080008080000080808000c0c0c0000000ff0000ff000000ffff00ff0000 ,
                        0x00ff00ff00ffff0000ffffff0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000733333333333333333333333333300007b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b8b83000078b8b8b8b8b8b8b8b8b8b8b8b8b300078b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b8b800007b8b8b8b8b8b8b8b8b8b8b8b8b88000078b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b8b700007b8b8b8b8b8b8b8b8b8b8b8b8b870007b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b8b030078b8b8b8b8b8b8b8b8b8b8b8b8b803007b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b87030078b8b8b8b8b8b8b8b8b8b8b8b8b703078b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b807307b8b8b8b8b8b8b8b8b8b8b8b8b8b073078b8b8 ,
                        0xb8b8b8b8b8b8b8b8b8b8b70b307b8b8b8b8b8b8b8b8b8b8b8b8b8708307fffff ,
                        0xfffffffffffffffffffff07b300778888888888888888888888888b8300007fb ,
                        0x8b8b8b8b8b8b8b8b8b8b8b8b300007f8b8b8b8b8b8b8b8b8b8b8b8b8300007fb ,
                        0x8b8b8b8b8b8b8b8b8b8b8b8b300007f8b8b8b8b8b8b8bfffffffffff000007fb ,
                        0x8b8b8b8b8b8b8777777777770000007fb8b8b8b8b8b870000000000000000007 ,
                        0xfb8b8b8b8b87000000000000000000007fffffffff7000000000000000000000 ,
                        0x0777777777000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000fffffffffffffffffffffffff0000001e00000 ,
                        0x00e0000000e0000000c0000000c0000000c0000000c000000080000000800000 ,
                        0x0080000000800000000000000000000000000000000000000000000000800000 ,
                        0x00e0000000e0000000e0000000e0000001e0000003f0001ffff8003ffffc007f ,
                        0xfffe00ffffffffffffffffffff28000000100000002000000001000400000000 ,
                        0x00c0000000000000000000000000000000000000000000000000008000008000 ,
                        0x000080800080000000800080008080000080808000c0c0c0000000ff0000ff00 ,
                        0x0000ffff00ff000000ff00ff00ffff0000ffffff000000000000000000000000 ,
                        0x000000000000000000000000000077777777777700007fb8b8b8b8b70007fb8b ,
                        0x8b8b8b807007f8b8b8b8b870707f8b8b8b8b8b07707ffffffffff70870777777 ,
                        0x7777777b7007f8b8b8b8b8b87007fb8b8b8fffff7007f8b8b8f7777770007fff ,
                        0xff7000000000077777000000000000000000000000ffff0000ffff0000e00000 ,
                        0x00c0000000c00000008000000080000000000000000000000000000000800000 ,
                        0x008000000080010000c07f0000e0ff0000ffff0000040000006c740000360400 ,
                        0x000000010002002020100000000000e802000026000000101010000000000028 ,
                        0x0100000e03000028000000200000004000000001000400000000008002000000 ,
                        0x0000000000000000000000000000000000000000008000008000000080800080 ,
                        0x000000800080008080000080808000c0c0c0000000ff0000ff000000ffff00ff ,
                        0x000000ff00ff00ffff0000ffffff000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000880000000000000 ,
                        0x0000000000000009077880000000000000000000000000990777788000000000 ,
                        0x0000000000000999800777788000000000000000000099980990077778800000 ,
                        0x0000000000099980999990077770000000000000009998099999999007700000 ,
                        0x0000000009998099999999999000000000000000999809999999999999900000 ,
                        0x0000000999809999999999999990000000000009980999999999999999000000 ,
                        0x0000000980999999999999999000000000000008099999999999999900000000 ,
                        0x0000000009999999999999900000000000000000000999999999990000000000 ,
                        0x0000000000000999999990000000000000000000000000099999000000000000 ,
                        0x0000000000000000099000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffc7ffffff81ffffff007ffffe001ffffc0007ff ,
                        0xf80001fff00003ffe00003ffc00003ff800001ff000003ff000007ff00000fff ,
                        0x00001fff80003fffe0007ffff800fffffe01ffffff83ffffffe7ffffffffffff ,
                        0xffffffffffffffffffffffffffffff2800000010000000200000000100040000 ,
                        0x000000c000000000000000000000000000000000000000000000000000800000 ,
                        0x8000000080800080000000800080008080000080808000c0c0c0000000ff0000 ,
                        0xff000000ffff00ff000000ff00ff00ffff0000ffffff00000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000907000000 ,
                        0x0000009807880000000009800077800000009809990070000009809999990000 ,
                        0x0008099999990000000099999990000000000999990000000000000990000000 ,
                        0x0000000000000000000000000000000000000000000000ffff0000ffff0000ff ,
                        0xff0000ff8f0000ff030000fe000000fc010000f8010000f0000000f0010000f0 ,
                        0x030000f8070000fe0f0000ff9f0000ffff0000ffff0000050000006c74000036 ,
                        0x0400000000010002002020100000000000e80200002600000010101000000000 ,
                        0x00280100000e0300002800000020000000400000000100040000000000800200 ,
                        0x0000000000000000000000000000000000000000000000800000800000008080 ,
                        0x0080000000800080008080000080808000c0c0c0000000ff0000ff000000ffff ,
                        0x00ff000000ff00ff00ffff0000ffffff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000088000000000 ,
                        0x0000000000000000000907788000000000000000000000000000000778800000 ,
                        0x000000000000700077770880077880000000000000070fff000008fff0077880 ,
                        0x000000000070fffff8808ffffff0077000000000070fffff8808ff77fffff000 ,
                        0x0000000070fffff8808fffff77fffff0000000070fffff8808ff77ffff77ff00 ,
                        0x00000070fffff8808fffff77fffff0000000070fffff8808ff77ffff77ff0000 ,
                        0x000070fffff8808fffff77fffff0000000000fffff8808ff77ffff77ff000000 ,
                        0x0000000ff880000fff77fffff000000000000000000000000fff77ff00000000 ,
                        0x0000000000000000000ffff000000000000000000000000000000f0000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffc7ffffff81ffffff007fffc0001fff8000 ,
                        0x07ff000001fe000003fc000003f8000003f0000007e000000fc000001f800000 ,
                        0x3f0000007fe02000fff87801fffffe03ffffff87ffffffefffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffff280000001000000020000000010004 ,
                        0x0000000000c00000000000000000000000000000000000000000000000000080 ,
                        0x00008000000080800080000000800080008080000080808000c0c0c0000000ff ,
                        0x0000ff000000ffff00ff000000ff00ff00ffff0000ffffff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000088000 ,
                        0x0000000000000880000000fff880f00880000fff880ffff00000fff880ffffff ,
                        0x0008ff880ffffff000008880ffffff000000000000fff0000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000ffff0000ffff00 ,
                        0x00ffff0000ff9f0000ff070000f0010000e0000000c001000080010000000300 ,
                        0x0000070000c40f0000ff1f0000ffff0000ffff0000ffff0000060000006c7400 ,
                        0x00360400000000010002002020100001000400e8020000260000001010100001 ,
                        0x000400280100000e030000280000002000000040000000010004000000000080 ,
                        0x0200000000000000000000100000000000000000000000000080000080000000 ,
                        0x80800080000000800080008080000080808000c0c0c0000000ff0000ff000000 ,
                        0xffff00ff000000ff00ff00ffff0000ffffff0000000077777777777777777777 ,
                        0x77770000000000000000000000000000000700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbb0000000000000bbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbb0000000000000bbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbb0000000000000bbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbb0000000000000bbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb0700000000bbbbbbbbbbbbbbbbbbbb ,
                        0xbb0700000000bbbbbbbbbbbbbbbbbbbbbb070000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000fc000003f8000003f8000003f8 ,
                        0x000003f8000003f8000003f8000003f8000003f8000003f8000003f8000003f8 ,
                        0x000003f8000003f8000003f8000003f8000003f8000003f8000003f8000007ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffff28000000100000002000000001 ,
                        0x00040000000000c0000000000000000000000000000000000000000000000000 ,
                        0x008000008000000080800080000000800080008080000080808000c0c0c00000 ,
                        0x00ff0000ff000000ffff00ff000000ff00ff00ffff0000ffffff000000000000 ,
                        0x00000000bfbfbfbfbfbf0000fb000000fbfb0000bfbfbfbfbfbf0000fb00000b ,
                        0xfbfb0000bfbfbfbfbfbf0000fb0000000bfb0000bfbfbfbfb55f0000fbfbfbfb ,
                        0xf55b0000bfbfbfbfbfbf00000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000008001000080 ,
                        0x0100008001000080010000800100008001000080010000800100008001000080 ,
                        0x01000080010000ffff0000ffff0000ffff0000ffff0000ffff0000070000006c ,
                        0x740000360400000000010002002020100001000400e802000026000000101010 ,
                        0x0001000400280100000e03000028000000200000004000000001000400000000 ,
                        0x0080020000000000000000000010000000000000000000000000008000008000 ,
                        0x000080800080000000800080008080000080808000c0c0c0000000ff0000ff00 ,
                        0x0000ffff00ff000000ff00ff00ffff0000ffffff000000007777777777777777 ,
                        0x777777770000000000000000000000000000000700000000ffffffffffffffff ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000fff0000000000000 ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000fff0000000000000ffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000fff0000000000000 ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000fff0000000000000ffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000ffffffffffffffffffffff0700000000ffffffffffffffff ,
                        0xffffff0700000000ffffffffffffffffffffff07000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000fc000003f8000003f80000 ,
                        0x03f8000003f8000003f8000003f8000003f8000003f8000003f8000003f80000 ,
                        0x03f8000003f8000003f8000003f8000003f8000003f8000003f8000003f80000 ,
                        0x07ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffff2800000010000000200000 ,
                        0x000100040000000000c000000000000000000000000000000000000000000000 ,
                        0x0000008000008000000080800080000000800080008080000080808000c0c0c0 ,
                        0x000000ff0000ff000000ffff00ff000000ff00ff00ffff0000ffffff00000000 ,
                        0x000000000000bfbfbfbfbfbf0000fb000000fbfb0000bfbfbfbfbfbf0000fb00 ,
                        0x000bfbfb0000bfbfbfbfbfbf0000fb0000000bfb0000bfbfbfbfb55f0000fbfb ,
                        0xfbfbf55b0000bfbfbfbfbfbf0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000800100 ,
                        0x0080010000800100008001000080010000800100008001000080010000800100 ,
                        0x008001000080010000ffff0000ffff0000ffff0000ffff0000ffff0000080000 ,
                        0x006c740000360400000000010002002020100001000400e80200002600000010 ,
                        0x10100001000400280100000e0300002800000020000000400000000100040000 ,
                        0x0000008002000000000000000000001000000000000000000000000000800000 ,
                        0x8000000080800080000000800080008080000080808000c0c0c0000000ff0000 ,
                        0xff000000ffff00ff000000ff00ff00ffff0000ffffff00000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000cc0 ,
                        0x0000000000000000000000000000ccc0000000000000000000000000000ccccc ,
                        0x00000000000000000000000000ccc0cc0000000000000000000000000ccc000c ,
                        0xc00000000000000000000000ccc0000cc0000000000000000000000000000000 ,
                        0xcc000000000000000000000000000000cc000000000000000000000000000000 ,
                        0x0cc000000000000000000000000000000cc00000000000000000000000000000 ,
                        0x00cc000000000000000000000000000000cc0000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffe7ffffffc7ffffff83ffffff13fffff ,
                        0xe39fffffc79fffffffcfffffffcfffffffe7ffffffe7fffffff3fffffff3ffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffff280000001000000020 ,
                        0x0000000100040000000000c00000000000000000000000000000000000000000 ,
                        0x00000000008000008000000080800080000000800080008080000080808000c0 ,
                        0xc0c0000000ff0000ff000000ffff00ff000000ff00ff00ffff0000ffffff0000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000070000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000000000ff ,
                        0xff0000feff0000fcff0000f8ff0000f07f0000e27f0000c63f0000ff3f0000ff ,
                        0x9f0000ff8f0000ffcf0000ffe70000fff30000fff90000fffd0000ffff000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000007300 ,
                        0xf81f5c0200005802200000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000ccc0cc0000000000000000000000000ccc000c ,
                        0xc00000000000000000000000ccc0000cc0000000000000000000000000000000 ,
                        0xcc000000000000000000000000000000cc000000000000000000000000000000 ,
                        0x0cc000000000000000000000000000000cc00000000000000000000000000000 ,
                        0x00cc000000000000000000000000000000cc0000000000000000000000000000 ,
                        0x000cc000000000000000000000000000000cc000000000000000000000000000 ,
                        0x0000cc000000000000000000000000000000cc00000000000000000000000000 ,
                        0x00000c0000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000ffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7ffffffe7ffffffc3ffffff83ffffff13fffff ,
                        0xe39fffffc79fffffffcfffffffcfffffffe7ffffffe7fffffff3fffffff3ffff ,
                        0xfff9fffffff9fffffffcfffffffcfffffffeffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffff280000001000000020 ,
                        0x0000000100040000000000c00000000000000000000000000000000000000000 ,
                        0x00000000008000008000000080800080000000800080008080000080808000c0 ,
                        0xc0c0000000ff0000ff000000ffff00ff000000ff00ff00ffff0000ffffff0000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000070000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000000000ff ,
                        0xff0000feff0000fcff0000f8ff0000f07f0000e27f0000c63f0000ff3f0000ff ,
                        0x9f0000ff8f0000ffcf0000ffe70000fff30000fff90000fffd0000ffff000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OLEClass ="ImageListCtrl"
                    Class ="MSComctlLib.ImageListCtrl.2"

                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =4491
                    Top =7426
                    Width =1482
                    Height =330
                    FontSize =9
                    TabIndex =24
                    Name ="btnSendEpost"
                    Caption ="Send epost"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =283
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
Public myDb As DAO.Database



Private Sub btnArbeidsplan_Click()
On Error GoTo Err_btnArbeidsplan_Click
    Dim rsPm As DAO.Recordset
    Dim rsLarer As DAO.Recordset
    Dim rsStilling As DAO.Recordset
    Dim rsKurs As DAO.Recordset
    Dim rsTerm As DAO.Recordset
    Dim sqlLarer As String, sqlStilling As String, strComment As String
    Dim MylvwEmne As ListView
    Dim strEmne As String, StAndel As Integer, dblArTimer As Double
    Dim strSted As String
    Dim strSheet As String, strAgrp1 As String, strAgrp2 As String
    Dim strSem1 As String, strSem2 As String
    'Set MylvwEmne = lvwKurs.Object
    Dim APath As String, AFile As String, TFile As String
    Dim strNavn As String, strStilling As String
    Dim CellNo As String
    Dim AntKurs As Integer, i As Integer
    
    ' Loading Excel (late or early)
    ' 0 if Late Binding
    ' 1 if Reference to Excel set.
    #Const ExcelRef = 0
    #If ExcelRef = 0 Then ' Late binding
        Dim ExApp As Object
        Dim APlan As Object
        Set ExApp = CreateObject("Excel.Application")
        On Error GoTo Err_btnArbeidsplan_Click
    #Else
        ' a reference to MS Excel <version number> Object Library must be specified
        Dim ExApp As Excel.Application
        Dim APlan As Excel.Worksheet
        Set ExApp = New Excel.Application
    #End If

    Set myDb = CurrentDb
   
    ' åpner kurstabell og sjekker innhold
    Set rsKurs = myDb.OpenRecordset("tblAplan")
    AntKurs = rsKurs.RecordCount
    If AntKurs = 0 Or GL_FID <= 0 Then
        MsgBox "Ingenting å overføre til arbeidsplan", , OIS_Title
        Exit Sub
    End If
    'AntKurs = MylvwEmne.ListItems.Count
    
    sqlLarer = "SELECT * FROM tblLarer WHERE LarerID = " & GL_FID & ";"
    Set rsLarer = myDb.OpenRecordset(sqlLarer)
    If Not IsNull(rsLarer!Andel) And rsLarer!Andel <> "" Then
        StAndel = rsLarer!Andel
    Else
        StAndel = 100
    End If
    If StAndel = 0 Then
        MsgBox "Trenger ikke arbeidsplan", vbInformation, OIS_Title
        Exit Sub
    End If
    strNavn = SnuddNavn(rsLarer!Navn)
    sqlStilling = "SELECT * FROM tblStilling WHERE StKode = '" & rsLarer!StKode & "';"
    Set rsStilling = myDb.OpenRecordset(sqlStilling)
    If rsStilling.RecordCount = 0 Then strStilling = "Ikke oppgitt"

    Set rsTerm = myDb.OpenRecordset("tblTermer")
    Set rsPm = myDb.OpenRecordset("tblParameter")
    Select Case rsLarer!SpraakAplan
        Case Is = "E" ' engelsk
            APath = rsPm!AvdPath
            TFile = APath & "\" & rsPm!ArbPlanE
            strSheet = rsTerm!TmSheetE
            strAgrp1 = rsTerm!TmAgrp1E
            strAgrp2 = rsTerm!TmAgrp2E
            strSem1 = rsTerm!TmSem1E
            strSem2 = rsTerm!TmSem2E
            strStilling = rsStilling!StNavnE
        Case Else 'hvis ingenting er oppgitt, eller er "N", er det norsk som gjelder
            APath = rsPm!AvdPath
            TFile = APath & "\" & rsPm!ArbPlan
            strSheet = rsTerm!TmSheet
            strAgrp1 = rsTerm!TmAgrp1
            strAgrp2 = rsTerm!TmAgrp2
            strSem1 = rsTerm!TmSem1
            strSem2 = rsTerm!TmSem2
            strStilling = rsStilling!StNavn
    End Select
    
    ' lager ny excel-fil med navnet til aktuell lærer
    
    AFile = APath & "\" & strNavn & "_" & rsPm!studyYear & ".xlsx"
    FileCopy TFile, AFile
    
    ExApp.Workbooks.Open (AFile)
    Set APlan = ExApp.Worksheets(strSheet) ' velger første side
    'overfører data til arbeidsplanen
    'APlan.LeftHeader = rsPm!studyYear
    APlan.Range("D1") = strNavn
    APlan.Range("A2") = rsPm!studyYear
    APlan.Range("H1") = strStilling
    APlan.Range("H2") = StAndel
    If rsLarer!Over60 = True Then
        APlan.Range("D2") = strAgrp2
        dblArTimer = 1650 * (StAndel / 100)
        APlan.Range("E2") = dblArTimer
    Else ' under 60 år
        APlan.Range("D2") = strAgrp1
        dblArTimer = 1687.5 * (StAndel / 100)
        APlan.Range("E2") = dblArTimer
    End If
    APlan.Range("F4") = strSem1 & " " & Left(rsPm!studyYear, 4)
    APlan.Range("H4") = strSem2 & " " & Right(rsPm!studyYear, 4)
    ' henter data fra kurstabell (tblAplan)
    i = 1
    rsKurs.MoveFirst
    Do While Not rsKurs.EOF
        APlan.Range("A" & CStr(5 + i)) = rsKurs!Kurslinje
        APlan.Range("E" & CStr(5 + i)) = rsKurs!StpSem
        i = i + 1
        rsKurs.MoveNext
    Loop
    'APlan.SaveAs (AFile)
    'viser arbeisplanen og overlater kontroll til bruker
    ExApp.Visible = True
    ExApp.UserControl = True
Exit_btnArbeidsplan_Click:
    rsLarer.Close
    rsPm.Close
    rsTerm.Close
    rsKurs.Close
    rsStilling.Close
    myDb.Close
    Exit Sub

Err_btnArbeidsplan_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnArbeidsplan_Click

End Sub

Private Sub btnDelete_Click()
On Error GoTo Err_btnDelete_Click
    Dim MylvwEmne As ListView
    Set MylvwEmne = Me.lvwKurs.Object
    Dim msg As String
    Dim response As Integer
    Dim sqlStr As String
 
    If MylvwEmne.ListItems.Count < 1 Or Me.btnUpdate.Caption <> strUpdate Then
        msg = "Ingenting å fjerne. Du må velge et emne."
        MsgBox msg, vbOKOnly + vbExclamation, OIS_Title
        Exit Sub
    End If
    
    'msg = "Fjerne " & Me.lstLarerEmne.Column(1) & " fra "
    'msg2 = vbNewLine & Me.lstFaglarer.Column(1)
    'response = MsgBox(msg & msg2, vbYesNo + vbExclamation + vbDefaultButton2, OIS_Title)
    'If response = vbNo Then
    '    Exit Sub
    'End If
    Set myDb = CurrentDb()
    sqlStr = "DELETE * FROM tblLarerEmne WHERE LarerID = " & GL_FID & " AND EmneID = " & GL_EID & ";"
    myDb.Execute (sqlStr)
    ' Reload listview
    Call ForeleserSineKurs(GL_FID)
    Me.btnUpdate.Caption = strInsert
Exit_btnDelete_Click:
    'myDb.Close
    Exit Sub

Err_btnDelete_Click:
    MsgBox Err.Description
    Resume Exit_btnDelete_Click

End Sub

Private Sub btnNullstill_Click()
    Dim MylvwKurs As ListView
    Dim NoOfEmne As Integer
    Dim stbStatus As CustomControl
    Set stbStatus = Me.stbStatusLine
    
    Set MylvwKurs = Me.lvwKurs.Object
    MylvwKurs.ListItems.Clear
    MylvwKurs.View = lvwReport
    stbStatus.Panels(2) = ""
    Me.txtLarer = ""
    GL_FID = 0
    Me.btnUpdate.Caption = strInsert

End Sub

Private Sub btnSendArbeidsplan_Click()
On Error GoTo Err_btnSendArbeidsplan_Click
    Dim rsPm As DAO.Recordset
    Dim rsLarer As DAO.Recordset
    Dim rsStilling As DAO.Recordset
    Dim OApp As Outlook.Application
    Dim OAppEpost As MailItem
    Dim sqlLarer As String, sqlStilling As String, strComment As String
    Dim MylvwEmne As ListView
    Dim fSize As Long
    Dim strEmne As String, StAndel As Integer, dblArTimer As Double
    Dim strSted As String, strStudyYear As String, strSpraak As String
    Dim SendTil As String, Vedlegg As String
    Set MylvwEmne = lvwKurs.Object
    Dim ExApp As Object
    Dim APlan As Excel.Worksheet
    Dim APath As String, AFile As String, TFile As String
    Dim Navn As String, Stilling As String
    Dim CellNo As String
    Dim AntKurs As Integer, i As Integer
    Set myDb = CurrentDb
    sqlLarer = "SELECT * FROM tblLarer WHERE LarerID = " & GL_FID & ";"
    Set rsLarer = myDb.OpenRecordset(sqlLarer)
    ' navn, stilling, sti og filnavn
    Navn = SnuddNavn(rsLarer!Navn)
    strSpraak = rsLarer!SpraakAplan
    If Not IsNull(rsLarer!Andel) And rsLarer!Andel <> "" Then
        StAndel = rsLarer!Andel
    Else
        StAndel = 100
    End If
    If StAndel = 0 Then
        MsgBox "Trenger ikke arbeidsplan, men skal ha kontrakt", vbInformation, OIS_Title
        Exit Sub
    End If
    
    Set rsPm = myDb.OpenRecordset("tblParameter")
    APath = rsPm!AvdPath
    strStudyYear = rsPm!studyYear
    AFile = APath & "\" & Navn & "_" & strStudyYear & ".xlsx"
    fSize = FileLen(AFile)
    If fSize > 0 Then 'arbeidsplanen funnet og eksisterer
        If rsLarer!SendtArbPlan = True Then 'har sendt arbeidsplan tidligere
            i = MsgBox(Navn & " har allerede mottatt arbeideplan. " & vbNewLine & "Vil du sende planen pånytt? ", vbYesNo, OIS_Title)
            If i = vbNo Then Exit Sub
        End If
        SendTil = ""
        If Not IsNull(rsLarer!Epost) And rsLarer!Epost <> "" Then
            SendTil = rsLarer!Epost
        Else
            MsgBox "Har ikke epostadressen til " & Navn, vbInformation, OIS_Title
            Exit Sub
        End If
        Set OApp = New Outlook.Application
        Set OAppEpost = OApp.CreateItem(olMailItem)
        With OAppEpost
            .To = SendTil
            If strSpraak = "E" Then
                .Subject = "Working plan " & strStudyYear
                .Body = "Draft of working plan for the study year " & strStudyYear & " is attached."
            Else
                .Subject = "Arbeidsplan " & strStudyYear
                .Body = "Utkast til arbeidsplan for " & strStudyYear & " følger vedlagt."
            End If
            .Attachments.Add (AFile)
            .Display
        End With
        rsLarer.Edit
            rsLarer!SendtArbPlan = True
        rsLarer.Update
        MsgBox "Arbeidsplan er sendt til " & Navn, vbInformation, OIS_Title
    End If
Exit_btnSendArbeidsplan_Click:
    rsLarer.Close
    rsPm.Close
    myDb.Close
    Exit Sub

Err_btnSendArbeidsplan_Click:
    If Err.Number = 53 Then  ' File not found
        MsgBox "Finner ikke arbeidsplanen til " & Navn & vbNewLine & "Bruk 'Overfør til arbeidsplan' knappen.", , OIS_Title
    Else
        MsgBox Err.Description, , OIS_Title
    End If
    Resume Exit_btnSendArbeidsplan_Click
End Sub

Private Sub btnSendEpost_Click()
On Error GoTo Err_btnSendEpost_Click
    Dim rsPm As DAO.Recordset
    Dim rsLarer As DAO.Recordset
    Dim rsStilling As DAO.Recordset
    Dim OApp As Outlook.Application
    Dim OAppEpost As MailItem
    Dim sqlLarer As String, sqlStilling As String, strComment As String
    Dim MylvwEmne As ListView
    Dim strEmne As String, StAndel As Integer, dblArTimer As Double
    Dim strSted As String, strStudyYear As String, strMelding As String
    Dim strSem As String
    Dim SendTil As String, Vedlegg As String
    Set MylvwEmne = lvwKurs.Object
    Dim Navn As String, Stilling As String
    Dim CellNo As String
    Dim AntKurs As Integer, i As Integer
    Set myDb = CurrentDb
    sqlLarer = "SELECT * FROM tblLarer WHERE LarerID = " & GL_FID & ";"
    Set rsLarer = myDb.OpenRecordset(sqlLarer)
    ' navn, stilling, sti og filnavn
    If IsNull(Me.txtLarer) Or Me.txtLarer = "" Then
        MsgBox "Ingen å sende til. Du må velge en person fra foreleserlisten", vbInformation, OIS_Title
        Exit Sub
    End If

    Navn = Fnavn(rsLarer!Navn)
    Set rsPm = myDb.OpenRecordset("tblParameter")
    strStudyYear = rsPm!studyYear
    SendTil = ""
    If Not IsNull(rsLarer!Epost) And rsLarer!Epost <> "" Then
        SendTil = rsLarer!Epost
    Else
        MsgBox "Finner ikke epostadressen til " & Navn, vbInformation, OIS_Title
        'Exit Sub
    End If
'Lager melding
    strMelding = ""
    AntKurs = MylvwEmne.ListItems.Count
    If MylvwEmne.ListItems.Count > 0 Then
        For i = 1 To AntKurs
            With MylvwEmne.ListItems(i)
                strEmne = .Text & " " & .SubItems(1) & ", "
                Select Case .SubItems(3)
                    Case "H": strSem = "Høst, "
                    Case "V": strSem = "Vår, "
                    Case "H+V": strSem = "Høst og vår, "
                    Case Else: strSem = ", "
                End Select
                strMelding = strMelding & vbNewLine & strEmne & .SubItems(2) & " stp, " & strSem & .SubItems(4)
            End With
        Next i
    End If
    
    Set OApp = New Outlook.Application
    Set OAppEpost = OApp.CreateItem(olMailItem)
    With OAppEpost
        .To = SendTil
        .Subject = "Undervisningsoppgaver " & strStudyYear
        .Body = "Hei " & Navn & "," & vbNewLine & vbNewLine & _
                "Dine undervisningsoppgaver for studieåret " & strStudyYear & " er foreløpig slik:" & vbNewLine & _
                strMelding & vbNewLine & vbNewLine & "Ta kontakt med studieleder om du har spørsmål eller kommentarer."
        .Display
    End With
Exit_btnSendEpost_Click:
    rsLarer.Close
    rsPm.Close
    myDb.Close
    Exit Sub

Err_btnSendEpost_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnSendEpost_Click

End Sub

Private Sub btnUpdate_Click()
On Error GoTo Err_btnUpdate_Click
    Dim rsLarerEmne         As DAO.Recordset
    Dim sqlLarerEmne        As String
    Dim msg                 As String
    Dim sqlStr              As String
 
    If Me.txtLarer = "" Then
        msg = "Du må velge en lærer før du kan legge til eller endre et emne."
        MsgBox msg, vbOKOnly + vbExclamation, OIS_Title
        Exit Sub
    End If
    
    Set myDb = CurrentDb
    Select Case btnUpdate.Caption
        Case strInsert
            sqlStr = "INSERT INTO tblLarerEmne (LarerID, EmneID, Studiepoeng, Kommentar) " & _
                    "VALUES('" & GL_FID & "', '" & GL_EID & "', '" & CSng(Me.txtBelastning) & "', '" & Me.txtMerknader & "');"
            myDb.Execute (sqlStr)
            Call ForeleserSineKurs(GL_FID)
        Case strUpdate
            sqlStr = "UPDATE tblLarerEmne SET Studiepoeng ='" & CSng(Me.txtBelastning) & "', " & _
                    "Kommentar = '" & Me.txtMerknader & "' " & _
                    "WHERE LarerID = " & GL_FID & " AND EmneID = " & GL_EID & ";"
            myDb.Execute (sqlStr)
            Call ForeleserSineKurs(GL_FID)
    End Select
    Me.btnUpdate.Caption = strUpdate

Exit_btnUpdate_Click:
    Exit Sub

Err_btnUpdate_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnUpdate_Click
End Sub


Private Sub btnOk_Click()
On Error GoTo Err_btnOk_Click

    DoCmd.Close

Exit_btnOk_Click:
    Exit Sub

Err_btnOk_Click:
    MsgBox Err.Description
    Resume Exit_btnOk_Click
End Sub

Private Sub cboStilling_Click()
    Dim ValgtStilling As String
    ValgtStilling = Me.cboStilling.Column(0)
    Call FyllPersonListe(ValgtStilling)
    Me.frPersoner = 0
End Sub

Private Sub cboEmnegruppe_Click()
    Dim ValgtFag As String
    ValgtFag = Me.cboEmnegruppe.Column(0)
    Call FyllEmneliste(ValgtFag, "")
    Me.frEmner = 0
End Sub


Private Sub chkFerdig_Click()
On Error GoTo Err_chkFerdig_Click
    Dim sqlstring As String
    Dim blnFerdig As Boolean
    Dim ValgtFag As String
    Dim StartFag As String
    Dim msg1 As String
    Dim msg2 As String
    If Me.txtEmne = "" Then
        msg1 = "Ingenting å ferdig-markere"
        msg2 = "Du må velge et emne"
        MsgBox msg1 & vbNewLine & msg2, vbExclamation, OIS_Title
        Exit Sub
    End If
    StartFag = Left(Me.txtEmne, 6)
    blnFerdig = Me.chkFerdig.Value

    sqlstring = "UPDATE tblEmne SET tblEmne.Ferdig =" & blnFerdig & _
                " WHERE EmneID = " & GL_EID & ";"
                
    Set myDb = CurrentDb()
    myDb.Execute (sqlstring)
    If Me.frEmner = 1 Then Call FyllEmneliste("", StartFag)
    If Me.cboEmnegruppe <> "" Then
        ValgtFag = Me.cboEmnegruppe.Column(0)
        Call FyllEmneliste(ValgtFag, StartFag)
    End If
Exit_chkFerdig_Click:
    Exit Sub

Err_chkFerdig_Click:
    MsgBox Err.Description, , GP_title
    Resume Exit_chkFerdig_Click
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load
    Dim sqlstr1 As String, sqlstr2 As String
    
    Dim itmX As ListItem
    Dim MylvwKurs As ListView
    Set MylvwKurs = Me.lvwKurs.Object
    MylvwKurs.ListItems.Clear
    MylvwKurs.View = lvwReport
    
    Call FyllPersonListe("")
    'Me.lstEmne.RowSource = sqlstr2
    Call FyllEmneliste("", "")
    Me.cboEmnegruppe.RowSource = "SELECT Kode from tblEmnekode ORDER BY Kode"
    Me.cboStilling.RowSource = "SELECT StNavn from tblStilling ORDER BY StNavn"
    Me.btnUpdate.Caption = strInsert
    Me.txtLarer = ""
    Me.frEmner = 1      'viser alle emner
    Me.frPersoner = 1   ' viser alle ansatte
    Call LockTxtboxes(Me)
Exit_Form_Load:
    Exit Sub

Err_Form_Load:
    MsgBox Err.Description
    Resume Exit_Form_Load
End Sub

Public Sub FyllEmneliste(ValgtEmne As String, StartEmne As String)
    Dim lvwRS As DAO.Recordset
    Dim itmX As ListItem
    Dim stbStatus As CustomControl
    Dim MylvwEmne As ListView
    Dim sqlStr As String, StEmne As String
    Dim NoOfEmne As Integer
    Set stbStatus = Me.stbStatusLine
    Set MylvwEmne = Me.lvwAlleEmner.Object
    StEmne = StartEmne
    If MylvwEmne.ListItems.Count < 17 Then StEmne = "" 'alle emner vises i listviewet
    
    If ValgtEmne = "" Then
        If StartEmne = "" Then
            sqlStr = "SELECT * FROM tblEmne " & _
                "WHERE Aktiv = True ORDER BY Emnekode;"
        Else ' startemne <> ""
            sqlStr = "SELECT * FROM tblEmne " & _
                "WHERE Aktiv = True AND Left(Emnekode,6) >= '" & StEmne & "' ORDER BY Emnekode;"
        End If
    Else  'valgt emne <> ""
        If StEmne = "" Then
            sqlStr = "SELECT * FROM tblEmne " & _
                "WHERE Aktiv = True AND Left(Emnekode,3) = '" & ValgtEmne & "' ORDER BY Emnekode;"
        Else ' startemne <> ""
            sqlStr = "SELECT * FROM tblEmne " & _
                "WHERE Aktiv = True AND Left(Emnekode,3) = '" & ValgtEmne & "' AND Left(Emnekode,6) >= '" & StEmne & "' ORDER BY Emnekode;"
        End If
    End If
    'Me.lstEmne.RowSource = sqlstr
    Set myDb = CurrentDb()
    Set lvwRS = myDb.OpenRecordset(sqlStr, dbOpenDynaset)
    
    MylvwEmne.ListItems.Clear
    MylvwEmne.View = lvwReport
         
    While Not lvwRS.EOF
    ' Kurskode
        If lvwRS.Fields("Emnekode") <> "" Then
            If lvwRS.Fields("Ferdig") = True Then
                Set itmX = MylvwEmne.ListItems.Add(, , lvwRS.Fields("Emnekode"), , 8)
            Else
                Set itmX = MylvwEmne.ListItems.Add(, , lvwRS.Fields("Emnekode"))
            End If
        Else
            Set itmX = MylvwEmne.ListItems.Add(, , " ")
        End If
    ' Subitem 1: Navn
        If lvwRS.Fields("Emnenavn") <> "" Then
            itmX.SubItems(1) = lvwRS.Fields("Emnenavn")
        Else
            itmX.SubItems(1) = ""
        End If
    ' Subitem 2: Studiepoeng
        If lvwRS.Fields("Studiepoeng") <> "" And Not IsNull(lvwRS.Fields("Studiepoeng")) Then
            itmX.SubItems(2) = lvwRS.Fields("Studiepoeng")
        Else
            itmX.SubItems(2) = ""
        End If
    
    ' Subitem 3: Semester
        If lvwRS.Fields("Semester") <> "" Then
            itmX.SubItems(3) = lvwRS.Fields("Semester")
        Else
            itmX.SubItems(3) = ""
        End If
        
    ' Subitem 4: Undervisningssted
        If lvwRS.Fields("Sted") <> "" Then
            itmX.SubItems(4) = lvwRS.Fields("Sted")
        Else
            itmX.SubItems(4) = "Molde"
        End If

    ' EmneID
        If lvwRS.Fields("EmneID") <> "" And Not IsNull(lvwRS.Fields("EmneID")) Then
            itmX.SubItems(5) = CStr(lvwRS.Fields("EmneID"))
        Else
            itmX.SubItems(5) = ""
        End If
    ' Subitem 6: Merknader
        If lvwRS.Fields("Comment") <> "" Then
            itmX.SubItems(6) = lvwRS.Fields("Comment")
        Else
            itmX.SubItems(6) = ""
        End If
    'Subitem 6: Ferdig planlagt
        If lvwRS.Fields("Ferdig") <> "" And Not IsNull(lvwRS.Fields("Ferdig")) Then
            itmX.SubItems(7) = lvwRS.Fields("Ferdig")
        Else
            itmX.SubItems(7) = 0
        End If

        lvwRS.MoveNext
    Wend
    stbStatus.Panels(3) = "Totalt antall emner: " & MylvwEmne.ListItems.Count
    lvwRS.Close
    myDb.Close

End Sub

Public Sub FyllPersonListe(ValgtStilling As String)
    Dim lvwRS As DAO.Recordset
    Dim itmX As ListItem
    Dim MylvwPerson As ListView
    Dim sqlStr As String
    Dim NoOfEmne As Integer
    Dim stbStatus As CustomControl
    Set stbStatus = Me.stbStatusLine
    Set myDb = CurrentDb()
    If ValgtStilling = "" Then
        sqlStr = "SELECT tblLarer.LarerID, tblLarer.Navn, tblStilling.StNavn, tblLarer.Stkode, tblLarer.Andel, tblLarer.Merk " & _
            "FROM tblLarer INNER JOIN tblStilling ON tblLarer.Stkode = tblStilling.StKode " & _
            "ORDER BY tblLarer.Navn;"
    Else
        sqlStr = "SELECT tblLarer.LarerID, tblLarer.Navn, tblStilling.StNavn, tblLarer.Stkode, tblLarer.Andel, tblLarer.Merk " & _
            "FROM tblLarer INNER JOIN tblStilling ON tblLarer.Stkode = tblStilling.StKode " & _
            "WHERE StNavn = '" & ValgtStilling & "' ORDER BY tblLarer.Navn;"
    End If
    Set lvwRS = myDb.OpenRecordset(sqlStr, dbOpenDynaset)
    
    Set MylvwPerson = Me.lvwPersonale.Object
    MylvwPerson.ListItems.Clear
    MylvwPerson.View = lvwReport
         
    While Not lvwRS.EOF
    ' Navn
        If lvwRS.Fields("Navn") <> "" Then
            Set itmX = lvwPersonale.ListItems.Add(, , lvwRS.Fields("Navn"))
        Else
            Set itmX = lvwPersonale.ListItems.Add(, , " ")
        End If
    
    ' Stilling
        If lvwRS.Fields("StNavn") <> "" Then
            itmX.SubItems(1) = lvwRS.Fields("StNavn")
        Else
            itmX.SubItems(1) = ""
        End If
    
    ' Stillingsandel
        If Not IsNull(lvwRS.Fields("Andel")) And IsNumeric(lvwRS.Fields("Andel")) Then
            itmX.SubItems(2) = CStr(lvwRS.Fields("Andel")) & " %"
        Else
            itmX.SubItems(2) = lvwRS.Fields("Andel")
        End If
    
    ' Merknad
        If lvwRS.Fields("Merk") <> "" Then
            itmX.SubItems(3) = lvwRS.Fields("Merk")
        Else
            itmX.SubItems(3) = ""
        End If
    ' ID
        If lvwRS.Fields("LarerID") <> "" And Not IsNull(lvwRS.Fields("LarerID")) Then
            itmX.SubItems(4) = CStr(lvwRS.Fields("LarerID"))
        Else
            itmX.SubItems(4) = ""
        End If
        lvwRS.MoveNext
    Wend
    lvwRS.Close
    myDb.Close
    'Me.txtBeskrivelse = ""
    'Me.txtEmne = ""
    'Me.txtStp = ""
    'Me.txtUndervisning = ""
    'Me.cboSemester = ""
    stbStatus.Panels(1) = "Antall ansatte: " & MylvwPerson.ListItems.Count

End Sub


Private Sub frEmner_Click()
    Call FyllEmneliste("", "")
    Me.cboEmnegruppe = ""
End Sub

Private Sub frPersoner_Click()
    Call FyllPersonListe("")
    Me.cboStilling = ""
End Sub

Private Sub lvwAlleEmner_ColumnClick(ByVal ColumnHeader As Object)
    Me.lvwAlleEmner.SortOrder = 1 - Me.lvwAlleEmner.SortOrder
    Me.lvwAlleEmner.SortKey = ColumnHeader.Index - 1
    Me.lvwAlleEmner.Sorted = True
End Sub
Private Sub lvwAlleEmner_Click()
    Dim MylvwEmne As ListView
    Dim strEmne As String, strL As String
    Set MylvwEmne = lvwAlleEmner.Object
    If MylvwEmne.ListItems.Count > 0 Then
        With MylvwEmne.SelectedItem
            strEmne = .Text & " " & .SubItems(1)
            Me.txtEmne = strEmne
            Me.txtBelastning = .SubItems(2)
            Me.txtTid = .SubItems(3)
            Me.txtSted = .SubItems(4)
            Me.txtKursID = .SubItems(5)
            Me.txtMerknader = .SubItems(6)
            Me.chkFerdig.Value = .SubItems(7)
        End With
        GL_EID = CLng(Me.txtKursID)
        Call UnlockTxtBoxes
        Me.btnUpdate.Caption = strInsert
        Call VisForelesere(GL_EID, strL)
        MylvwEmne.SelectedItem.ToolTipText = strL
    End If
End Sub

Private Sub lvwKurs_Click()
    Dim MylvwEmne As ListView
    Dim strEmne As String
    Set MylvwEmne = lvwKurs.Object
    If MylvwEmne.ListItems.Count > 0 Then
        With MylvwEmne.SelectedItem
            strEmne = .Text & " " & .SubItems(1)
            Me.txtEmne = strEmne
            Me.txtBelastning = .SubItems(2)
            Me.txtTid = .SubItems(3)
            Me.txtSted = .SubItems(4)
            Me.txtKursID = .SubItems(5)
            Me.txtMerknader = .SubItems(6)
        End With
        GL_EID = CLng(Me.txtKursID)
        Call UnlockTxtBoxes
        Me.btnUpdate.Caption = strUpdate
    End If
End Sub


Private Sub lvwPersonale_Click()
    Dim MylvwPersonale As ListView
    Dim PID As Long
    Dim strLarer As String
    Set MylvwPersonale = lvwPersonale.Object
    If MylvwPersonale.ListItems.Count > 0 Then
        With MylvwPersonale.SelectedItem
            strLarer = .Text
            PID = CLng(.SubItems(4))
        End With
        Me.txtLarer = SnuddNavn(strLarer) & " underviser følgende emner:"
        Call ForeleserSineKurs(PID)
        GL_FID = PID
    End If
    Me.btnUpdate.Caption = strInsert
End Sub

Public Sub ForeleserSineKurs(id As Long)
    Dim lvwRS As DAO.Recordset
    Dim rsKurs As DAO.Recordset
    Dim itmX As ListItem
    Dim MylvwKurs As ListView
    Dim sqlStr As String
    Dim NoOfStp As Single, NoOfPStp As Single
    Dim stbStatus As CustomControl
    Set stbStatus = Me.stbStatusLine
    Set myDb = CurrentDb()
    'MsgBox "Studium: " & GL_SID
    sqlStr = "SELECT tblEmne.*, tblLarerEmne.* " & _
            "FROM tblEmne INNER JOIN tblLarerEmne ON tblEmne.EmneID = tblLarerEmne.EmneID " & _
            "WHERE (tblLarerEmne.LarerID= " & id & " AND tblEmne.Aktiv = True);"

    Set lvwRS = myDb.OpenRecordset(sqlStr, dbOpenDynaset)
    ' åpner kurstabell og sletter eventuelle tidligere kurslinjer
    Set rsKurs = myDb.OpenRecordset("tblAplan")
    If rsKurs.RecordCount > 0 Then
        Do While Not rsKurs.EOF
            rsKurs.Delete
            rsKurs.MoveNext
        Loop
    End If
    Set MylvwKurs = Me.lvwKurs.Object
    MylvwKurs.ListItems.Clear
    MylvwKurs.View = lvwReport
    NoOfStp = 0
    NoOfPStp = 0

    While Not lvwRS.EOF
    ' Kode
        If lvwRS.Fields("Emnekode") <> "" Then
            If lvwRS.Fields("Ferdig") = True Then
                Set itmX = MylvwKurs.ListItems.Add(, , lvwRS.Fields("Emnekode"), , 8)
            Else
                Set itmX = MylvwKurs.ListItems.Add(, , lvwRS.Fields("Emnekode"))
            End If
        Else
            Set itmX = MylvwKurs.ListItems.Add(, , " ")
        End If
    
    ' Betegnelse
        If lvwRS.Fields("Emnenavn") <> "" Then
            itmX.SubItems(1) = lvwRS.Fields("Emnenavn")
        Else
            itmX.SubItems(1) = ""
        End If
    ' Studiepoeng
        If lvwRS.Fields("tblLarerEmne.Studiepoeng") <> "" And Not IsNull(lvwRS.Fields("tblLarerEmne.Studiepoeng")) Then
            itmX.SubItems(2) = Format(lvwRS.Fields("tblLarerEmne.Studiepoeng"), "#0.0")
            NoOfPStp = NoOfPStp + lvwRS.Fields("tblLarerEmne.Studiepoeng")
            If lvwRS.Fields("Ferdig") = True Then
                NoOfStp = NoOfStp + lvwRS.Fields("tblLarerEmne.Studiepoeng")
            End If
        Else
            itmX.SubItems(2) = ""
        End If
    ' Semester
        If lvwRS.Fields("Semester") <> "" Then
            itmX.SubItems(3) = lvwRS.Fields("Semester")
        Else
            itmX.SubItems(3) = ""
        End If
    ' Undervisningssted
        If lvwRS.Fields("Sted") <> "" Then
            itmX.SubItems(4) = lvwRS.Fields("Sted")
        Else
            itmX.SubItems(4) = "Molde"
        End If
    ' EmneID
        If lvwRS.Fields("tblEmne.EmneID") <> "" And Not IsNull(lvwRS.Fields("tblEmne.EmneID")) Then
            itmX.SubItems(5) = CStr(lvwRS.Fields("tblEmne.EmneID"))
        Else
            itmX.SubItems(5) = ""
        End If
    ' Merknader
        If lvwRS.Fields("Kommentar") <> "" And Not IsNull(lvwRS.Fields("Kommentar")) Then
            itmX.SubItems(6) = lvwRS.Fields("Kommentar")
        Else
            itmX.SubItems(6) = ""
        End If
    ' legger data i ny linje i kurstabell
        rsKurs.AddNew
            rsKurs!Kurslinje = CStr(itmX) & " " & CStr(itmX.SubItems(1))
            If lvwRS.Fields("Sted") <> "" And lvwRS.Fields("Sted") <> "Molde" Then
                rsKurs!Kurslinje = rsKurs!Kurslinje & " (" & lvwRS.Fields("Sted") & ")"
            End If
            rsKurs!StpSem = CStr(itmX.SubItems(2)) & " (" & CStr(itmX.SubItems(3)) & ")"
        rsKurs.Update
        
        lvwRS.MoveNext
    Wend
    If NoOfStp = NoOfPStp Then
        stbStatus.Panels(2) = "Sum studiepoeng: " & Format(NoOfStp, "##0.0")
    Else
        stbStatus.Panels(2) = "Sum studiepoeng: " & Format(NoOfStp, "##0.0") & " " & Format(NoOfPStp, "(##0.0)")
    End If
    lvwRS.Close
    myDb.Close

End Sub

Private Sub btnAddForeleser_Click()
On Error GoTo Err_btnAddForeleser_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frmAddForeleser"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_btnAddForeleser_Click:
    Exit Sub

Err_btnAddForeleser_Click:
    MsgBox Err.Description
    Resume Exit_btnAddForeleser_Click
    
End Sub

Public Sub UnlockTxtBoxes()
    Me.txtBelastning.Enabled = True
    Me.txtBelastning.Locked = False
    Me.txtBelastning.BackColor = clrWhite
    Me.txtMerknader.Enabled = True
    Me.txtMerknader.Locked = False
    Me.txtMerknader.BackColor = clrWhite
    Me.btnUpdate.Caption = strInsert
End Sub
Private Sub lvwKurs_ColumnClick(ByVal ColumnHeader As Object)
    Me.lvwKurs.SortOrder = 1 - Me.lvwKurs.SortOrder
    Me.lvwKurs.SortKey = ColumnHeader.Index - 1
    Me.lvwKurs.Sorted = True
End Sub
Private Sub lvwPersonale_ColumnClick(ByVal ColumnHeader As Object)
    Me.lvwPersonale.SortOrder = 1 - Me.lvwPersonale.SortOrder
    Me.lvwPersonale.SortKey = ColumnHeader.Index - 1
    Me.lvwPersonale.Sorted = True
End Sub


Public Sub VisForelesere(id As Long, strListe As String)
    Dim myDb As DAO.Database
    Dim lvwRS As DAO.Recordset
    Dim sqlStr As String
    Dim NoOfEmne As Integer
    Dim strLarer As String
    'Dim stbStatusLine As CustomControl
    Set myDb = CurrentDb()
    
    sqlStr = "SELECT tblLarer.Navn, tblLarer.LarerID, tblLarerEmne.* " & _
            "FROM tblLarer INNER JOIN tblLarerEmne ON tblLarer.LarerID = tblLarerEmne.LarerID " & _
            "WHERE (tblLarerEmne.EmneID= " & id & ");"
    
    Set lvwRS = myDb.OpenRecordset(sqlStr, dbOpenDynaset)
    
    strListe = ""
    While Not lvwRS.EOF
        strLarer = Trim(lvwRS.Fields("Navn")) & ": " & Trim(lvwRS.Fields("Studiepoeng")) & " stp  "
        strListe = strListe & strLarer
        lvwRS.MoveNext
    Wend
    lvwRS.Close
    myDb.Close

End Sub


Public Function Fnavn(strNavn As String) As String
    Dim Fornavn As String, Etternavn As String
    Dim lenN As Integer, lenF As Integer, lenE As Integer
    lenN = Len(strNavn)     'lenght of name
    lenE = InStr(1, strNavn, ",") - 1
    lenF = lenN - lenE - 1
    Etternavn = Mid(strNavn, 1, lenE)
    Fornavn = Mid(strNavn, lenE + 2)
    Fnavn = Fornavn
End Function
