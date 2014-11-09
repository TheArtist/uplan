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
    Width =13492
    DatasheetFontHeight =10
    ItemSuffix =122
    Left =1620
    Top =180
    Right =15795
    Bottom =8565
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x00529b40d4c5e240
    End
    Caption ="Faglærere"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
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
            Height =6633
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin OptionGroup
                    OverlapFlags =93
                    Left =226
                    Top =443
                    Width =13093
                    Height =5611
                    ColumnOrder =7
                    Name ="frPersoner"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =345
                            Top =345
                            Width =2145
                            Height =285
                            FontSize =9
                            Name ="Label17"
                            Caption ="Rediger faglæreroversikt"
                            FontName ="Tahoma"
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =8503
                            Top =680
                            Width =1021
                            Height =170
                            OptionValue =1
                            Name ="chkAllePersoner"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =8730
                                    Top =623
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
                    Left =5526
                    Top =6179
                    Width =2502
                    Height =345
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
                    Left =10826
                    Top =6179
                    Width =2502
                    Height =345
                    FontSize =9
                    TabIndex =2
                    Name ="btnClose"
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
                    Left =2876
                    Top =6179
                    Width =2502
                    Height =345
                    FontSize =9
                    TabIndex =3
                    Name ="btnUpdate"
                    Caption ="Oppdater"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CustomControl
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =342
                    Top =1025
                    Width =12804
                    Height =4536
                    AutoActivate =1
                    TabIndex =4
                    Name ="lvwPersonale"
                    OleData = Begin
                        0x00160000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
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
                        0xf0283628000000000000000000000000e030285c3baac9010700000080020000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000d4010000 ,
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
                        0xf0283628000000000000000000000000e0e74a8718aac9010500000080020000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000d3010000 ,
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
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000004bf0d1bd8b85d111b16a00c0f0283628214334120800000043580000 ,
                        0x361f00004e087deb010006001c00000000000000000000000086010043580000 ,
                        0x01efcdab0000050078fe9c0e07002a0008000080050000805050eb0d00000000 ,
                        0x00000000000000001fdeecbd01000500c94feb0d0352e30b918fce119de300aa ,
                        0x004bb851010000009001dc7c010005417269616c080020000000000000003a11 ,
                        0x000005000000000000000000000000000000050000004e61766e002000010000 ,
                        0x000000290f000009000000000000000000000000000000090000005374696c6c ,
                        0x696e67002000020002000000e406000006000000000000000000000000000000 ,
                        0x06000000416e64656c002000030002000000e406000008000000000000000000 ,
                        0x000000000000080000004f766572203630002000040000000000f31200000800 ,
                        0x0000000000000000000000000000080000004d65726b6e616400200005000000 ,
                        0x0000000000000900000000000000000000000000000009000000506572736f6e ,
                        0x4944002001000000feffffff0300000004000000050000000600000007000000 ,
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
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000004bf0d1bd8b85d111b16a00c0f0283628214334120800000043580000 ,
                        0x361f00004e087deb010006001c00000000000000000000000086010043580000 ,
                        0x01efcdab0000050008bb390007002a000800008005000080d02ca91000000000 ,
                        0x00000000000000001fdeecbd01000500492ca9100352e30b918fce119de300aa ,
                        0x004bb851010000009001905f0100065461686f6d61080020000000000000003a ,
                        0x11000005000000000000000000000000000000050000004e61766e0020000100 ,
                        0x00000000290f000009000000000000000000000000000000090000005374696c ,
                        0x6c696e67002000020002000000e4060000060000000000000000000000000000 ,
                        0x0006000000416e64656c002000030002000000e4060000080000000000000000 ,
                        0x00000000000000080000004f766572203630002000040000000000f312000008 ,
                        0x000000000000000000000000000000080000004d65726b6e6164002000050000 ,
                        0x000000000000000900000000000000000000000000000009000000506572736f ,
                        0x6e49440000060000000000000000000e0000000000000000000000000000000e ,
                        0x0000005374696c6c696e67736b6f6465002000070000000000ac140000060000 ,
                        0x000000000000000000000000000600000045706f737400000000000000000000 ,
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
                        0x000000002000060000000000000000000e000000000000000000000000000000 ,
                        0x0e0000005374696c6c696e67736b6f6465002000070000000000ac1400000600 ,
                        0x00000000000000000000000000000600000045706f73740065006d0070007500 ,
                        0x73002000530061006e0073002000490054004300000056006900760061006c00 ,
                        0x640069000000570069006e006700640069006e00670073002000320000005700 ,
                        0x69006e006700640069006e00670073002000330000004a006f006b0065007200 ,
                        0x6d0061006e0000004a007500690063006500200049005400430000004d005300 ,
                        0x20005200650066006500720065006e00630065002000310000004d0053002000 ,
                        0x5200650066006500720065006e00630065002000320000004d00530020005200 ,
                        0x650066006500720065006e00630065002000530061006e007300200053006500 ,
                        0x72006900660000004d00530020005200650066006500720065006e0063006500 ,
                        0x20005300700065006300690061006c007400790000004100670065006e006300 ,
                        0x7900200046004200000041006c00670065007200690061006e00000041007200 ,
                        0x690061006c00200052006f0075006e0064006500640020004d00540020004200 ,
                        0x6f006c006400000041007200690061006c00200055006e00690063006f006400 ,
                        0x650020004d0053000000400041007200690061006c00200055006e0069006300 ,
                        0x6f00640000000000
                    End
                    OLEClass ="ListViewCtrl"
                    Class ="MSComctlLib.ListViewCtrl.2"

                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =11029
                    Top =623
                    Width =2117
                    Height =284
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =5
                    Name ="cboStilling"
                    RowSourceType ="Table/Query"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =10028
                            Top =623
                            Width =912
                            Height =228
                            FontSize =9
                            Name ="Label79"
                            Caption ="Vis bare :"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =340
                    Top =5669
                    Width =2553
                    Height =283
                    FontSize =9
                    TabIndex =6
                    Name ="txtNavn"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =7111
                    Top =5669
                    Width =2712
                    Height =283
                    FontSize =9
                    TabIndex =7
                    Name ="txtMerknad"
                    FontName ="Tahoma"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =226
                    Top =6179
                    Width =2502
                    Height =345
                    FontSize =9
                    TabIndex =8
                    Name ="btnNew"
                    Caption ="Ny"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2932
                    Top =5669
                    Width =2109
                    Height =283
                    FontSize =9
                    TabIndex =9
                    Name ="cboStilling2"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6107
                    Top =5669
                    Width =965
                    Height =283
                    FontSize =9
                    TabIndex =10
                    Name ="cboOver60"
                    RowSourceType ="Value List"
                    RowSource ="\"Yes\";\"No\""
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5104
                    Top =5669
                    Width =964
                    Height =283
                    FontSize =9
                    TabIndex =11
                    Name ="cboAndel"
                    RowSourceType ="Value List"
                    RowSource ="\"0\";\"10\";\"20\";\"30\";\"40\";\"50\";\"60\";\"70\";\"80\";\"90\";\"100\""
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =9862
                    Top =5669
                    Width =3282
                    Height =283
                    FontSize =9
                    TabIndex =12
                    Name ="txtEpost"
                    FontName ="Tahoma"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8176
                    Top =6179
                    Width =2502
                    Height =345
                    FontSize =9
                    TabIndex =13
                    Name ="btnOutlook"
                    Caption ="Hent epostadr  fra Outlook"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =396
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CustomControl
                    Enabled = NotDefault
                    TabStop = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =226
                    Top =68
                    Width =13095
                    Height =270
                    AutoActivate =1
                    Name ="stbStatusline"
                    OleData = Begin
                        0x00220000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000000000000000000000100000 ,
                        0x0a00000001000000feffffff0000000001000000ffffffffffffffffffffffff ,
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
                        0xf0283628000000000000000000000000d04110963baac9010b000000c0080000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000020000001d080000 ,
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
                        0xf028362800000000000000000000000080b9741c4debc8010500000080080000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000fc070000 ,
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
                        0x2000000021000000feffffffffffffffffffffffffffffffffffffffffffffff ,
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
                        0x0000000000000000000000000c000000000000000c0000000000000000000000 ,
                        0x00000000a367388e8685d111b16a00c0f02836282143341208000000d0590000 ,
                        0xdc010000887ee1e6000006007600000000000500ffff130001efcdab00000500 ,
                        0xf8191b0006007200ffffffffffffffff000000000000000002000000a0040000 ,
                        0x0949000009490000030000000700000041006e00740061006c006c0020000700 ,
                        0x000041006e00740061006c006c002000b5040000ea110000ea11000002000000 ,
                        0x05000000310038003a003100360001000000020000006c740000360700000000 ,
                        0x01000300202002000000000030010000360000002020080000000000e8020000 ,
                        0x660100002020100000000000e80200004e040000280000002000000040000000 ,
                        0x0100010000000000000100000000000000000000000000000000000000000000 ,
                        0xffffff0000000000000000003ffffff83ffffff83000001837ffffd837feffd8 ,
                        0x37ffffd837ffffd8377fffd837bfffd837dfffd837efffd837f6ffd837faffd8 ,
                        0x37fcffd835f8035837fe7fd837febfd837feffd837feffd837feffd837feffd8 ,
                        0x37feffd837ffffd837feffd837ffffd8300000183ffffff83ffffff800000000 ,
                        0x00000000ffffffffc00000078000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x80000003800000038000000380000003800000038000000380000003c0000007 ,
                        0xffffffff28000000200000004000000001000400000000008002000000000000 ,
                        0x0000000000000000000000000000000000008000008000000080800080000000 ,
                        0x800080008080000080808000c0c0c0000000ff0000ff000000ffff00ff000000 ,
                        0xff00ff00ffff0000ffffff000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000099999999999999999999999999900000999999 ,
                        0x9999999999999999999990000099000000000000000000000009900000990fff ,
                        0xffffffffffffffffff09900000990ffffffffff0ffffffffff09900000990fff ,
                        0xffffffffffffffffff09900000990fffffffffffffffffffff09900000990fff ,
                        0x0fffffffffffffffff09900000990ffff0ffffffffffffffff09900000990fff ,
                        0xff0fffffffffffffff09900000990ffffff0ffffffffffffff09900000990fff ,
                        0xffff0ff0ffffffffff09900000990ffffffff0f0ffffffffff09900000990fff ,
                        0xffffff00ffffffffff09900000990f0ffffff000000000ff0f09900000990fff ,
                        0xfffffff00fffffffff09900000990ffffffffff0f0ffffffff09900000990fff ,
                        0xfffffff0ffffffffff09900000990ffffffffff0ffffffffff09900000990fff ,
                        0xfffffff0ffffffffff09900000990ffffffffff0ffffffffff09900000990fff ,
                        0xfffffff0ffffffffff09900000990fffffffffffffffffffff09900000990fff ,
                        0xfffffff0ffffffffff09900000990fffffffffffffffffffff09900000990000 ,
                        0x0000000000000000000990000099999999999999999999999999900000999999 ,
                        0x9999999999999999999990000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000ffffffffc0000007800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x80000003c0000007ffffffff2800000020000000400000000100040000000000 ,
                        0x8002000000000000000000000000000000000000000000000000800000800000 ,
                        0x0080800080000000800080008080000080808000c0c0c0000000ff0000ff0000 ,
                        0x00ffff00ff000000ff00ff00ffff0000ffffff00000077777777777777777777 ,
                        0x7777777000077777777777777777777777777777000000000000000000000000 ,
                        0x0000077700111111111111111111111111111077009999999999999999999999 ,
                        0x999910770091000000000000000000000009107700910fffffffffffffffffff ,
                        0xff09107700910ffffffffff0ffffffffff09107700910fffffffffffffffffff ,
                        0xff09107700910fffffffffffffffffffff09107700910fff0fffffffffffffff ,
                        0xff09107700910ffff0ffffffffffffffff09107700910fffff0fffffffffffff ,
                        0xff09107700910ffffff0ffffffffffffff09107700910fffffff0ff0ffffffff ,
                        0xff09107700910ffffffff0f0ffffffffff09107700910fffffffff00ffffffff ,
                        0xff09107700910f0ffffff000000000ff0f09107700910ffffffffff00fffffff ,
                        0xff09107700910ffffffffff0f0ffffffff09107700910ffffffffff0ffffffff ,
                        0xff09107700910ffffffffff0ffffffffff09107700910ffffffffff0ffffffff ,
                        0xff09107700910ffffffffff0ffffffffff09107700910ffffffffff0ffffffff ,
                        0xff09107700910fffffffffffffffffffff09107700910ffffffffff0ffffffff ,
                        0xff09107700910fffffffffffffffffffff091077009100000000000000000000 ,
                        0x0009107700911111111111111111111111191070009999999999999999999999 ,
                        0x9999100000000000000000000000000000000000f0000001e0000000c0000000 ,
                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                        0x80000000800000008000000180000003c00000071fdeecbd0100050000000000 ,
                        0x690061006c0020004e006100720072006f007700200053007000650063006900 ,
                        0x61006c00200047003100000041007200690061006c0020004e00610072007200 ,
                        0x6f00770020005300700065006300690061006c00200047003200000054006900 ,
                        0x6d006500730020004e0065007700200052006f006d0061006e00200053007000 ,
                        0x65006300690061006c002000470031000000540069006d006500730020004e00 ,
                        0x65007700200052006f006d0061006e0020005300700065006300690061006c00 ,
                        0x20004700320000005a005700410064006f00620065004600000045006e006700 ,
                        0x7200610076006500720046006f006e0074004500780074007200610073000000 ,
                        0x45006e0067007200610076006500720046006f006e0074005300650074000000 ,
                        0x45006e00670072006100760065007200540069006d006500000045006e006700 ,
                        0x7200610076006500720054006500780074004e0043005300000045006e006700 ,
                        0x7200610076006500720054006500780074004800000045006e00670072006100 ,
                        0x7600650001000000feffffff0300000004000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x200000002100000022000000feffffffffffffffffffffffffffffffffffffff ,
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
                        0x0000000000000000000000000c000000000000000c0000000000000000000000 ,
                        0x00000000a367388e8685d111b16a00c0f028362821433412080000003a5a0000 ,
                        0xdc010000887ee1e6000006007600000000000600ffff130001efcdab00000500 ,
                        0x0000000006007200ffffffffffffffff000000000000000002000000a0040000 ,
                        0x0949000009490000030000000700000041006e00740061006c006c0020000700 ,
                        0x000041006e00740061006c006c002000b5040000ea110000ea11000002000000 ,
                        0x05000000310036003a003400320001000000020000006c740000360700000000 ,
                        0x01000300202002000000000030010000360000002020080000000000e8020000 ,
                        0x660100002020100000000000e80200004e040000280000002000000040000000 ,
                        0x0100010000000000000100000000000000000000000000000000000000000000 ,
                        0xffffff0000000000000000003ffffff83ffffff83000001837ffffd837feffd8 ,
                        0x37ffffd837ffffd8377fffd837bfffd837dfffd837efffd837f6ffd837faffd8 ,
                        0x37fcffd835f8035837fe7fd837febfd837feffd837feffd837feffd837feffd8 ,
                        0x37feffd837ffffd837feffd837ffffd8300000183ffffff83ffffff800000000 ,
                        0x00000000ffffffffc00000078000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x80000003800000038000000380000003800000038000000380000003c0000007 ,
                        0xffffffff28000000200000004000000001000400000000008002000000000000 ,
                        0x0000000000000000000000000000000000008000008000000080800080000000 ,
                        0x800080008080000080808000c0c0c0000000ff0000ff000000ffff00ff000000 ,
                        0xff00ff00ffff0000ffffff000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000099999999999999999999999999900000999999 ,
                        0x9999999999999999999990000099000000000000000000000009900000990fff ,
                        0xffffffffffffffffff09900000990ffffffffff0ffffffffff09900000990fff ,
                        0xffffffffffffffffff09900000990fffffffffffffffffffff09900000990fff ,
                        0x0fffffffffffffffff09900000990ffff0ffffffffffffffff09900000990fff ,
                        0xff0fffffffffffffff09900000990ffffff0ffffffffffffff09900000990fff ,
                        0xffff0ff0ffffffffff09900000990ffffffff0f0ffffffffff09900000990fff ,
                        0xffffff00ffffffffff09900000990f0ffffff000000000ff0f09900000990fff ,
                        0xfffffff00fffffffff09900000990ffffffffff0f0ffffffff09900000990fff ,
                        0xfffffff0ffffffffff09900000990ffffffffff0ffffffffff09900000990fff ,
                        0xfffffff0ffffffffff09900000990ffffffffff0ffffffffff09900000990fff ,
                        0xfffffff0ffffffffff09900000990fffffffffffffffffffff09900000990fff ,
                        0xfffffff0ffffffffff09900000990fffffffffffffffffffff09900000990000 ,
                        0x0000000000000000000990000099999999999999999999999999900000999999 ,
                        0x9999999999999999999990000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000ffffffffc0000007800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x8000000380000003800000038000000380000003800000038000000380000003 ,
                        0x80000003c0000007ffffffff2800000020000000400000000100040000000000 ,
                        0x8002000000000000000000000000000000000000000000000000800000800000 ,
                        0x0080800080000000800080008080000080808000c0c0c0000000ff0000ff0000 ,
                        0x00ffff00ff000000ff00ff00ffff0000ffffff00000077777777777777777777 ,
                        0x7777777000077777777777777777777777777777000000000000000000000000 ,
                        0x0000077700111111111111111111111111111077009999999999999999999999 ,
                        0x999910770091000000000000000000000009107700910fffffffffffffffffff ,
                        0xff09107700910ffffffffff0ffffffffff09107700910fffffffffffffffffff ,
                        0xff09107700910fffffffffffffffffffff09107700910fff0fffffffffffffff ,
                        0xff09107700910ffff0ffffffffffffffff09107700910fffff0fffffffffffff ,
                        0xff09107700910ffffff0ffffffffffffff09107700910fffffff0ff0ffffffff ,
                        0xff09107700910ffffffff0f0ffffffffff09107700910fffffffff00ffffffff ,
                        0xff09107700910f0ffffff000000000ff0f09107700910ffffffffff00fffffff ,
                        0xff09107700910ffffffffff0f0ffffffff09107700910ffffffffff0ffffffff ,
                        0xff09107700910ffffffffff0ffffffffff09107700910ffffffffff0ffffffff ,
                        0xff09107700910ffffffffff0ffffffffff09107700910ffffffffff0ffffffff ,
                        0xff09107700910fffffffffffffffffffff09107700910ffffffffff0ffffffff ,
                        0xff09107700910fffffffffffffffffffff091077009100000000000000000000 ,
                        0x0009107700911111111111111111111111191070009999999999999999999999 ,
                        0x9999100000000000000000000000000000000000f0000001e0000000c0000000 ,
                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                        0x8000000080000000800000008000000080000000800000008000000080000000 ,
                        0x80000000800000008000000180000003c00000071fdeecbd0100050001000000 ,
                        0x0352e30b918fce119de300aa004bb851010000009001905f0100065461686f6d ,
                        0x610045006e00670072006100760065007200540069006d006500000045006e00 ,
                        0x67007200610076006500720054006500780074004800000045006e0067007200 ,
                        0x610076006500720054006500780074004e0043005300000045006e0067007200 ,
                        0x61007600650072005400650078007400540000004a0061007a007a0000004a00 ,
                        0x61007a007a0043006f007200640000004a0061007a007a005000650072006300 ,
                        0x00004a0061007a007a00540065007800740045007800740065006e0064006500 ,
                        0x640000004a0061007a007a005400650078007400000050006500740072007500 ,
                        0x630063006900000053006500760069006c006c0065000000540061006d006200 ,
                        0x750072006f00000041006e00640061006c00650020004d006f006e006f002000 ,
                        0x49005000410000004d00530020005200650066006500720065006e0063006500 ,
                        0x2000530065007200690066000000650020007000000000000000000007000000 ,
                        0x0000000000000000
                    End
                    OLEClass ="SBarCtrl"
                    Class ="MSComctlLib.SBarCtrl.2"

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
Public PID As Long

Private Sub btnClose_Click()
On Error GoTo Err_btnClose_Click

    DoCmd.Close

Exit_btnClose_Click:
    Exit Sub

Err_btnClose_Click:
    MsgBox Err.Description
    Resume Exit_btnClose_Click

End Sub

Private Sub btnDelete_Click()
On Error GoTo Err_btnDelete_Click
    Dim myDb                As DAO.Database
    Dim msg                 As String
    Dim response            As Integer
    Dim sqlStr              As String
    Dim strStilling         As String
    
    If IsNull(Me.cboStilling) Or Me.cboStilling = "" Then
        strStilling = ""
    Else
        strStilling = Me.cboStilling
    End If
    Set myDb = CurrentDb
    If Me.txtNavn = "" Or PID = 0 Then
        msg1 = "Ingenting å slette. Du må velge et navn."
        MsgBox msg1, vbExclamation + vbOKOnly, OIS_Title
        Exit Sub
    Else
        msg1 = "Vil du slette '" & Me.txtNavn & "' fra databasen?"
        If MsgBox(msg1, vbExclamation + vbYesNo, OIS_Title) = vbNo Then Exit Sub
        sqlDelete = "DELETE * FROM tblLarer WHERE LarerID=" & PID
        myDb.Execute (sqlDelete)
    ' sletter lærer-emne forekomster for dette emnet
        sqlDelete = "DELETE * FROM tblLarerEmne WHERE LarerID=" & PID
        myDb.Execute (sqlDelete)
    End If
    Call FyllPersonListe(strStilling)
 
Exit_btnDelete_Click:
    myDb.Close
    Exit Sub

Err_btnDelete_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnDelete_Click

End Sub


Private Sub btnNew_Click()
On Error GoTo Err_btnNew_Click
    Me.txtNavn = ""
    Me.cboStilling2 = ""
    Me.cboAndel = ""
    Me.txtMerknad = ""
    Me.txtEpost = ""
    Me.txtNavn.SetFocus
    Me.btnUpdate.Caption = strInsert
Exit_btnNew_Click:
    Exit Sub

Err_btnNew_Click:
    MsgBox Err.Description
    Resume Exit_btnNew_Click
End Sub

Private Sub btnOutlook_Click()
On Error GoTo Err_btnOutlook_Click
Dim oOutlookApp As Outlook.Application
Set oOutlookApp = New Outlook.Application
    Dim oRecipient As Recipient
    Dim oNameSpace As Namespace
    Dim strName As String
    Set oNameSpace = GetNamespace("MAPI")
    If Me.txtNavn <> "" Then
        strName = SnuddNavn(Me.txtNavn)
        Set oRecipient = oNameSpace.CreateRecipient(strName)
        oRecipient.Resolve
        If oRecipient.Resolved Then
            Me.txtEpost = oRecipient
        Else
            MsgBox "Epostadresse ikke funnet i Outlook", vbInformation, OIS_Title
        End If
    Else
       MsgBox "Ingen faglærer er valgt.", vbInformation, OIS_Title
    End If
    
Exit_btnOutlook_Click:
    Exit Sub

Err_btnOutlook_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnOutlook_Click
End Sub


Private Sub btnUpdate_Click()
On Error GoTo Err_btnUpdate_Click
    Dim myDb                As DAO.Database
    Dim rsLarer             As DAO.Recordset
    Dim sqlLarer            As String
    Dim msg                 As String
    Dim sqlStr              As String
    Dim strStilling         As String
    Dim strStkode           As String
    
    If IsNull(Me.cboStilling) Or Me.cboStilling = "" Then
        strStilling = ""
    Else
        strStilling = Me.cboStilling
    End If
    
    strStkode = StillingsKode(Me.cboStilling2)
    Set myDb = CurrentDb
    Select Case btnUpdate.Caption
        Case strInsert
            sqlStr = "INSERT INTO tblLarer (Navn, Stkode, Andel, Over60, Merk, Epost) " & _
                    "VALUES('" & Me.txtNavn & "', '" & strStkode & "', '" & CInt(Me.cboAndel) & "','" & Me.cboOver60 & "','" & Me.txtMerknad & "','" & Me.txtEpost & "');"
            myDb.Execute (sqlStr)
            Call FyllPersonListe(strStilling)
        Case "Oppdater"
            If Me.txtNavn = "" Or PID = 0 Then
                msg = "Ingenting å oppdatere. Du må velge en faglærer."
                MsgBox msg, vbOKOnly + vbExclamation, OIS_Title
                Exit Sub
            End If
            sqlStr = "UPDATE tblLarer SET Navn ='" & Me.txtNavn & "', " & _
                    "Stkode = '" & strStkode & "', " & _
                    "Andel = " & CInt(Me.cboAndel) & ", " & _
                    "Over60 = " & Me.cboOver60 & ", " & _
                    "Merk = '" & Me.txtMerknad & "', " & _
                    "Epost = '" & Me.txtEpost & "' " & _
               "WHERE LarerID = " & PID & ";"
            myDb.Execute (sqlStr)
            Call FyllPersonListe(strStilling)
    End Select
    Me.txtNavn = ""
    Me.cboStilling2 = ""
    Me.cboAndel = ""
    Me.cboOver60 = ""
    Me.txtMerknad = ""
    Me.txtEpost = ""
    Me.btnUpdate.Caption = "Oppdater"

Exit_btnUpdate_Click:
    Exit Sub

Err_btnUpdate_Click:
    MsgBox Err.Description, , OIS_Title
    Resume Exit_btnUpdate_Click
End Sub



Private Sub cboStilling_Click()
    Dim ValgtStilling As String
    ValgtStilling = Me.cboStilling.Column(0)
    Call FyllPersonListe(ValgtStilling)
    Me.frPersoner = 0
End Sub


Private Sub Form_Load()
On Error GoTo Err_Form_Load
    Dim sqlstr1 As String, sqlstr2 As String
    
    Dim itmX As ListItem
    
    Call FyllPersonListe("")
    Me.cboStilling.RowSource = "SELECT StNavn from tblStilling ORDER BY StNavn"
    Me.cboStilling2.RowSource = "SELECT StNavn from tblStilling ORDER BY StNavn"
    Me.btnUpdate.Caption = "Oppdater"
    Me.frPersoner = 1   ' viser alle ansatte
Exit_Form_Load:
    Exit Sub

Err_Form_Load:
    MsgBox Err.Description
    Resume Exit_Form_Load
End Sub



Public Sub FyllPersonListe(ValgtStilling As String)
    Dim myDb As DAO.Database
    Dim lvwRS As DAO.Recordset
    Dim itmX As ListItem
    Dim MylvwPerson As ListView
    Dim sqlStr As String
    Dim NoOfEmne As Integer
    Dim stbStatus As CustomControl
    Set stbStatus = Me.stbStatusLine
    Set myDb = CurrentDb()
    If ValgtStilling = "" Then
        sqlStr = "SELECT tblLarer.LarerID, tblLarer.Navn, tblStilling.StNavn, tblLarer.Stkode, tblLarer.Andel, tblLarer.Over60, tblLarer.Merk, tblLarer.Epost " & _
            "FROM tblLarer INNER JOIN tblStilling ON tblLarer.Stkode = tblStilling.StKode " & _
            "ORDER BY tblLarer.Navn;"
    Else
        sqlStr = "SELECT tblLarer.LarerID, tblLarer.Navn, tblStilling.StNavn, tblLarer.Stkode, tblLarer.Andel, tblLarer.Over60, tblLarer.Merk, tblLarer.Epost " & _
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
            Set itmX = Me.lvwPersonale.ListItems.Add(, , lvwRS.Fields("Navn"))
        Else
            Set itmX = Me.lvwPersonale.ListItems.Add(, , " ")
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
    
    ' Over 60 år
        If lvwRS.Fields("Over60") <> "" Then
            itmX.SubItems(3) = lvwRS.Fields("Over60")
        Else
            itmX.SubItems(3) = ""
        End If
    ' Merknad
        If lvwRS.Fields("Merk") <> "" Then
            itmX.SubItems(4) = lvwRS.Fields("Merk")
        Else
            itmX.SubItems(4) = ""
        End If
    ' ID
        If lvwRS.Fields("LarerID") <> "" And Not IsNull(lvwRS.Fields("LarerID")) Then
            itmX.SubItems(5) = CStr(lvwRS.Fields("LarerID"))
        Else
            itmX.SubItems(5) = ""
        End If
    ' Stillingskode
        If lvwRS.Fields("StKode") <> "" Then
            itmX.SubItems(6) = lvwRS.Fields("StKode")
        Else
            itmX.SubItems(6) = ""
        End If
    ' Epostadresse
        If lvwRS.Fields("Epost") <> "" Then
            itmX.SubItems(7) = lvwRS.Fields("Epost")
        Else
            itmX.SubItems(7) = ""
        End If

        lvwRS.MoveNext
    Wend
    Me.txtNavn = ""
    Me.txtMerknad = ""
    Me.cboStilling2 = ""
    Me.cboAndel = ""
    Me.cboOver60 = ""
    Me.txtEpost = ""
    stbStatus.Panels(1) = "Antall ansatte: " & MylvwPerson.ListItems.Count
    lvwRS.Close
    myDb.Close
End Sub


Private Sub frPersoner_Click()
    Call FyllPersonListe("")
    Me.cboStilling = ""
End Sub


Private Sub lvwPersonale_Click()
    Dim MylvwPersonale As ListView
    Dim strLarer As String
    Set MylvwPersonale = lvwPersonale.Object
    If MylvwPersonale.ListItems.Count > 0 Then
        With MylvwPersonale.SelectedItem
            Me.txtNavn = .Text
            Me.cboStilling2 = .SubItems(1)
            Me.cboAndel = Val(.SubItems(2))
            Me.cboOver60 = .SubItems(3)
            Me.txtMerknad = .SubItems(4)
            PID = CLng(.SubItems(5))
            Me.txtEpost = .SubItems(7)
        End With
        Me.btnUpdate.Caption = "Oppdater"
    End If
End Sub



Private Sub lvwPersonale_ColumnClick(ByVal ColumnHeader As Object)
    Me.lvwPersonale.SortOrder = 1 - Me.lvwPersonale.SortOrder
    Me.lvwPersonale.SortKey = ColumnHeader.Index - 1
    Me.lvwPersonale.Sorted = True
End Sub


Public Function StillingsKode(strStilling As String) As String
    Dim myDb As DAO.Database
    Dim rsStilling As DAO.Recordset
    Dim sqlStr As String
    Set myDb = CurrentDb()
    sqlStr = "SELECT * FROM tblStilling WHERE StNavn = '" & strStilling & "';"
    Set rsStilling = myDb.OpenRecordset(sqlStr, dbOpenDynaset)
    If rsStilling.RecordCount = 1 Then
        StillingsKode = rsStilling("StKode")
    Else
        StillingsKode = "99"
    End If
End Function
