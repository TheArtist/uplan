Operation =1
Option =0
Begin InputTables
    Name ="tblParameter"
    Name ="tblLarer"
    Name ="tblStilling"
    Name ="tblLarerEmne"
End
Begin OutputColumns
    Expression ="tblLarer.Navn"
    Expression ="tblLarer.Stkode"
    Expression ="tblLarer.Andel"
    Expression ="tblLarer.Merk"
    Expression ="tblStilling.StKode"
    Expression ="tblStilling.StNavn"
    Expression ="tblLarerEmne.Studiepoeng"
    Expression ="tblParameter.studyYear"
End
Begin Joins
    LeftTable ="tblLarer"
    RightTable ="tblStilling"
    Expression ="tblLarer.Stkode = tblStilling.StKode"
    Flag =1
    LeftTable ="tblLarerEmne"
    RightTable ="tblLarer"
    Expression ="tblLarerEmne.LarerID = tblLarer.LarerID"
    Flag =1
End
Begin OrderBy
    Expression ="tblLarer.Navn"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9196785a30ea594f95920b19587035ec
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblLarer.Navn"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x09f8d830ad3eea4e83597181b087e78b
        End
    End
    Begin
        dbText "Name" ="tblLarer.Stkode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x515b8eeb391646468f600f428374f49d
        End
    End
    Begin
        dbText "Name" ="tblLarer.Andel"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x40b8866f345a0346a96749b82a902703
        End
    End
    Begin
        dbText "Name" ="tblLarer.Merk"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf5caa68b35392b4ab6f34159a6d3ab91
        End
    End
    Begin
        dbText "Name" ="tblStilling.StKode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x41eecf73b16f724ca4a4c81ba5934068
        End
    End
    Begin
        dbText "Name" ="tblStilling.StNavn"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x091670686579444ea8ae704312a3ffad
        End
    End
    Begin
        dbText "Name" ="tblParameter.studyYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1dd327a539eb9d4fb721b6a5a2bc802f
        End
    End
    Begin
        dbText "Name" ="tblLarerEmne.Studiepoeng"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =18
    Top =8
    Right =1404
    Bottom =436
    Left =-1
    Top =-1
    Right =1354
    Bottom =129
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =618
        Top =9
        Right =738
        Bottom =149
        Top =0
        Name ="tblParameter"
        Name =""
    End
    Begin
        Left =229
        Top =10
        Right =349
        Bottom =150
        Top =0
        Name ="tblLarer"
        Name =""
    End
    Begin
        Left =416
        Top =11
        Right =575
        Bottom =113
        Top =0
        Name ="tblStilling"
        Name =""
    End
    Begin
        Left =23
        Top =10
        Right =167
        Bottom =154
        Top =0
        Name ="tblLarerEmne"
        Name =""
    End
End
