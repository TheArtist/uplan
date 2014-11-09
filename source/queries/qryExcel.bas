Operation =1
Option =0
Where ="(((tblEmne.Aktiv)=Yes))"
Begin InputTables
    Name ="tblLarer"
    Name ="tblLarerEmne"
    Name ="tblEmne"
End
Begin OutputColumns
    Expression ="tblEmne.Emnekode"
    Expression ="tblEmne.Emnenavn"
    Expression ="tblEmne.Studiepoeng"
    Expression ="tblEmne.Semester"
    Expression ="tblEmne.Sted"
    Expression ="tblEmne.Aktiv"
    Expression ="tblEmne.Ferdig"
    Expression ="tblLarerEmne.Studiepoeng"
    Expression ="tblLarer.Navn"
    Expression ="tblEmne.FagID"
    Expression ="tblLarerEmne.Kommentar"
End
Begin Joins
    LeftTable ="tblLarer"
    RightTable ="tblLarerEmne"
    Expression ="tblLarer.LarerID = tblLarerEmne.LarerID"
    Flag =3
    LeftTable ="tblLarerEmne"
    RightTable ="tblEmne"
    Expression ="tblLarerEmne.EmneID = tblEmne.EmneID"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x1caf56a7b8ba044395ab0451cacd9369
End
Begin
    Begin
        dbText "Name" ="tblLarerEmne.Studiepoeng"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3ded9dfedbc8124da069efede5ebbfb6
        End
    End
    Begin
        dbText "Name" ="tblEmne.Emnekode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7dff768e38a2014cb38261e9655ba144
        End
    End
    Begin
        dbText "Name" ="tblEmne.Emnenavn"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xebba73d4bb2a0b408d16044ebddcfb1d
        End
    End
    Begin
        dbText "Name" ="tblEmne.Studiepoeng"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ef09f3de295df49a8426e0e5ea0c117
        End
    End
    Begin
        dbText "Name" ="tblEmne.Semester"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x88b8c184595c1047af1fb767d038d9c1
        End
    End
    Begin
        dbText "Name" ="tblEmne.Sted"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8c237457e930b49b130367a0b9172d0
        End
    End
    Begin
        dbText "Name" ="tblEmne.Aktiv"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2a8684e2a0ee6646bb8b0c2350e74858
        End
    End
    Begin
        dbText "Name" ="tblEmne.Ferdig"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf876eb0586af3745b13092c1ff102400
        End
    End
    Begin
        dbText "Name" ="tblLarer.Navn"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4d93a6919c326f4b8cc16bdb50d94a31
        End
    End
    Begin
        dbText "Name" ="tblEmne.FagID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarerEmne.Kommentar"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =2
    Top =61
    Right =1294
    Bottom =419
    Left =-1
    Top =-1
    Right =1268
    Bottom =151
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =433
        Top =14
        Right =553
        Bottom =130
        Top =0
        Name ="tblLarer"
        Name =""
    End
    Begin
        Left =248
        Top =9
        Right =368
        Bottom =110
        Top =0
        Name ="tblLarerEmne"
        Name =""
    End
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =123
        Top =0
        Name ="tblEmne"
        Name =""
    End
End
