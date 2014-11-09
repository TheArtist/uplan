Operation =5
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
End
Begin Joins
    LeftTable ="tblLarer"
    RightTable ="tblLarerEmne"
    Expression ="tblLarer.LarerID = tblLarerEmne.LarerID"
    Flag =1
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
dbBinary "GUID" = Begin
    0x5b4850c8fae32e419ea341075c293a69
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="tblLarerEmne.Studiepoeng"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x32212d42648641478d3b5afb6d453ab9
        End
    End
    Begin
        dbText "Name" ="tblEmne.Emnekode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x04ba95594be7a94ca9237726311d1e24
        End
    End
    Begin
        dbText "Name" ="tblEmne.Emnenavn"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9c81578abdf13c4d9d6ff6f256f31841
        End
    End
    Begin
        dbText "Name" ="tblEmne.Studiepoeng"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00c5e1c56cf8494bb361484080af893e
        End
    End
    Begin
        dbText "Name" ="tblEmne.Semester"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4ec9154e04763b41ad8f1f44acad360c
        End
    End
    Begin
        dbText "Name" ="tblEmne.Sted"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x98a963676ac0844ca99437c105973e7d
        End
    End
    Begin
        dbText "Name" ="tblEmne.Aktiv"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1a9f504b15ec72478b20cbc871fa0b49
        End
    End
    Begin
        dbText "Name" ="tblEmne.Ferdig"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbf73bf45103d5640807778193ccb05d6
        End
    End
    Begin
        dbText "Name" ="tblLarer.Navn"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3c5199bdda197d47bacf4c5a5e9d6b95
        End
    End
    Begin
        dbText "Name" ="tblParameter.studyYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x28da5d18f4c442459e8ce4b9d1eb2693
        End
    End
End
Begin
    State =0
    Left =-6
    Top =-72
    Right =1286
    Bottom =506
    Left =-1
    Top =-1
    Right =1268
    Bottom =163
    Left =0
    Top =0
    ColumnsShown =771
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
