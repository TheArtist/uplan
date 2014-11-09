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
End
Begin Joins
    LeftTable ="tblLarer"
    RightTable ="tblLarerEmne"
    Expression ="tblLarer.LarerID=tblLarerEmne.LarerID"
    Flag =3
    LeftTable ="tblLarerEmne"
    RightTable ="tblEmne"
    Expression ="tblLarerEmne.EmneID=tblEmne.EmneID"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x84f845f70b002a45a59e2e9a482c745c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblLarerEmne.Studiepoeng"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Emnekode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Emnenavn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Studiepoeng"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Semester"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Sted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Aktiv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Ferdig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarer.Navn"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =18
    Top =61
    Right =1310
    Bottom =419
    Left =-1
    Top =-1
    Right =1268
    Bottom =163
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
