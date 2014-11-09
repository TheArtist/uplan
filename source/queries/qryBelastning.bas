Operation =1
Option =0
Where ="(((tblEmne.Aktiv)=Yes) AND ((tblLarer.Andel)>0))"
Begin InputTables
    Name ="tblLarer"
    Name ="tblLarerEmne"
    Name ="tblEmne"
End
Begin OutputColumns
    Expression ="tblEmne.Emnekode"
    Expression ="tblEmne.Emnenavn"
    Expression ="tblEmne.Semester"
    Expression ="tblEmne.Sted"
    Expression ="tblEmne.Aktiv"
    Expression ="tblLarerEmne.Studiepoeng"
    Expression ="tblLarer.FagID"
    Expression ="tblLarer.Navn"
    Expression ="tblLarer.Andel"
End
Begin Joins
    LeftTable ="tblLarer"
    RightTable ="tblLarerEmne"
    Expression ="tblLarer.LarerID = tblLarerEmne.LarerID"
    Flag =2
    LeftTable ="tblLarerEmne"
    RightTable ="tblEmne"
    Expression ="tblLarerEmne.EmneID = tblEmne.EmneID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x100230493965214bbd92e9978598543d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblEmne.Emnekode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarer.FagID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Emnenavn"
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
        dbText "Name" ="tblLarerEmne.Studiepoeng"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarer.Navn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblParameter.studyYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarer.Andel"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =2
    Top =8
    Right =1294
    Bottom =484
    Left =-1
    Top =-1
    Right =1260
    Bottom =181
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =433
        Top =14
        Right =553
        Bottom =143
        Top =0
        Name ="tblLarer"
        Name =""
    End
    Begin
        Left =248
        Top =9
        Right =368
        Bottom =123
        Top =0
        Name ="tblLarerEmne"
        Name =""
    End
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =136
        Top =0
        Name ="tblEmne"
        Name =""
    End
End
