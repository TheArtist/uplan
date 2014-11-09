Operation =1
Option =0
Where ="(((tblEmne.Aktiv)=True))"
Begin InputTables
    Name ="tblEmne"
    Name ="tblParameter"
End
Begin OutputColumns
    Expression ="tblEmne.Emnekode"
    Expression ="tblEmne.Emnenavn"
    Expression ="tblEmne.Studiepoeng"
    Expression ="tblEmne.Semester"
    Expression ="tblEmne.Sted"
    Expression ="tblEmne.Comment"
    Expression ="tblParameter.studyYear"
    Expression ="tblEmne.Aktiv"
End
Begin OrderBy
    Expression ="tblEmne.Emnekode"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf337b1a357c83a4c81134878d7cd06a1
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
        dbText "Name" ="tblEmne.Comment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblParameter.studyYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Aktiv"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =1
    Top =2
    Right =1293
    Bottom =358
    Left =-1
    Top =-1
    Right =1268
    Bottom =146
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =136
        Top =0
        Name ="tblEmne"
        Name =""
    End
    Begin
        Left =216
        Top =7
        Right =336
        Bottom =136
        Top =0
        Name ="tblParameter"
        Name =""
    End
End
