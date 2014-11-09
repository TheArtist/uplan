Operation =1
Option =0
Where ="(((tblEmne.Aktiv)=True))"
Begin InputTables
    Name ="qryBelastning00"
    Name ="tblEmne"
End
Begin OutputColumns
    Expression ="qryBelastning00.LarerID"
    Expression ="qryBelastning00.Navn"
    Expression ="qryBelastning00.Andel"
    Expression ="qryBelastning00.Studiepoeng"
    Alias ="Stp"
    Expression ="[qryBelastning00].[Studiepoeng]*Abs(CInt([Ferdig]))"
    Expression ="tblEmne.Emnekode"
    Expression ="tblEmne.Emnenavn"
    Expression ="qryBelastning00.FagID"
    Expression ="tblEmne.Aktiv"
    Expression ="tblEmne.Ferdig"
    Expression ="tblEmne.Sted"
    Expression ="tblEmne.Semester"
End
Begin Joins
    LeftTable ="qryBelastning00"
    RightTable ="tblEmne"
    Expression ="qryBelastning00.EmneID = tblEmne.EmneID"
    Flag =2
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
    0x55bc95930069b64cb20305d1505afb43
End
Begin
    Begin
        dbText "Name" ="qryBelastning00.LarerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBelastning00.Navn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBelastning00.Andel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBelastning00.Studiepoeng"
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
        dbText "Name" ="tblEmne.Aktiv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Ferdig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Sted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblEmne.Semester"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBelastning00.FagID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stp"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x035aea5f15b6304e9072bcf7d294ce54
        End
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
    Bottom =191
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =76
        Top =6
        Right =220
        Bottom =150
        Top =0
        Name ="qryBelastning00"
        Name =""
    End
    Begin
        Left =377
        Top =1
        Right =538
        Bottom =166
        Top =0
        Name ="tblEmne"
        Name =""
    End
End
