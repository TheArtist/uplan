Operation =1
Option =0
Begin InputTables
    Name ="tblLarer"
    Name ="tblLarerEmne"
End
Begin OutputColumns
    Expression ="tblLarer.LarerID"
    Expression ="tblLarer.Navn"
    Expression ="tblLarer.FagID"
    Expression ="tblLarer.Andel"
    Expression ="tblLarerEmne.EmneID"
    Expression ="tblLarerEmne.Studiepoeng"
End
Begin Joins
    LeftTable ="tblLarer"
    RightTable ="tblLarerEmne"
    Expression ="tblLarer.LarerID = tblLarerEmne.LarerID"
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
    0x67f619b7f4d47c4aa2efec7fa3eb9353
End
Begin
    Begin
        dbText "Name" ="tblLarerEmne.Studiepoeng"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarer.Navn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarer.Andel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarerEmne.EmneID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarer.LarerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLarer.FagID"
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
    Bottom =161
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =360
        Top =11
        Right =480
        Bottom =140
        Top =0
        Name ="tblLarer"
        Name =""
    End
    Begin
        Left =115
        Top =11
        Right =235
        Bottom =142
        Top =0
        Name ="tblLarerEmne"
        Name =""
    End
End
