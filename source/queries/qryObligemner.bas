Operation =1
Option =0
Begin InputTables
    Name ="tblEmne"
    Name ="tblObligemne"
End
Begin OutputColumns
    Expression ="tblEmne.EmneID"
    Expression ="tblEmne.Emnekode"
    Expression ="tblEmne.Emnenavn"
    Expression ="tblEmne.Aktiv"
    Expression ="tblObligemne.StudieID"
    Expression ="tblObligemne.Emnekode"
End
Begin Joins
    LeftTable ="tblEmne"
    RightTable ="tblObligemne"
    Expression ="tblEmne.Emnekode=tblObligemne.Emnekode"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6124d0627aa58046ae92d64c8be77f54
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblEmne.EmneID"
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
        dbText "Name" ="tblObligemne.StudieID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblObligemne.Emnekode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =18
    Top =40
    Right =1218
    Bottom =360
    Left =-1
    Top =-1
    Right =1176
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =109
        Top =7
        Right =205
        Bottom =121
        Top =0
        Name ="tblEmne"
        Name =""
    End
    Begin
        Left =293
        Top =9
        Right =389
        Bottom =104
        Top =0
        Name ="tblObligemne"
        Name =""
    End
End
