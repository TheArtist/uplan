Operation =1
Option =0
Where ="(((tblEmne.Emnekode)=[tblObligemne].[Emnekode]))"
Begin InputTables
    Name ="tblObligemne"
    Name ="tblEmne"
End
Begin OutputColumns
    Expression ="tblObligemne.StudieID"
    Expression ="tblObligemne.Emnekode"
    Expression ="tblEmne.Emnekode"
    Expression ="tblEmne.Emnenavn"
End
Begin Joins
    LeftTable ="tblObligemne"
    RightTable ="tblEmne"
    Expression ="tblObligemne.Emnekode=tblEmne.Emnekode"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9cc1ba8f9f2d5a4d89f470bc886f2d01
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblObligemne.StudieID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblObligemne.Emnekode"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1049
    Bottom =601
    Left =-1
    Top =-1
    Right =1025
    Bottom =282
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblObligemne"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tblEmne"
        Name =""
    End
End
