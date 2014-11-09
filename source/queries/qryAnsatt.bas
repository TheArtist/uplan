Operation =1
Option =0
Begin InputTables
    Name ="tblParameter"
    Name ="tblLarer"
    Name ="tblStilling"
End
Begin OutputColumns
    Expression ="tblLarer.Navn"
    Expression ="tblLarer.Stkode"
    Expression ="tblLarer.Andel"
    Expression ="tblLarer.Merk"
    Expression ="tblStilling.StKode"
    Expression ="tblStilling.StNavn"
    Expression ="tblParameter.studyYear"
End
Begin Joins
    LeftTable ="tblLarer"
    RightTable ="tblStilling"
    Expression ="tblLarer.Stkode=tblStilling.StKode"
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
    0x52b8aea1a8258c4f824a6ae7c32553bb
End
Begin
End
Begin
    State =0
    Left =18
    Top =8
    Right =1531
    Bottom =397
    Left =-1
    Top =-1
    Right =1502
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =423
        Top =7
        Right =543
        Bottom =147
        Top =0
        Name ="tblParameter"
        Name =""
    End
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =147
        Top =0
        Name ="tblLarer"
        Name =""
    End
    Begin
        Left =216
        Top =7
        Right =375
        Bottom =109
        Top =0
        Name ="tblStilling"
        Name =""
    End
End
