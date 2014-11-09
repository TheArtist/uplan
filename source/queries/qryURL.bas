Operation =3
Name ="tblURL"
Option =0
Begin InputTables
    Name ="tblEmne"
End
Begin OutputColumns
    Name ="Ekode"
    Expression ="tblEmne.Emnekode"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x682369780095c449ac31d12fb5ab0d02
End
Begin
    Begin
        dbText "Name" ="tblEmne.Emnekode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1054
    Bottom =692
    Left =-1
    Top =-1
    Right =1030
    Bottom =373
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblEmne"
        Name =""
    End
End
