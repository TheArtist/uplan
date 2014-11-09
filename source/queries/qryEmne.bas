Operation =1
Option =0
Begin InputTables
    Name ="tblEmne"
End
Begin OutputColumns
    Expression ="tblEmne.Emnekode"
    Expression ="tblEmne.Emnenavn"
    Expression ="tblEmne.Aktiv"
    Alias ="FirstOfSemester"
    Expression ="First(tblEmne.Semester)"
End
Begin OrderBy
    Expression ="tblEmne.Emnekode"
    Flag =0
End
Begin Groups
    Expression ="tblEmne.Emnekode"
    GroupLevel =0
    Expression ="tblEmne.Emnenavn"
    GroupLevel =0
    Expression ="tblEmne.Aktiv"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x917e3c6145b45b488e4580de9b836bc8
End
Begin
    Begin
        dbText "Name" ="tblEmne.Emnekode"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FirstOfSemester"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =18
    Top =8
    Right =1310
    Bottom =643
    Left =-1
    Top =-1
    Right =1273
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =123
        Top =3
        Name ="tblEmne"
        Name =""
    End
End
