Operation =1
Option =0
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
    0xeeb314ff3b312045a716be3111f77ca3
End
Begin
End
Begin
    State =0
    Left =18
    Top =8
    Right =1310
    Bottom =364
    Left =-1
    Top =-1
    Right =1285
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =136
        Top =7
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
