Operation =1
Option =0
Where ="(((tblEmne.Studiepoeng)<30) AND ((tblEmne.Aktiv)=True) AND ((tblEmne.Ferdig)=Fal"
    "se))"
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
    Expression ="tblEmne.Ferdig"
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
    0x2f6f0a076a1597438406627045a16800
End
Begin
End
Begin
    State =0
    Left =18
    Top =9
    Right =1310
    Bottom =365
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
        Bottom =121
        Top =7
        Name ="tblEmne"
        Name =""
    End
    Begin
        Left =216
        Top =7
        Right =336
        Bottom =121
        Top =0
        Name ="tblParameter"
        Name =""
    End
End
