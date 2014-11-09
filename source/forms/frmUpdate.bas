Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    DatasheetFontHeight =10
    ItemSuffix =1
    Left =1560
    Top =24
    Right =9012
    Bottom =3756
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4e2925685846e340
    End
    DatasheetFontName ="Arial"
    Begin
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin Section
            Height =1701
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1133
                    Top =510
                    Width =945
                    Height =405
                    Name ="btnUpdate"
                    Caption ="Oppdater"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnUpdate_Click()
On Error GoTo Err_btnUpdate_Click

    Dim myDb As DAO.Database
    Dim rsLarer As DAO.Recordset
    Dim strKode As String
    Dim sNavn As String
    Dim sEpost As String
    Dim ant As Integer
    Set myDb = CurrentDb()
    Set rsLarer = myDb.OpenRecordset("tblLarer", dbOpenDynaset)
    ant = 0
    Do While Not rsLarer.EOF
        If rsLarer!Andel <= 20 Then 'II'erstillinger og timelærere tas ikke med
            'sNavn = rsLarer!Navn
            'sEpost = LagEpostNavn(sNavn)
            rsLarer.Edit
                rsLarer!Epost = ""
            rsLarer.Update
            ant = ant + 1
        End If
        rsLarer.MoveNext
    Loop
    rsLarer.Close
    MsgBox "Oppdatering ferdig. " & ant & " epostadresser oppdatert."

Exit_btnUpdate_Click:
    Exit Sub

Err_btnUpdate_Click:
    MsgBox Err.Description
    Resume Exit_btnUpdate_Click
    
End Sub
Public Function LagEpostNavn(strNavn) As String
    Dim Fnavn As String, Enavn As String
    Const eAdresse = "@himolde.no"
    Dim lenN As Integer, lenF As Integer, lenE As Integer
    lenN = Len(strNavn)     'lenght of name
    lenE = InStr(1, strNavn, ",") - 1
    lenF = lenN - lenE - 1
    Enavn = Mid(strNavn, 1, lenE)
    Fnavn = Mid(strNavn, lenE + 2)
    LagEpostNavn = Trim(Fnavn & "." & Enavn & eAdresse)
End Function
