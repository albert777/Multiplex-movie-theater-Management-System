Attribute VB_Name = "registration"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public Sub connectdb()
If con.State = 1 Then con.Close
con.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=ilahia_vb_Multiplex;Data Source=LENOVO-PC"
con.Open
End Sub
Public Function insert(dataArray() As Variant, ByVal sqlTable As String) As String
On Error GoTo ErrHandler
Dim errMsg As String
Dim qry As String
qry = "insert into " & sqlTable & " values('"
qry = qry + Join(dataArray, "','")
qry = qry & "')"

con.Execute (qry)

insert = "Data Added Successfully completed"
Exit Function

ErrHandler:
errMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
'GoLogTheError smsg

insert = errMsg
End Function
Public Function getId(ByVal txt As String) As Integer
Dim i As Integer
Dim id As Integer
i = 0
For i = 1 To Len(txt)
    If Mid(txt, i, 1) = "," Then
        id = Mid(txt, 1, i - 1)
        Exit For
    End If
Next i
getId = id
End Function

Public Function bindComboBox(ByVal combo As ComboBox, ByVal sqlTxt As String) As Integer
If rs.State Then rs.Close

rs.Open sqlTxt, con, adOpenDynamic, adLockOptimistic

combo.clear

If rs.BOF = False Then
    rs.MoveFirst
    combo.Text = rs.Fields(0) & ", " & rs.Fields(1)
    While Not rs.EOF
        combo.AddItem (rs.Fields(0) & ", " & rs.Fields(1))
        rs.MoveNext
    Wend
End If
End Function
Public Function chkDataExistency(ByVal dbTable As String, ByVal dbField As String, ByVal data As String) As Integer
If rs.State Then rs.Close
Dim s As String
s = "select count(*) from " & dbTable & " where " & dbField & "='" & data & "'"
rs.Open s, con
chkDataExistency = rs(0)
End Function



