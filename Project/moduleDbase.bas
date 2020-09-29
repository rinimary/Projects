Attribute VB_Name = "Module1"
Global con As New ADODB.Connection
Global rs As New ADODB.Recordset

Global con1 As New ADODB.Connection
Global rs1 As New ADODB.Recordset

Global con2 As New ADODB.Connection
Global rs2 As New ADODB.Recordset

Global con3 As New ADODB.Connection
Global rs3 As New ADODB.Recordset

Global con4 As New ADODB.Connection
Global rs4 As New ADODB.Recordset



Public Sub dbConnect()

Dim fname As String

fname = App.Path & "\dbCluster.mdb"

Set con = New ADODB.Connection

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fname & ";JET OLEDB:Database password=cando; Persist Security Info=False"

Set rs = Nothing
Set rs = New ADODB.Recordset

End Sub

Public Sub dbController()

Dim fname As String

fname = App.Path & "\dbController.mdb"

Set con4 = New ADODB.Connection

con4.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fname & ";JET OLEDB:Database password=cando; Persist Security Info=False"

Set rs4 = Nothing
Set rs4 = New ADODB.Recordset

End Sub

Public Sub dbConnectServer11()

Dim fname As String

fname = App.Path & "\dbServer11.mdb"

Set con1 = New ADODB.Connection

con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fname & ";JET OLEDB:Database password=cando; Persist Security Info=False"

Set rs1 = Nothing
Set rs1 = New ADODB.Recordset

End Sub
Public Sub dbConnectServer21()

Dim fname As String

fname = App.Path & "\dbServer21.mdb"

Set con1 = New ADODB.Connection

con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fname & ";JET OLEDB:Database password=cando; Persist Security Info=False"

Set rs1 = Nothing
Set rs1 = New ADODB.Recordset

End Sub

Public Sub dbConnectServer22()

Dim fname As String

fname = App.Path & "\dbServer22.mdb"

Set con1 = New ADODB.Connection

con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fname & ";JET OLEDB:Database password=cando; Persist Security Info=False"

Set rs1 = Nothing
Set rs1 = New ADODB.Recordset

End Sub
Public Sub dbConnectServer31()

Dim fname As String

fname = App.Path & "\dbServer31.mdb"

Set con1 = New ADODB.Connection

con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fname & ";JET OLEDB:Database password=cando; Persist Security Info=False"

Set rs1 = Nothing
Set rs1 = New ADODB.Recordset

End Sub

Public Sub dbConnectServer32()

Dim fname As String

fname = App.Path & "\dbServer32.mdb"

Set con1 = New ADODB.Connection

con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fname & ";JET OLEDB:Database password=cando; Persist Security Info=False"

Set rs1 = Nothing
Set rs1 = New ADODB.Recordset

End Sub

Public Sub dbConnectServer12()

Dim fname As String

fname = App.Path & "\dbServer12.mdb"

Set con2 = New ADODB.Connection

con2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fname & ";JET OLEDB:Database password=cando; Persist Security Info=False"

Set rs2 = Nothing
Set rs2 = New ADODB.Recordset

End Sub



Public Sub dbConnectServer13()

Dim fname As String

fname = App.Path & "\dbServer13.mdb"

Set con3 = New ADODB.Connection

con3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fname & ";JET OLEDB:Database password=cando; Persist Security Info=False"

Set rs3 = Nothing
Set rs3 = New ADODB.Recordset

End Sub

