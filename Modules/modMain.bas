Attribute VB_Name = "modMain"
Public c As New ADODB.Connection
Public r As New ADODB.Recordset
Public rSearch As New ADODB.Recordset
Public rCheck As New ADODB.Recordset
Public s As String
Public search As String


Public Sub Connect()
s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\AddressBook.mdb;Persist Security Info=False;Jet OLEDB:Database Password=hamzas007;"
c.Open s
r.Open "Addresses", c, adOpenDynamic, adLockOptimistic
End Sub

Public Sub CenterForm(frmTemp As Form)
    On Error Resume Next
    frmTemp.Left = (Screen.Width - frmTemp.Width) / 2
    frmTemp.Top = (Screen.Height - frmTemp.Height) / 2 - 500
    Exit Sub
End Sub

Public Sub Main()
frmLogon.Show
End Sub
