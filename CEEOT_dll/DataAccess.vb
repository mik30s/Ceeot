Option Strict Off
Option Explicit On
Module DataAccess
 'Public Function LocalDB(ByRef SQLString As String) As ADODB.Recordset

 '	Dim adoConnection As ADODB.Connection
 '	Dim sc As String

 '	adoConnection = New ADODB.Connection
 '	LocalDB = New ADODB.Recordset
 '	'UPGRADE_WARNING: Couldn't resolve default property of object Initial.Output_files. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
 '	sc = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Initial.Output_files & "\Local.mdb;Persist Security Info=False"
 '	adoConnection.Open(sc)

 '	LocalDB.Open(SQLString, adoConnection,  , ADOR.LockTypeEnum.adLockOptimistic)

 '	adoConnection.Close()

 'End Function

 'Public Function parmDB(ByRef SQLString As String) As ADODB.Recordset

 '	Dim adoConnection As ADODB.Connection
 '	Dim sc As String

 '	adoConnection = New ADODB.Connection
 '	parmDB = New ADODB.Recordset
 '	'UPGRADE_WARNING: Couldn't resolve default property of object Initial.Dir1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
 '	sc = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Initial.Dir1 & "\Project_Parameters.mdb;Persist Security Info=False"
 '	adoConnection.Open(sc)

 '	parmDB.Open(SQLString, adoConnection,  , ADOR.LockTypeEnum.adLockOptimistic)

 'End Function
	
End Module