Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Compatibility

Module DataEnvironment_DeScenarios_Module
	Friend DeScenarios As DataEnvironment_DeScenarios = New DataEnvironment_DeScenarios()
End Module

Friend Class DataEnvironment_DeScenarios
	Inherits VB6.BaseDataEnvironment
	Public WithEvents ConSummary As ADODB.Connection
	Public WithEvents rsScenarios_Comparision As ADODB.Recordset
	Private m_Scenarios_Comparision As ADODB.Command
	Public Sub New()
		MyBase.New()

		ConSummary = New ADODB.Connection()
		ConSummary.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Project_Parameters.mdb;Persist Security Info=False;"
		m_Connections.Add(ConSummary, "ConSummary")
		m_Scenarios_Comparision = New ADODB.Command()
		rsScenarios_Comparision = New ADODB.Recordset()
		m_Scenarios_Comparision.Name = "Scenarios_Comparision"
		m_Scenarios_Comparision.CommandText = "`Scenarios_Comparision`"
		m_Scenarios_Comparision.CommandType = ADODB.CommandTypeEnum.adCmdTable
		rsScenarios_Comparision.CursorLocation = ADOR.CursorLocationEnum.adUseClient
		rsScenarios_Comparision.CursorType = ADOR.CursorTypeEnum.adOpenStatic
		rsScenarios_Comparision.LockType = ADOR.LockTypeEnum.adLockReadOnly
		rsScenarios_Comparision.Source = m_Scenarios_Comparision
		m_Commands.Add(m_Scenarios_Comparision, "Scenarios_Comparision")
		m_Recordsets.Add(rsScenarios_Comparision, "Scenarios_Comparision")
	End Sub
	Public Sub Scenarios_Comparision()
		If ConSummary.State = ADODB.ObjectStateEnum.adStateClosed Then
			ConSummary.Open()
		End If
		If rsScenarios_Comparision.State = ADODB.ObjectStateEnum.adStateOpen Then
			rsScenarios_Comparision.Close()
		End If
		m_Scenarios_Comparision.ActiveConnection = ConSummary
		rsScenarios_Comparision.Open()
	End Sub
End Class