'
' Created by SharpDevelop.
' User: cwg
' Date: 03/08/2011
' Time: 10:49 AM
' 
' To change this template use Tools | Options | Coding | Edit Standard Headers.
Option Strict Off
Imports System.IO

Public Module Variables
    'Public mw As MapWindow.Interfaces.IMapWin
    Public Create_Files_Form As Create_Files
    Public Wait_Form As Wait
    Public SelectFolder_Form As SelectFolder
    Public Out_Files_Form As Out_Files
    'Public SWT_Files_Form As SWT_Files
    'Public Maintenance_Form As Maintenance
    Public frmEvalues_Form As frmEvalues
    Public frmEvaluesYear_Form As frmEvaluesYear
    Public esriMap_Form As esriMap
    Public MeasuredData_Form As MeasuredData
    Public Measured_Predicted_EValues_Form As Measured_Predicted_EValues
    'Public frmOpFiles_Form As frmOpFiles
    Public SelectFile_Form As SelectFile
    Public MeasuredValues_Form As MeasuredValues
    Public Edit_Subbasins_Inputs_Form As Edit_Subbasins_Inputs
    Public frmSelectSubBasins_Form As frmSelectSubBasins
    Public ResultsFEMfrm_Form As ResultsFEMfrm
    Public Land_exclude_Frm_Form As Land_exclude_Frm
    'Public Apex_Files1_Form As Apex_Files1
    'Public Crops1_Form As Crops1
    'Public Till1_Form As Till1
    'Public Operations_Parameters_Form As Operations_Parameters
    Public Initial1_Form As Initial1
    Public FilesToAddFrm_Form As FilesToAddFrm
    Public Resultsfrm_Form As Resultsfrm
    'Public ResultsTotal_Form As ResultsTotal
    'Public ModLog_Form As ModLog
    Public General_Parametrs_Form As General_Parametrs
    'Public frmCtrA_Form As frmCtrA
    Public MeasuredCompare_Form As MeasuredCompare
    'Public Scenarios_Graph_Form As Scenarios_Graph
    Public ScenarioGraph_Form As ScenarioGraph
    Public List_Form As List
    'Public ResultsYear_Form As ResultsYear
    'Public MeasuredCompare_Form As MeasuredCompare
    'Public Scenarios_Graph_Form As Scens_Graph
    Public Scenarios_List_Form As Scenarios_List
    Public ScenariosList_Form As ScenariosList
    Public Scenarios_FEM_Form As Scenarios_FEM
    Public Scenarios_Summary_Form As Scenarios_Summary
    Public NewScenario_Form As NewScenario
    Public NewScenario1_Form As NewScenario1
    Public Subbasins_Included_Frm_Form As Subbasins_Included_Frm
    Public Animals_in_Stream_Form As Animals_in_Stream
    Public SWATGeneral_Form As SWATGeneral
    Public Operations_Form As Operations
    Public txtFiles_Form As txtFiles
    Public Edit_APEX_inputs_Form As Edit_APEX_Inputs
    Public Edit_APEX_inputs1_Form As Edit_APEX_Inputs1
    Public BMPsParameters_Form As BMPsParameters
    Public APEX_Oper_Form As APEX_Oper
    Public APEX_Soil_Form As APEX_Soil
    Public FileList_Form As FileList
    Public APEX_Subarea_Form As APEX_Subarea
    Public APEX_Site_Form As APEX_Site
    Public APEX_Parm_Form As APEX_Parm
    Public Apex_Cont_Form As Apex_Cont
    Public SWAT_Routing_Form As SWAT_Routing
    Public SWAT_Ponds_Form As SWAT_Ponds
    Public SWAT_Water_Form As SWAT_Water
    Public SWAT_Soil_Form As SWAT_Soil
    Public SWAT_HRUs_Form As SWAT_HRUs
    Public SWAT_GroundWater_Form As SWAT_GroundWater
    Public SWAT_WaterUse_Form As SWAT_WaterUse
    Public SWAT_Chemical_Form As SWAT_Chemical
    Public SWAT_Oper_Form As SWAT_Oper
    Public SWAT_Subarea_Form As SWAT_Subarea
    Public index(150, 50) As Single
    Public deletedRows(150, 10) As String
    Public recCnt As Short
    Public qpSF, qnSF, ynSF, ypSF As Single 'Yang 2/19/2013 
    Public selectedSubbasin As String 'Yang 2/19/2013       
    Public Const ac_to_km2 As Single = 0.00404686
    Public Const ft_to_km As Single = 0.0003048
    Public Const ac_to_ha As Single = 0.4046863
    Public Const ft_to_m As Single = 0.3048
    Public Const in_to_mm As Single = 25.4
    Public Const ha_to_km2 As Single = 0.01
    Public Const km2_to_ha As Single = 100
    Public Const lbs_to_kg As Single = 0.453592
    Public Const in_to_m As Single = 0.0254

    Public Function IsAvailable(ByRef form As System.Windows.Forms.Form) As Boolean
        Return (Not IsNothing(form)) AndAlso Not form.IsDisposed
    End Function

    Public Sub Close_Stream_File(ByVal a As Object)
        If Not a Is Nothing Then
            a.Close()
            a.Dispose()
            a = Nothing
        End If
    End Sub

    Public Function Get_land_use(opcs_number As String) As String
        Dim srOpcs As StreamReader = New StreamReader(Path.Combine(Initial.Dir2 + "\APEX", Initial.Opcs))
        Dim temp() As String
        Dim srOpcsFile As StreamReader = Nothing

        temp = srOpcs.ReadLine.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        Do While opcs_number <> temp(0)
            temp = srOpcs.ReadLine.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        Loop

        srOpcsFile = New StreamReader(Path.Combine(Initial.Dir2 + "\APEX", temp(1)))
        temp(0) = srOpcsFile.ReadLine()
        Dim pos As UShort = temp(0).IndexOf("Luse:")
        Return temp(0).Substring(pos + 5, 4).Trim + ", " + temp(1)
    End Function

    Public Sub CreateTable(ByRef tableName As String, path As String)
        Dim cn As ADODB.Connection
        Dim Cat As ADOX.Catalog
        Dim objTable As Object
        Dim i As Short
        Dim tempiMonth, tempfMonth As Object
        Dim tempSite As Boolean


        Try
            cn = New ADODB.Connection
            Cat = New ADOX.Catalog
            objTable = New ADOX.Table
            tempiMonth = False
            tempfMonth = False : tempSite = False
            'Open the connection
            If Initial.Output_files = "" Then
                MsgBox("You need to open a project first - Open your project and try again")
                Exit Sub
            End If
            cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & "\local.mdb")

            'Open the Catalog
            Cat.ActiveConnection = cn

            'Create the table
            For i = 0 To CShort(Cat.Tables.Count - 1)
                If Cat.Tables(i).Name = tableName Then Exit Sub
            Next
            objTable.name = tableName
            'Create and Append a new field to the tablename Columns Collection
            Select Case tableName
                Case "BMPs"
                    objTable.Columns.Append("Scenario", ADOR.DataTypeEnum.adLongVarChar)
                    objTable.Columns.Append("BMP_Name", ADOR.DataTypeEnum.adLongVarChar)
                    objTable.Columns.Append("BMP", ADOR.DataTypeEnum.adLongVarChar)
                    objTable.Columns.Append("Other", ADOR.DataTypeEnum.adLongVarChar)
                    objTable.Columns.Append("Input1", ADOR.DataTypeEnum.adSingle)
                    objTable.Columns.Append("Input2", ADOR.DataTypeEnum.adSingle)
                    objTable.Columns.Append("Input3", ADOR.DataTypeEnum.adSingle)
                    objTable.Columns.Append("Input4", ADOR.DataTypeEnum.adSingle)
                Case "BMP_Subareas"
                    objTable.Columns.Append("BMP", ADOR.DataTypeEnum.adBSTR)
                    objTable.Columns.Append("Subarea", ADOR.DataTypeEnum.adBSTR)
                    objTable.Columns.Append("HRU", ADOR.DataTypeEnum.adBSTR)
                    objTable.Columns.Append("Land_Use", ADOR.DataTypeEnum.adBSTR)
                    objTable.Columns.Append("Soil", ADOR.DataTypeEnum.adBSTR)
            End Select
            'Create and Append a new key. Note that we are merely passing
            'the "PimaryKey_Field" column as the source of the primary key. This
            'new Key will be Appended to the Keys Collection of "Test_Table"
            'objTable.Keys.Append "PrimaryKey", adKeyPrimary, "PrimaryKey_Field"
            'Append the newly created table to the Tables Collection
            Cat.Tables.Append(objTable)
            ' clean up objects
            objTable = Nothing
            Cat = Nothing
            cn.Close()
            cn = Nothing

            Exit Sub

        Catch ex As Exception
            MsgBox(ex.Message & "-- program Update --> CreateTable")
        End Try
    End Sub
End Module
