Option Strict Off
Option Explicit On
Module Scenarios_Module
'	Dim con As ADODB.Connection
' Dim con1 As ADODB.Connection
'	Dim con2 As ADODB.Connection
'	Dim con3 As ADODB.Connection
'	Dim rec3 As ADODB.Recordset
'	Dim rec As ADODB.Recordset
'	Dim rec1 As ADODB.Recordset
'	Dim rec2 As ADODB.Recordset
'	Dim Conv As Convertion
'	Dim FirstTime As Byte
'	Dim Total_Cost_ha_Watershed As Object
'	Dim NumOfSce As Short
'	Dim Label3(12) As Double
'	Dim Label3_(4) As Double

'Public Sub Main_Program()

'    con3 = New ADODB.Connection
'    rec3 = New ADODB.Recordset

'    Wait_Form.Label1(3).Text = "Cost for Scenarios is being Calculated"
'    Wait_Form.Label1(4).Text = "Please Wait"
'    Wait_Form.Show()

'    con3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir1 & "\Project_Parameters.mdb"
'    con3.Open()

'    Call List1_Click()

'    Wait_Form.Close()
'    If NumOfSce > 10 Then NumOfSce = 10

'End Sub

'Private Sub Graph_Click()
'    If Not IsAvailable(Scenarios_Graph_Form) Then Scenarios_Graph_Form = New Scenarios_Graph()
'    Scenarios_Graph_Form.ShowDialog()
'End Sub

'Private Sub Form_Unload(ByRef cancel As Short)
'    On Error Resume Next
'    If rec.State = 1 Then rec.Close()
'    If rec1.State = 1 Then rec1.Close()
'    If rec2.State = 1 Then rec2.Close()
'    If con.State = 1 Then con.Close()
'    If con1.State = 1 Then con1.Close()
'    If con2.State = 1 Then con2.Close()
'End Sub

'Private Sub List1_Click()
'Dim conect1 As ADODB.Connection
'Dim record1 As ADODB.Recordset

'conect1 = New ADODB.Connection
'record1 = New ADODB.Recordset
'rec3.Open("DELETE * FROM Scenarios_Comparision", con3)
'rec3.Open("SELECT * FROM Scenarios_Comparision", con3, ADOR.CursorTypeEnum.adOpenDynamic, ADOR.LockTypeEnum.adLockOptimistic)
'Dir_Bas = "\APEX"
'Initial.Dir_Bas = Initial.Dir1 & Dir_Bas
'Initial.Dir_BasF = Initial.Dir1 & "\FEM"
'Initial.Bas = "Baseline"

'conect1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir1 & "\Project_Parameters.mdb"
'conect1.Open()

'With record1
'    .Open("SELECT Scenario, Folder, APEX, FEM FROM paths", conect1)
'    NumOfSce = 0

'    Do While .EOF <> True
'        If .Fields("Scenario").Value <> "Baseline" Then
'            Dir_Sce = .Fields("APEX").Value
'            Initial.Dir_SceF = .Fields("Folder").Value & .Fields("FEM").Value
'            Initial.Dir_Sce = .Fields("Folder").Value & Dir_Sce
'            Initial.Sce = .Fields("Scenario").Value
'            Call Calc_Per()
'            NumOfSce = NumOfSce + 1
'            If NumOfSce > 1 Then
'            End If
'        End If
'        record1.MoveNext()
'    Loop

'    .Close()
'    conect1.Close()
'End With
'End Sub

'Private Sub Calc_Per()
'    Dim Highest_Area As Single
'    con = New ADODB.Connection
'    con1 = New ADODB.Connection
'    rec = New ADODB.Recordset
'    rec1 = New ADODB.Recordset
'    con2 = New ADODB.Connection
'    rec2 = New ADODB.Recordset
'    Conv = New Convertion

'    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir1 & "\Project_Parameters.mdb"

'    con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir_Bas & "\Local.mdb"
'    con2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir_Sce & "\Local.mdb"
'    con1.Open()
'    con.Open()
'    rec.Open("DELETE * FROM Reach_Combined", con)
'    con2.Open()

'    con.CursorLocation = ADOR.CursorLocationEnum.adUseClient

'    rec.Open("SELECT * FROM Reach_Combined", con, ADOR.CursorTypeEnum.adOpenDynamic, ADOR.LockTypeEnum.adLockOptimistic)
'    rec2.Open("SELECT * FROM Reach_Total", con2)
'    rec1.Open("SELECT * FROM Reach_Total", con1)
'    Highest_Area = 0

'    With rec
'        Do While rec1.EOF <> True
'            .AddNew()
'            .Fields("RCH").Value = rec1.Fields("RCH").Value

'            .Fields("Flow_Out").Value = VB6.Format(-1 * (1 - rec2.Fields("Flow_Out").Value / rec1.Fields("Flow_Out").Value) * 100, "##0.00")
'            .Fields("Flow_In").Value = VB6.Format(rec2.Fields("Flow_Out").Value - rec1.Fields("Flow_Out").Value, "##0.00")

'            If rec1.Fields("Sed_Out").Value = 0 Then
'                .Fields("Sed_Out").Value = VB6.Format(0, "##0.00")
'                .Fields("Sed_In").Value = VB6.Format(0, "##0.00")
'            Else
'                .Fields("Sed_Out").Value = VB6.Format(-1 * (1 - rec2.Fields("Sed_Out").Value / rec1.Fields("Sed_Out").Value) * 100, "##0.00")
'                .Fields("Sed_In").Value = VB6.Format(rec2.Fields("Sed_Out").Value - rec1.Fields("Sed_Out").Value, "##0.00")
'            End If

'            If rec1.Fields("OrgN_Out").Value = 0 Then
'                .Fields("OrgN_Out").Value = VB6.Format(0, "##0.00")
'                .Fields("OrgN_In").Value = VB6.Format(0, "##0.00")
'            Else
'                .Fields("OrgN_Out").Value = VB6.Format(-1 * (1 - rec2.Fields("OrgN_Out").Value / rec1.Fields("OrgN_Out").Value) * 100, "##0.00")
'                .Fields("OrgN_In").Value = VB6.Format(rec2.Fields("OrgN_Out").Value - rec1.Fields("OrgN_Out").Value, "##0.00")
'            End If

'            .Fields("OrgN_In").Value = VB6.Format(rec2.Fields("OrgN_Out").Value - rec1.Fields("OrgN_Out").Value, "##0.00")
'            .Fields("OrgP_Out").Value = VB6.Format(-1 * (1 - rec2.Fields("OrgP_Out").Value / rec1.Fields("OrgP_Out").Value) * 100, "##0.00")
'            .Fields("OrgP_In").Value = VB6.Format(rec2.Fields("OrgP_Out").Value - rec1.Fields("OrgP_Out").Value, "##0.00")
'            .Fields("NO3_Out").Value = VB6.Format(-1 * (1 - rec2.Fields("NO3_Out").Value / rec1.Fields("NO3_Out").Value) * 100, "##0.00")
'            .Fields("NO3_In").Value = VB6.Format(rec2.Fields("NO3_Out").Value - rec1.Fields("NO3_Out").Value, "##0.00")
'            .Fields("MinP_Out").Value = VB6.Format(-1 * (1 - rec2.Fields("MinP_Out").Value / rec1.Fields("MinP_Out").Value) * 100, "##0.00")
'            .Fields("MinP_In").Value = VB6.Format(rec2.Fields("MinP_Out").Value - rec1.Fields("MinP_Out").Value, "##0.00")
'            If rec1.Fields("Area").Value > Highest_Area Then
'                Highest_Area = rec1.Fields("Area").Value
'                Label3(0) = CDbl(VB6.Format(.Fields("Sed_In").Value / (Highest_Area * 100), "######0.00"))
'                Label3(1) = CDbl(VB6.Format(.Fields("OrgN_In").Value / (Highest_Area * 100), "######0.00"))
'                Label3(2) = CDbl(VB6.Format(.Fields("OrgP_In").Value / (Highest_Area * 100), "######0.00"))
'                Label3(3) = CDbl(VB6.Format(.Fields("NO3_In").Value / (Highest_Area * 100), "######0.00"))
'                Label3(4) = CDbl(VB6.Format(.Fields("MinP_In").Value / (Highest_Area * 100), "######0.00"))

'                Label3_(0) = rec2.Fields("Sed_Out").Value / rec1.Fields("Sed_Out").Value * 1 - 1
'                Label3_(1) = rec2.Fields("OrgN_Out").Value / rec1.Fields("OrgN_Out").Value * 1 - 1
'                Label3_(2) = rec2.Fields("OrgP_Out").Value / rec1.Fields("OrgP_Out").Value * 1 - 1
'                Label3_(3) = rec2.Fields("NO3_Out").Value / rec1.Fields("NO3_Out").Value * 1 - 1
'                Label3_(4) = rec2.Fields("MinP_Out").Value / rec1.Fields("MinP_Out").Value * 1 - 1
'            End If
'            .Fields("Area").Value = rec1.Fields("Area").Value
'            .Update()
'            rec1.MoveNext()
'            rec2.MoveNext()
'        Loop
'        .Close()
'    End With

'    rec.Open("SELECT RCH,Flow_Out,Sed_Out,OrgN_Out,OrgP_Out,NO3_Out,MinP_Out FROM Reach_Combined", con, ADOR.CursorTypeEnum.adOpenDynamic, ADOR.LockTypeEnum.adLockOptimistic)

'    rec2.Close()
'    rec1.Close()

'    Call Calc_Cost()

'    Label3(5) = CDbl(VB6.Format(Total_Cost_ha_Watershed, "#####0.00000"))
'    If Label3(0) = 0 Then
'        Label3(12) = 0
'    Else
'        Label3(12) = CDbl(VB6.Format(Val(CStr(Label3(5))) / Val(CStr(Label3(0))), "#####0.00000"))
'    End If
'    If Label3(1) = 0 Then
'        Label3(11) = 0
'    Else
'        Label3(11) = CDbl(VB6.Format(Val(CStr(Label3(5))) / Val(CStr(Label3(1))), "#####0.00000"))
'    End If
'    If Label3(2) = 0 Then
'        Label3(10) = 0
'    Else
'        Label3(10) = CDbl(VB6.Format(Val(CStr(Label3(5))) / Val(CStr(Label3(2))), "#####0.00000"))
'    End If
'    If Label3(3) = 0 Then
'        Label3(9) = 0
'    Else
'        Label3(9) = CDbl(VB6.Format(Val(CStr(Label3(5))) / Val(CStr(Label3(3))), "#####0.00000"))
'    End If
'    If Label3(4) = 0 Then
'        Label3(8) = 0
'    Else
'        Label3(8) = CDbl(VB6.Format(Val(CStr(Label3(5))) / Val(CStr(Label3(4))), "#####0.00000"))
'    End If

'    Call Update_Summary()

'End Sub

'Private Sub Calc_Cost()
'    Dim Total_Cost_ha_Sce As Object
'    Dim Cost_ha2 As Object
'    Dim Total_Cost_ha_Bas As Object
'    Dim Cost_ha1 As Object
'    Dim Total_Factor2 As Object
'    Dim Total_Cost_ha2 As Object
'    Dim Total_Factor1 As Object
'    Dim Total_Cost_ha1 As Object
'    Dim cons1 As ADODB.Connection
'    Dim cons2 As ADODB.Connection
'    Dim recs1 As ADODB.Recordset
'    Dim recs2 As ADODB.Recordset

'    cons1 = New ADODB.Connection
'    cons2 = New ADODB.Connection
'    recs1 = New ADODB.Recordset
'    recs2 = New ADODB.Recordset

'    Total_Cost_ha1 = 0
'    Total_Factor1 = 0
'    Total_Cost_ha2 = 0
'    Total_Factor2 = 0

'    cons1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir_BasF & "\SWAPPFEMOut.mdb"
'    cons1.Open()

'    With recs1
'        .Open("SELECT * FROM [FEM Output Summary]", cons1)
'        Do While recs1.EOF <> True
'            Cost_ha1 = .Fields("Net Returns").Value / .Fields("Total Hectares").Value
'            Total_Cost_ha1 = Total_Cost_ha1 + (Cost_ha1 * .Fields("Weight").Value)
'            Total_Factor1 = Total_Factor1 + .Fields("Weight").Value
'            .MoveNext()
'        Loop
'        Total_Cost_ha_Bas = Total_Cost_ha1 / Total_Factor1
'        .Close()
'    End With
'    cons1.Close()

'    cons2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir_SceF & "\SWAPPFEMOut.mdb"
'    cons2.Open()

'    With recs2
'        .Open("SELECT * FROM [FEM Output Summary]", cons2)
'        Do While recs2.EOF <> True
'            Cost_ha2 = .Fields("Net Returns").Value / .Fields("Total Hectares").Value
'            Total_Cost_ha2 = Total_Cost_ha2 + (Cost_ha2 * .Fields("Weight").Value)
'            Total_Factor2 = Total_Factor2 + .Fields("Weight").Value
'            .MoveNext()
'        Loop
'        Total_Cost_ha_Sce = Total_Cost_ha2 / Total_Factor2
'        .Close()
'    End With
'    cons2.Close()

'    Total_Cost_ha_Watershed = Total_Cost_ha_Bas - Total_Cost_ha_Sce

'End Sub

'Private Sub Update_Summary()

'With rec3
'    .AddNew()
'    .Fields("Scenario").Value = Initial.Sce
'    .Fields("Sediment_Units").Value = Label3(0)
'    .Fields("OrgN_Units").Value = Label3(1)
'    .Fields("OrgP_Units").Value = Label3(2)
'    .Fields("NO3_Units").Value = Label3(3)
'    .Fields("PO4_Units").Value = Label3(4)
'    .Fields("Cost_Units").Value = Label3(5)
'    .Fields("Sediment_$").Value = Label3(12)
'    .Fields("OrgN_$").Value = Label3(11)
'    .Fields("OrgP_$").Value = Label3(10)
'    .Fields("NO3_$").Value = Label3(9)
'    .Fields("PO4_$").Value = Label3(8)

'    .Fields("Sediment_%").Value = Label3_(0)
'    .Fields("OrgN_%").Value = Label3_(1)
'    .Fields("OrgP_%").Value = Label3_(2)
'    .Fields("NO3_%").Value = Label3_(3)
'    .Fields("PO4_%").Value = Label3_(4)
'    .Update()
'    .Requery()
'End With


'End Sub
End Module