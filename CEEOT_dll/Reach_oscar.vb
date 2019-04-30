Option Strict Off
Option Explicit On

Imports System.Data.OleDb

Module Reach
    'Dim fs As Object
    'Dim a, z As Object
    Dim temp As String
    Dim name1 As String
    Dim rs_Day, rs_Month, rs_Year, rs_Total As DataTable
    'Dim cn As ADODB.Connection
    Dim year_Renamed, MonthNum As Short
    Dim jDay As Short
    Dim totalArea As Single
    Dim dsBas As New Data.DataSet()
    Dim myConnection As OleDb.OleDbConnection
    Dim dbConnectString As String
    Dim command As OleDb.OleDbCommand
    Dim writer As System.IO.StreamWriter

    Private Function Basinsrch() As Short
        Dim year1 As Short = 0
        Dim current_line As Short
        Dim ADORecordset As DataTable
        Dim TakeField As Convertion
        Dim i As Short
        On Error GoTo goError

        TakeField = New Convertion
        ADORecordset = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'Basins.rch' and version = " & "'" & Initial.Version & "'" & "ORDER BY line, field")

        With ADORecordset
            current_line = .Rows(0).Item("Line")

            For i = 0 To .Rows.Count - 1
                If (.Rows(i).Item("SwatFile") <> "") Then
                    TakeField.filename = Initial.Swat_Output & "\" & Trim(.Rows(i).Item("SwatFile"))
                    TakeField.Leng = .Rows(i).Item("Leng")
                    TakeField.LineNum = .Rows(i).Item("Lines")
                    TakeField.Inicia = .Rows(i).Item("Inicia")
                    Basinsrch = TakeField.value()
                Else
                    Basinsrch = .Rows(i).Item("Value")
                End If
                year1 = year1 + Val(Basinsrch)
            Next
        End With

        ADORecordset.Dispose()
        ADORecordset = Nothing

        Basinsrch = year1
        Exit Function
goError:
        MsgBox(Err.Description & " " & name1, , "Function: Basinsrch on Reach Module")

    End Function

    Public Sub ReadFile()
        Dim sc As Object
        Dim temp1(10) As String

        On Error GoTo goError

        'fs = CreateObject("Scripting.FileSystemObject")

        Select Case Initial.Version
            Case "1.0.0", "2.0.0", "4.0.0"
                name1 = Initial.Swat_Output & "\Basins.rch"
            Case "1.1.0", "2.1.0", "2.3.0", "4.1.0", "4.2.0", "4.3.0"
                name1 = Initial.Swat_Output & "\output.rch"
        End Select

        year_Renamed = Basinsrch()

        Exit Sub
goError:
        MsgBox(Err.Description & " " & name1, , "Function: ReadFile in Reach Module")

    End Sub

    Public Sub ReadSaveFile()
        Dim pointPos As Integer
        Dim sc As Object
        Dim temp1(10) As String
        Dim ADORecordset As DataTable
        Dim dt As New DataTable
        Dim query As String
        Dim i, j As Integer

        On Error GoTo goError

        dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & Initial.Output_files & "\Local.mdb;"
        myConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = dbConnectString
        myConnection.Open()
        'fs = CreateObject("Scripting.FileSystemObject")
        rs_Day = New DataTable
        ADORecordset = getDBDataTable("SELECT * FROM Subbasins" & " ORDER BY subbasin")
        year_Renamed = Basinsrch()
        modifyLocalRecords("DELETE * FROM Reach_Total", Initial.Output_files)
        modifyLocalRecords("DELETE * FROM Reach_Year", Initial.Output_files)
        modifyLocalRecords("DELETE * FROM Reach_Month", Initial.Output_files)
        'modifyLocalRecords("DELETE * FROM Reach_Day", Initial.Output_files)

        Wait_Form.Label1(2).Text = "File Being Transfered:"
        'Wait_Form.Show()

        dt = getLocalDataTable("SELECT Subbasin FROM Runs", Initial.Output_files)

        Select Case Initial.Version
            Case "1.0.0", "2.0.0", "4.0.0"
                name1 = Initial.Swat_Output & "\Basins.rch"
            Case "1.1.0", "2.1.0", "2.3.0", "4.1.0", "4.2.0", "4.3.0"
                name1 = Initial.Swat_Output & "\output.rch"
        End Select

        With ADORecordset
            For i = 0 To .Rows.Count - 1
                If dt.Rows.Count <= 0 Then
                    modifyLocalRecords("DELETE * FROM Reach_Day", Initial.Output_files)
                    Call read_SaveDay()
                Else
                    'Call read_SaveDay()
                    For j = 0 To dt.Rows.Count - 1
                        If .Rows(i).Item("Subbasin") = dt.Rows(j).Item("Subbasin") Then
                            modifyLocalRecords("DELETE * FROM Reach_Day WHERE rch = " & .Rows(i).Item("File_Number"), Initial.Output_files)
                            Call read_SaveDay(.Rows(i).Item("File_Number"), .Rows(i).Item("Subbasin"))
                        End If
                    Next
                End If
            Next
        End With

        writer.Close()
        writer.Dispose()
        writer = Nothing
        'Wait_Form.Close()
        ADORecordset.Dispose()
        ADORecordset = Nothing

        'save CVS file to local DB
        'dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & Initial.Output_files & "\Local.mdb;"
        'myConnection = New OleDb.OleDbConnection
        'myConnection.ConnectionString = dbConnectString
        'myConnection.Open()
        'Existing table
        'Dim AccessCommand As New System.Data.OleDb.OleDbCommand _
        '("INSERT INTO [Reach_Day] (Subbasin, rch, [Year], [Mon], Area, Flow_Out, Sed_Out, Orgn_Out, OrgP_Out, NO3_Out, MinP_Out, TotalN, TotalP) SELECT F1, F2, F3, F4, F5,F6, F7, F8, F9, F10,F11, F12, F13 FROM [Text;Database=" & Initial.Swat_Output & ";Hdr=No].[Reach.csv]", myConnection)
        'AccessCommand.ExecuteNonQuery()
        'If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()
        'myConnection.Dispose()
        'myConnection = Nothing
        query = "INSERT INTO [Reach_Day] (Subbasin, rch, [Year], [Mon], Area, Flow_Out, Sed_Out, Orgn_Out, OrgP_Out, NO3_Out, MinP_Out, TotalN, TotalP)"
        query = query & " SELECT F1, F2, F3, F4, F5,F6, F7, F8, F9, F10,F11, F12, F13 "
        query = query & "FROM [Text;Database=" & Initial.Swat_Output & ";Hdr=No].[Reach.csv]"
        modifyLocalRecords(query, Initial.Output_files)

        query = "INSERT INTO Reach_Month (Subbasin, rch, [Year], [Mon], Area, Flow_Out, Sed_Out, Orgn_Out, OrgP_Out, NO3_Out, MinP_Out, TotalN, TotalP)"
        query = query & " SELECT Subbasin, RCH, Year, MON, Avg(Area), Avg(flow_Out), Sum(Sed_Out), Sum(OrgN_Out), Sum(OrgP_Out), Sum(NO3_Out), Sum(MinP_Out), sum(TotalN), sum(TotalP)"
        query = query & " FROM Reach_Day"
        query = query & " GROUP BY Reach_Day.Subbasin, Reach_Day.RCH, Reach_Day.[YEAR], Reach_Day.[MON];"
        modifyLocalRecords(query, Initial.Output_files)

        query = "INSERT INTO Reach_Year (Subbasin, rch, [Year], Area, Flow_Out, Sed_Out, Orgn_Out, OrgP_Out, NO3_Out, MinP_Out, TotalN, TotalP)"
        query = query & " SELECT Subbasin, RCH, [Year], Avg(Area), Avg(flow_Out), Sum(Sed_Out), Sum(OrgN_Out), Sum(OrgP_Out), Sum(NO3_Out), Sum(MinP_Out), sum(TotalN), sum(TotalP)"
        query = query & " FROM Reach_Month"
        query = query & " GROUP BY Reach_Month.Subbasin, Reach_Month.rch, Reach_Month.RCH, Reach_Month.[YEAR] ;"
        modifyLocalRecords(query, Initial.Output_files)

        query = "INSERT INTO Reach_Total (Subbasin, rch, Area, Flow_Out, Sed_Out, Orgn_Out, OrgP_Out, NO3_Out, MinP_Out, TotalN, TotalP)"
        query = query & " SELECT Subbasin, RCH,  Avg(Area), Avg(flow_Out), Avg(Sed_Out), Avg(OrgN_Out), Avg(OrgP_Out), Avg(NO3_Out), Avg(MinP_Out), Avg(TotalN), Avg(TotalP)"
        query = query & " FROM Reach_Year"
        query = query & " GROUP BY Reach_Year.Subbasin, Reach_Year.rch, Reach_Year.RCH;"
        modifyLocalRecords(query, Initial.Output_files)

        Exit Sub
goError:
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()
        myConnection.Dispose()
        myConnection = Nothing

        MsgBox(Err.Description & " " & name1, , "Function: ReadSaveFile in Reach Module")
    End Sub

    '	Private Sub read_Month()
    '		Dim Offset As Object
    '		Dim date1 As Object
    '		Dim i As Object

    '		On Error GoTo goError
    '		z = fs.createtextfile(Initial.Swat_Output & "\Titles.rch")
    '		a = fs.OpenTextFile(name1)
    '		rs_Month = New ADODB.Recordset

    '		rs_Month.Open("Reach_Month", cn, ADOR.CursorTypeEnum.adOpenDynamic, ADOR.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdTable)

    '		Wait_Form.Label1(2).Text = "Number of Monthly Records Transfered:"
    '		Wait_Form.Show()
    '		For i = 1 To 9
    '			z.WriteLine(a.ReadLine)
    '		Next 

    '		z.Close()
    '		temp = a.ReadLine
    '		i = 1

    '		Do While a.AtEndOfStream <> True
    '			date1 = Val(Mid(temp, 20, 6))
    '			Select Case date1
    '				Case Is > 1900
    '					year_Renamed = date1 + 1
    '				Case Else
    '					If Mid(temp, 24, 1) = "." Then
    '					Else
    '						rs_Month.AddNew()
    '						rs_Month.Fields.Item("Reach").Value = Left(temp, 5)
    '						rs_Month.Fields.Item("RCH").Value = Val(Mid(temp, 6, 5))
    '						rs_Month.Fields.Item("year").Value = year_Renamed
    '						rs_Month.Fields.Item("MON").Value = Val(Mid(temp, 20, 6))
    '						Offset = 12
    '						rs_Month.Fields.Item("area").Value = Val(Mid(temp, 26 + (Offset * 0), 12))
    '						rs_Month.Fields.Item("FLOW_Out").Value = Val(Mid(temp, 26 + (Offset * 2), 12))
    '						rs_Month.Fields.Item("Sed_Out").Value = Val(Mid(temp, 26 + (Offset * 6), 12))
    '						rs_Month.Fields.Item("OrgN_Out").Value = Val(Mid(temp, 26 + (Offset * 9), 12))
    '						rs_Month.Fields.Item("OrgP_Out").Value = Val(Mid(temp, 26 + (Offset * 11), 12))
    '						rs_Month.Fields.Item("NO3_Out").Value = Val(Mid(temp, 26 + (Offset * 13), 12))
    '						rs_Month.Fields.Item("MinP_Out").Value = Val(Mid(temp, 26 + (Offset * 19), 12))
    '					End If
    '			End Select

    '			Wait_Form.Label1(3).Text = i
    '			Wait_Form.Refresh()
    '			temp = a.ReadLine
    '			i = i + 1
    '		Loop 

    '		If rs_Month.EOF = False Then rs_Month.Update()
    '		'add the last record

    '		Wait_Form.Label1(3).Text = i
    '		Wait_Form.Refresh()
    '		Wait_Form.Close()
    '		rs_Month.Close()

    '		Exit Sub
    'goError: 
    '		MsgBox(Err.Description & " " & name1,  , "Function: ReadFile in Reach Module")
    '	End Sub

    Private Sub read_SaveDay()
        Dim date1 As Short
        Dim slashPos As Short
        Dim pointPos As Short
        Dim i As Short
        Dim query As String
        Dim newRow As DataRow
        'Dim table As DataTable = GetTable()
        Dim fnCSV As String = Initial.Swat_Output & "\Reach.csv"
        Dim a As System.IO.StreamReader

        On Error GoTo goError

        'a = fs.OpenTextFile(name1)
        a = New System.IO.StreamReader(name1)

        For i = 1 To 9
            a.ReadLine()
        Next

        temp = a.ReadLine
        pointPos = InStrRev(name1, ".")
        slashPos = InStrRev(name1, "\")
        date1 = year_Renamed - 1
        'start
        Do While a.EndOfStream <> True
            If Val(Mid(temp, 20, 6)) = 1 And Val(Mid(temp, 6, 5)) = 1 Then date1 += 1
            buildCVS(date1, temp)
            'If Val(Mid(temp, 20, 6)) = 1 And Val(Mid(temp, 6, 5)) = 1 Then date1 += 1
            'jDay = Val(Mid(temp, 23, 3))
            'Call jmonth(date1)
            '' Create an array of four objects and add it as a row.
            '' Get middle part of string: Dim m As String = Mid(value, 5, 3)
            'Dim v(12) As Object
            ''Subbasin
            'v(0) = Val(Mid(temp, 1, 10))
            ''rch
            'v(1) = Val(Mid(temp, 6, 5))
            ''Year
            'v(2) = date1
            ''Mon
            'v(3) = MonthNum
            ''Area
            'v(4) = 0
            ''Flow_Out
            'v(5) = Val(Mid(temp, 50, 12))
            ''Sed_Out
            'v(6) = Val(Mid(temp, 98, 12))
            ''Orgn_Out
            'v(7) = Val(Mid(temp, 134, 12))
            ''OrgP_Out
            'v(8) = Val(Mid(temp, 158, 12))
            ''NO3_Out
            'v(9) = Val(Mid(temp, 182, 12))
            ''MinP_Out
            'v(10) = Val(Mid(temp, 254, 12))
            ''TotalN
            'v(11) = Val(Mid(temp, 134, 12)) + Val(Mid(temp, 182, 12))
            ''TotalP
            'v(12) = Val(Mid(temp, 158, 12)) + Val(Mid(temp, 254, 12))
            'table.Rows.Add(v)
            'temp = a.ReadLine
        Loop
        'end
        a.Close()
        'DataTable2CSV(table, fnCSV)
        'AccessConn.Open()
'        dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & Initial.Output_files & "\Local.mdb;"
'        myConnection = New OleDb.OleDbConnection
'        myConnection.ConnectionString = dbConnectString
'        myConnection.Open()
'        'Existing table
'        Dim AccessCommand As New System.Data.OleDb.OleDbCommand _
'("INSERT INTO [Reach_Day] (Subbasin, rch, [Year], [Mon], Area, Flow_Out, Sed_Out, Orgn_Out, OrgP_Out, NO3_Out, MinP_Out, TotalN, TotalP) SELECT F1, F2, F3, F4, F5,F6, F7, F8, F9, F10,F11, F12, F13 FROM [Text;Database=" & Initial.Swat_Output & ";Hdr=No].[Reach.csv]", myConnection)
'        AccessCommand.ExecuteNonQuery()
'        myConnection.Close()
        Exit Sub
goError:
        MsgBox(Err.Description & " " & name1, , "Function: ReadFile in Reach Module")
    End Sub

    Private Sub read_SaveDay(ByVal subNumber As String, ByVal subName As String)
        Dim slashPos As Short
        Dim pointPos As Short
        Dim i As Short
        Dim query As String
        Dim date1 As Short
        'Dim newRow As DataRow
        'Dim table As DataTable = GetTable()
        Dim a As System.IO.StreamReader

        On Error GoTo goError

        'a = fs.OpenTextFile(name1)
        a = New System.IO.StreamReader(name1)

        For i = 1 To 9
            a.ReadLine()
        Next

        temp = a.ReadLine
        pointPos = InStrRev(name1, ".")
        slashPos = InStrRev(name1, "\")
        date1 = year_Renamed - 1
        Do While a.EndOfStream <> True
            If Val(Mid(temp, 20, 6)) = 1 And Val(Mid(temp, 6, 5)) = 1 Then date1 += 1
            i = Val(Mid(temp, 6, 5))
            If i = subNumber Then
                buildCVS(date1, temp)
            End If
            temp = a.ReadLine
        Loop

        a.Close()
        'DataTable2CSV(table, fnCSV)
        'AccessConn.Open()
        Exit Sub
goError:
        MsgBox(Err.Description & " " & name1, , "Function: ReadFile in Reach Module")
    End Sub

Public Sub buildCVS(ByRef date1 As Short, ByRef temp As String)
    Dim fnCSV As String = Initial.Swat_Output & "\Reach.csv"
    Dim sep As String = ""
    Dim builder As System.Text.StringBuilder

    If writer Is Nothing Then writer = New System.IO.StreamWriter(fnCSV)
    builder = New System.Text.StringBuilder
    sep = ""


    jDay = Val(Mid(temp, 23, 3))
    Call jmonth(date1)

    ' Create an array of four objects and add it as a row.
    ' Get middle part of string: Dim m As String = Mid(value, 5, 3)
    Dim v(12) As Object
    'Subbasin
    builder.Append(sep).Append(Val(Mid(temp, 1, 10)))
    sep = ","
    'v(0) = Val(Mid(temp, 1, 10))
    'rch
    builder.Append(sep).Append(Val(Mid(temp, 6, 5)))
    'v(1) = Val(Mid(temp, 6, 5))
    'Year
    builder.Append(sep).Append(date1)
    'v(2) = date1
    'Mon
    builder.Append(sep).Append(MonthNum)
    'v(3) = MonthNum
    'Area
    builder.Append(sep).Append(0)
    'v(4) = 0
    'Flow_Out
    builder.Append(sep).Append(Val(Mid(temp, 50, 12)))
    'v(5) = Val(Mid(temp, 50, 12))
    'Sed_Out
    builder.Append(sep).Append(Val(Mid(temp, 98, 12)))
    'v(6) = Val(Mid(temp, 98, 12))
    'Orgn_Out
    builder.Append(sep).Append(Val(Mid(temp, 134, 12)))
    'v(7) = Val(Mid(temp, 134, 12))
    'OrgP_Out
    builder.Append(sep).Append(Val(Mid(temp, 158, 12)))
    'v(8) = Val(Mid(temp, 158, 12))
    'NO3_Out
    builder.Append(sep).Append(Val(Mid(temp, 182, 12)))
    'v(9) = Val(Mid(temp, 182, 12))
    'MinP_Out
    builder.Append(sep).Append(Val(Mid(temp, 254, 12)))
    'v(10) = Val(Mid(temp, 254, 12))
    'TotalN
    builder.Append(sep).Append(Val(Mid(temp, 134, 12)) + Val(Mid(temp, 182, 12)))
    'v(11) = Val(Mid(temp, 134, 12)) + Val(Mid(temp, 182, 12))
    'TotalP
    builder.Append(sep).Append(Val(Mid(temp, 158, 12)) + Val(Mid(temp, 254, 12)))
    'v(12) = Val(Mid(temp, 158, 12)) + Val(Mid(temp, 254, 12))
    writer.WriteLine(builder.ToString())
    'table.Rows.Add(v)
End Sub
    Public Sub jmonth(ByVal yearNum As Short)
        Dim leapmo2 As Object
        Dim leapmo1 As Object
        'determined month from julian day
        Dim leap As Double
        Dim k As Short

        leapmo1 = New Object() {0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365}
        leapmo2 = New Object() {0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366}
        leap = yearNum Mod 4.0#
        If (leap > 0) Then
            ' --- -  this is NOT a leap year
            For k = 11 To 0 Step -1
                If (jDay > leapmo1(k)) Then
                    ' --- -        month is needed for monthly evaporation table
                    MonthNum = k + 1
                    Exit For
                End If
            Next
        Else
            ' --- -  this IS a leap year
            For k = 11 To 0 Step -1
                If (jDay > leapmo2(k)) Then
                    ' --- -        month is needed for monthly evaporation table
                    MonthNum = k + 1
                    Exit For
                End If
            Next
        End If
    End Sub
    ''call the actual function that writes datatable to CSV file
    'Sub DataTable2CSV(ByVal table As DataTable, ByVal filename As String)
    '    DataTable2CSV(table, filename, ",")
    'End Sub
    ''write datatable to CSV file
    'Sub DataTable2CSV(ByVal table As DataTable, ByVal filename As String, ByVal sepChar As String)
    '    Dim writer As System.IO.StreamWriter
    '    Try
    '        writer = New System.IO.StreamWriter(filename)

    '        ' first write a line with the columns name
    '        Dim sep As String = ""
    '        Dim builder As New System.Text.StringBuilder
    '        'For Each col As DataColumn In table.Columns
    '        '    builder.Append(sep).Append(col.ColumnName)
    '        '    sep = sepChar
    '        'Next
    '        'writer.WriteLine(builder.ToString())

    '        ' then write all the rows
    '        For Each row As DataRow In table.Rows
    '            sep = ""
    '            builder = New System.Text.StringBuilder

    '            For Each col As DataColumn In table.Columns
    '                builder.Append(sep).Append(row(col.ColumnName))
    '                sep = sepChar
    '            Next
    '            writer.WriteLine(builder.ToString())
    '        Next
    '    Finally
    '        If Not writer Is Nothing Then writer.Close()
    '    End Try
    'End Sub
    ''create table Reach_Day structure: Subbasin, rch, [Year], [Mon], Area, Flow_Out, Sed_Out, Orgn_Out, OrgP_Out, NO3_Out, MinP_Out, TotalN, TotalP
    'Function GetTable() As DataTable
    '    ' Generate a new DataTable.
    '    ' ... Add columns.
    '    Dim table As DataTable = New DataTable
    '    table.Columns.Add("Subbasin", GetType(String))
    '    table.Columns.Add("rch", GetType(Integer))
    '    table.Columns.Add("Year", GetType(Integer))
    '    table.Columns.Add("Mon", GetType(Integer))
    '    table.Columns.Add("Area", GetType(Double))
    '    table.Columns.Add("Flow_Out", GetType(Double))
    '    table.Columns.Add("Sed_Out", GetType(Double))
    '    table.Columns.Add("Orgn_Out", GetType(Double))
    '    table.Columns.Add("OrgP_Out", GetType(Double))
    '    table.Columns.Add("NO3_Out", GetType(Double))
    '    table.Columns.Add("MinP_Out", GetType(Double))
    '    table.Columns.Add("TotalN", GetType(Double))
    '    table.Columns.Add("TotalP", GetType(Double))
    '    Return table
    'End Function
    '	Private Sub read_Year()
    '		Dim Offset As Object
    '		Dim date1 As Object
    '		Dim i As Object

    '		z = fs.createtextfile(Initial.Swat_Output & "\Titles.rch")
    '		a = fs.OpenTextFile(name1)
    '		rs_Year = New ADODB.Recordset

    '		rs_Year.Open("Reach_Year", cn, ADOR.CursorTypeEnum.adOpenKeyset, ADOR.LockTypeEnum.adLockOptimistic)

    '		Wait_Form.Label1(2).Text = "Number of Records Transfered:"
    '		Wait_Form.Show()
    '		For i = 1 To 9
    '			z.WriteLine(a.ReadLine)
    '		Next 

    '		z.Close()
    '		temp = a.ReadLine
    '		i = 1

    '		Do While a.AtEndOfStream <> True

    '			date1 = Val(Mid(temp, 20, 6))
    '			If date1 > 1900 Then
    '				year_Renamed = date1 + 1
    '				rs_Year.AddNew()
    '				rs_Year.Fields.Item("Reach").Value = Left(temp, 5)
    '				rs_Year.Fields.Item("RCH").Value = Val(Mid(temp, 6, 5))
    '				rs_Year.Fields.Item("year").Value = Val(Mid(temp, 20, 6))
    '				Offset = 12
    '				rs_Year.Fields.Item("area").Value = Val(Mid(temp, 26 + (Offset * 0), 12))
    '				rs_Year.Fields.Item("FLOW_Out").Value = Val(Mid(temp, 26 + (Offset * 2), 12))
    '				rs_Year.Fields.Item("Sed_Out").Value = Val(Mid(temp, 26 + (Offset * 6), 12))
    '				rs_Year.Fields.Item("OrgN_Out").Value = Val(Mid(temp, 26 + (Offset * 9), 12))
    '				rs_Year.Fields.Item("OrgP_Out").Value = Val(Mid(temp, 26 + (Offset * 11), 12))
    '				rs_Year.Fields.Item("NO3_Out").Value = Val(Mid(temp, 26 + (Offset * 13), 12))
    '				rs_Year.Fields.Item("MinP_Out").Value = Val(Mid(temp, 26 + (Offset * 19), 12))
    '				rs_Year.Update()
    '			End If

    '			Wait_Form.Label1(3).Text = i
    '			Wait_Form.Refresh()
    '			temp = a.ReadLine
    '			i = i + 1
    '		Loop 

    '		Wait_Form.Label1(3).Text = i
    '		Wait_Form.Refresh()
    '		Wait_Form.Close()
    '		rs_Year.Close()

    '		Exit Sub
    'goError: 
    '		MsgBox(Err.Description & " " & name1,  , "Function: ReadFile in Reach Module")
    '	End Sub

    '	Private Sub read_Total()
    '		Dim Offset As Object
    '		Dim date1 As Object
    '		Dim i As Object

    '		'UPGRADE_WARNING: Couldn't resolve default property of object Initial.Swat_Output. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		'UPGRADE_WARNING: Couldn't resolve default property of object fs.createtextfile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		z = fs.createtextfile(Initial.Swat_Output & "\Titles.rch")
    '		'UPGRADE_WARNING: Couldn't resolve default property of object fs.OpenTextFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		a = fs.OpenTextFile(name1)
    '		rs_Total = New ADODB.Recordset

    '		rs_Total.Open("Reach_Total", cn, ADOR.CursorTypeEnum.adOpenKeyset, ADOR.LockTypeEnum.adLockOptimistic)

    '		Wait_Form.Label1(2).Text = "Number of Records Transfered:"
    '		Wait_Form.Show()
    '		For i = 1 To 9
    '			'UPGRADE_WARNING: Couldn't resolve default property of object a.ReadLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object z.WriteLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			z.WriteLine(a.ReadLine)
    '		Next 

    '		'UPGRADE_WARNING: Couldn't resolve default property of object z.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		z.Close()
    '		'UPGRADE_WARNING: Couldn't resolve default property of object a.ReadLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		temp = a.ReadLine
    '		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		i = 1

    '		'UPGRADE_WARNING: Couldn't resolve default property of object a.AtEndOfStream. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		Do While a.AtEndOfStream <> True
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object date1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			date1 = Val(Mid(temp, 20, 6))
    '			Select Case date1
    '				Case Is > 1900
    '				Case Else
    '					'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '					If Mid(temp, 24, 1) = "." Then
    '						'With rs_Total
    '						rs_Total.AddNew()
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("Reach").Value = Left(temp, 5)
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("RCH").Value = Val(Mid(temp, 6, 5))
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("GIS").Value = Val(Mid(temp, 11, 9))
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("MON").Value = Val(Mid(temp, 20, 6))
    '						'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						Offset = 12
    '						'                    For j = 1 To 43
    '						'                        rs_Total.Fields(j + 3) = Val(Mid(temp, 26 + Offset, 12))
    '						'                        Offset = Offset + 12
    '						'                    Next
    '						'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("area").Value = Val(Mid(temp, 26 + (Offset * 0), 12))
    '						'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("FLOW_Out").Value = Val(Mid(temp, 26 + (Offset * 2), 12))
    '						'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("Sed_Out").Value = Val(Mid(temp, 26 + (Offset * 6), 12))
    '						'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("OrgN_Out").Value = Val(Mid(temp, 26 + (Offset * 9), 12))
    '						'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("OrgP_Out").Value = Val(Mid(temp, 26 + (Offset * 11), 12))
    '						'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("NO3_Out").Value = Val(Mid(temp, 26 + (Offset * 13), 12))
    '						'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '						rs_Total.Fields.Item("MinP_Out").Value = Val(Mid(temp, 26 + (Offset * 19), 12))
    '						rs_Total.Update()
    '						'rs_Total.Requery
    '						'End With
    '					End If
    '			End Select

    '			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			Wait_Form.Label1(3).Text = i
    '			Wait_Form.Refresh()
    '			'UPGRADE_WARNING: Couldn't resolve default property of object a.ReadLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			temp = a.ReadLine
    '			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			i = i + 1
    '		Loop 

    '		'add the last record
    '		With rs_Total
    '			.AddNew()
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			.Fields("Reach").Value = Left(temp, 5)
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			.Fields("RCH").Value = Val(Mid(temp, 6, 5))
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			.Fields("GIS").Value = Val(Mid(temp, 11, 9))
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			.Fields("MON").Value = Val(Mid(temp, 20, 6))
    '			'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			Offset = 12
    '			'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			rs_Total.Fields.Item("area").Value = Val(Mid(temp, 26 + (Offset * 0), 12))
    '			'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			rs_Total.Fields.Item("FLOW_Out").Value = Val(Mid(temp, 26 + (Offset * 2), 12))
    '			'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			rs_Total.Fields.Item("Sed_Out").Value = Val(Mid(temp, 26 + (Offset * 6), 12))
    '			'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			rs_Total.Fields.Item("OrgN_Out").Value = Val(Mid(temp, 26 + (Offset * 9), 12))
    '			'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			rs_Total.Fields.Item("OrgP_Out").Value = Val(Mid(temp, 26 + (Offset * 11), 12))
    '			'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			rs_Total.Fields.Item("NO3_Out").Value = Val(Mid(temp, 26 + (Offset * 13), 12))
    '			'UPGRADE_WARNING: Couldn't resolve default property of object Offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '			rs_Total.Fields.Item("MinP_Out").Value = Val(Mid(temp, 26 + (Offset * 19), 12))
    '			rs_Total.Update()
    '			.Requery()
    '		End With

    '		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		Wait_Form.Label1(3).Text = i
    '		Wait_Form.Refresh()
    '		Wait_Form.Close()
    '		rs_Total.Close()

    '		Exit Sub
    'goError: 
    '		MsgBox(Err.Description & " " & name1,  , "Function: ReadFile in Reach Module")
    '	End Sub

    '	Public Sub CreateTable(ByRef tableName As String)
    '		Dim cn As ADODB.Connection
    '		Dim Cat As ADOX.Catalog
    '		Dim objTable As Object
    '		Dim objTable1 As ADOX.Table
    '		Dim i As Short
    '		Dim tempiMonth, tempfMonth As Object
    '		Dim tempSite As Boolean

    '		On Error GoTo Errors
    '		cn = New ADODB.Connection
    '		Cat = New ADOX.Catalog
    '		objTable = New ADOX.Table
    '		objTable1 = New ADOX.Table
    '		'UPGRADE_WARNING: Couldn't resolve default property of object tempiMonth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		tempiMonth = False
    '		'UPGRADE_WARNING: Couldn't resolve default property of object tempfMonth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		tempfMonth = False : tempSite = False
    '		'Open the connection
    '		'UPGRADE_WARNING: Couldn't resolve default property of object Initial.Output_files. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Initial.Output_files & "\Local.mdb")

    '		'Open the Catalog
    '		Cat.ActiveConnection = cn

    '		'Create the table
    '		For i = 0 To Cat.Tables.Count - 1
    '			If Cat.Tables(i).name = tableName Then GoTo finish
    '		Next 
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.name = tableName
    '		'Create and Append a new field to the tablename Columns Collection
    '		'objTable.Columns.Append "reach", adString
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("rch", ADOR.DataTypeEnum.adInteger)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("Year", ADOR.DataTypeEnum.adInteger)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("Mon", ADOR.DataTypeEnum.adInteger)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("FLOW_OUT", ADOR.DataTypeEnum.adSingle)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("SED_OUT", ADOR.DataTypeEnum.adSingle)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("ORGN_OUT", ADOR.DataTypeEnum.adSingle)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("ORGP_OUT", ADOR.DataTypeEnum.adSingle)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("NO3_OUT", ADOR.DataTypeEnum.adSingle)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("MINP_OUT", ADOR.DataTypeEnum.adSingle)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("TotalN", ADOR.DataTypeEnum.adSingle)
    '		'UPGRADE_WARNING: Couldn't resolve default property of object objTable.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		objTable.Columns.Append("TotalP", ADOR.DataTypeEnum.adSingle)

    '		'Create and Append a new key. Note that we are merely passing
    '		'the "PimaryKey_Field" column as the source of the primary key. This
    '		'new Key will be Appended to the Keys Collection of "Test_Table"
    '		'objTable.Keys.Append "PrimaryKey", adKeyPrimary, "PrimaryKey_Field"

    '		'Append the newly created table to the Tables Collection
    '		Cat.Tables.Append(objTable)

    '		objTable1.name = "Reach_Day"
    'finish: 
    '		' clean up objects
    '		'Set objKey = Nothing
    '		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		objTable = Nothing
    '		'UPGRADE_NOTE: Object objTable1 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		objTable1 = Nothing
    '		'UPGRADE_NOTE: Object Cat may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		Cat = Nothing
    '		cn.Close()
    '		'UPGRADE_NOTE: Object cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		cn = Nothing

    '		'MsgBox "CEEOT-SWAPP was updated succesfully - Now you can calculate E-Values for this scenario directly"
    '		Exit Sub

    'Errors: 
    '		MsgBox(Err.Description & "-- program Update --> CreateTable")
    '	End Sub

    '	Public Sub CreateTotalN_P(ByRef tableName As String)
    '		Dim j As Object
    '		Dim cn As ADODB.Connection
    '		Dim Cat As ADOX.Catalog
    '		Dim objTable As Object
    '		Dim objTable1 As ADOX.Table
    '		Dim i As Short
    '		Dim totalN, totalP As Object
    '		Dim tempSite As Boolean

    '		On Error GoTo Errors
    '		cn = New ADODB.Connection
    '		Cat = New ADOX.Catalog
    '		objTable = New ADOX.Table
    '		objTable1 = New ADOX.Table
    '		'UPGRADE_WARNING: Couldn't resolve default property of object totalN. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		totalN = False
    '		'UPGRADE_WARNING: Couldn't resolve default property of object totalP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		totalP = False
    '		'Open the connection
    '		'UPGRADE_WARNING: Couldn't resolve default property of object Initial.Output_files. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		If Initial.Output_files = "" Then
    '			MsgBox("You need to open a project first - Open your project and try again")
    '			Exit Sub
    '		End If
    '		'UPGRADE_WARNING: Couldn't resolve default property of object Initial.Output_files. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '		cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Initial.Output_files & "\local.mdb")

    '		'Open the Catalog
    '		Cat.ActiveConnection = cn

    '		'Create the table
    '		For i = 0 To Cat.Tables.Count - 1
    '			If Cat.Tables(i).name = tableName Then
    '				'Append the newly created table to the Tables Collection
    '				For j = 0 To Cat.Tables(tableName).Columns.Count - 1
    '					Select Case StrConv(Cat.Tables(tableName).Columns(j).name, VbStrConv.LowerCase)
    '						Case "totaln"
    '							'UPGRADE_WARNING: Couldn't resolve default property of object totalN. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '							totalN = True
    '						Case "totalp"
    '							'UPGRADE_WARNING: Couldn't resolve default property of object totalP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '							totalP = True
    '					End Select
    '				Next 

    '				'UPGRADE_WARNING: Couldn't resolve default property of object totalN. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '				If totalN = False Then Cat.Tables(tableName).Columns.Append("totalN", ADOR.DataTypeEnum.adInteger)
    '				'UPGRADE_WARNING: Couldn't resolve default property of object totalP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '				If totalP = False Then Cat.Tables(tableName).Columns.Append("totalP", ADOR.DataTypeEnum.adInteger)
    '			End If
    '		Next 
    '		' clean up objects
    '		'Set objKey = Nothing
    '		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		objTable = Nothing
    '		'UPGRADE_NOTE: Object objTable1 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		objTable1 = Nothing
    '		'UPGRADE_NOTE: Object Cat may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		Cat = Nothing
    '		cn.Close()
    '		'UPGRADE_NOTE: Object cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '		cn = Nothing

    '		'MsgBox "CEEOT-SWAPP was updated succesfully - Now you can calculate E-Values for this scenario directly"
    '		Exit Sub

    'Errors: 
    '		MsgBox(Err.Description & "-- program Update --> CreateTable")
    '	End Sub
End Module