Option Strict Off
Option Explicit On

Imports System.IO
Imports VB = Microsoft.VisualBasic

Public Class General
    Private mvarfilehru As String 'local copy
    Private mvarfileName As String 'local copy
    Private mvarfilepnd As String 'local copy
    Private mvarfilerte As String 'local copy
    Private mvarnumber As Short 'local copy
    Private mvarflag As Boolean 'local copy
    Private mvarfilesub As String 'local copy
    Private mvarfilesol As String 'local copy
    Private mvarfilesoi As String 'local copy
    Private mvarfilemgt As String 'local copy
    Private mvarfilewht As String 'local copy
    Private mvarLast As Object 'local copy
    Private mvarLasttmp As Object 'local copy
    Private mvarfilesit As String 'local copy
    Private mvarfilewp1 As String 'local copy
    Private mvarfilewgn As String 'local copy
    Private mvarfilewpm As String 'local copy
    Private mvarfilechm As String 'local copy
    'Grazing Information
    Private mvarManureID As Short
    Private mvarBioConsumed As Single
    Private mvarManureProduced As Single

    Dim value As String
    Dim NumberFormat As String
    Dim roundformat As Integer
    Dim lenFormat As Integer
    Dim convertFormat As Convertion
    Private wsa1 As Single

    'local variable(s) to hold property value(s)


    Public Property filechm() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.filechm
            If IsReference(mvarfilechm) Then
                filechm = mvarfilechm
            Else
                filechm = mvarfilechm
            End If
        End Get
        Set(ByVal Value As String)
            If IsReference(Value) And Not TypeOf Value Is String Then
                'used when assigning an Object to the property, on the left side of a Set statement.
                'Syntax: Set x.filechm = Form1
                mvarfilechm = Value
            Else
                'used when assigning a value to the property, on the left side of an assignment.
                'Syntax: X.filechm = 5
                mvarfilechm = Value
            End If
        End Set
    End Property

    Public WriteOnly Property filewpm() As String
        Set(ByVal Value As String)
            mvarfilewpm = Value
        End Set
    End Property

    Public WriteOnly Property filewgn() As String
        Set(ByVal Value As String)
            mvarfilewgn = Value
        End Set
    End Property

    Public WriteOnly Property filewp1() As String
        Set(ByVal Value As String)
            mvarfilewp1 = Value
        End Set
    End Property

    Public WriteOnly Property filesit() As String
        Set(ByVal Value As String)
            mvarfilesit = Value
        End Set
    End Property

    Public Property Last() As Object
        Get
            Last = mvarLast
        End Get
        Set(ByVal Value As Object)
            mvarLast = Value
        End Set
    End Property


    Public Property Lasttmp() As Object
        Get
            Lasttmp = mvarLasttmp
        End Get
        Set(ByVal Value As Object)
            mvarLasttmp = Value
        End Set
    End Property
    Public WriteOnly Property filewht() As String
        Set(ByVal Value As String)
            mvarfilewht = Value
        End Set
    End Property

    Public WriteOnly Property filemgt() As String
        Set(ByVal Value As String)
            mvarfilemgt = Value
        End Set
    End Property

    Public WriteOnly Property filesol() As String
        Set(ByVal Value As String)
            mvarfilesol = Value
        End Set
    End Property

    Public WriteOnly Property filesoi() As String
        Set(ByVal Value As String)
            mvarfilesoi = Value
        End Set
    End Property

    Public WriteOnly Property filename() As String
        Set(ByVal Value As String)
            mvarfileName = Value
        End Set
    End Property

    Public WriteOnly Property fileSub() As String
        Set(ByVal Value As String)
            mvarfilesub = Value
        End Set
    End Property

    Public WriteOnly Property filehru() As String
        Set(ByVal Value As String)
            mvarfilehru = Value
        End Set
    End Property

    Public WriteOnly Property filerte() As String
        Set(ByVal Value As String)
            mvarfilerte = Value
        End Set
    End Property

    Public WriteOnly Property filepnd() As String
        Set(ByVal Value As String)
            mvarfilepnd = Value
        End Set
    End Property

    Public WriteOnly Property number() As Short
        Set(ByVal Value As Short)
            mvarnumber = Value
        End Set
    End Property

    Public WriteOnly Property flag() As Boolean
        Set(ByVal Value As Boolean)
            mvarflag = Value
        End Set
    End Property

    Public Sub wpm11310()
        Dim temp As String
        Dim longitude As String
        Dim altitude As String = String.Empty
        Dim current_line, i, j As Integer
        Dim p As Object
        Dim fs As Object
        'Dim adoConnection As ADODB.Connection
        Dim ADORecordset As DataTable
        Dim TakeField As Convertion
        Dim value As String

        On Error GoTo goError

        TakeField = New Convertion
        'adoConnection = New ADODB.Connection
        'adoConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir1 & "\Project_Parameters.mdb"
        'adoConnection.Open()
        ADORecordset = New DataTable
        ADORecordset = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'Wpm11310' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")
        convertFormat = New Convertion
        fs = CreateObject("Scripting.FileSystemObject")

        p = fs.OpenTextFile(Initial.Output_files & "\" & Initial.wpm1, 8, True)

        With ADORecordset
            '.MoveFirst()
            current_line = .Rows(0).Item("Line")
            For j = 0 To .Rows.Count - 1
                'Do While .EOF <> True
                If (.Rows(j).Item("SwatFile") <> "") Then
                    Select Case .Rows(j).Item("SwatFile")
                        Case "*.wgn"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilewgn)
                    End Select

                    TakeField.Leng = .Rows(j).Item("Leng")
                    TakeField.LineNum = .Rows(j).Item("Lines")
                    TakeField.Inicia = .Rows(j).Item("Inicia")
                    value = TakeField.value()
                Else
                    If .Rows(j).Item("Value") = "Blank" Then
                        value = " "
                    Else
                        value = .Rows(j).Item("Value")
                    End If
                End If
                If .Rows(j).Item("Line") = 1 Then
                    altitude = convertFormat.Convert(System.Math.Abs(CDbl(Val(value))), "####0.00")
                Else
                    longitude = convertFormat.Convert(System.Math.Abs(CDbl(Val(value))), "####0.00")
                    temp = convertFormat.Convert(mvarnumber, "####0") & Initial.Espace1 & mvarfilewp1.PadLeft(12) & altitude & longitude & "    " & mvarfilewp1
                    p.WriteLine(temp)
                End If

                '.MoveNext()
            Next
        End With
        p.Close()

        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub
    Public Sub suba(ByRef field_i As Short)
        Dim rchl, CHL As Single
        Dim totalare As Single
        Dim current_line, i As Integer
        Dim e As Object
        Dim fs As Object
        'Dim cn As ADODB.Connection
        Dim subarea As DataTable
        Dim ADORecordset As DataSet
        Dim grazing As DataTable
        Dim TakeField As Convertion
        Dim value, temp As String
        Dim totalArea As Single
        Dim pcof As Single
        Dim rsae As Single = 0
        Dim rsap As Single
        Dim values() As String
        Dim swat, apex As String
        Dim area As Single

        On Error GoTo goError

        TakeField = New Convertion
        'cn = New ADODB.Connection
        subarea = New DataTable
        grazing = New DataTable

        'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir1 & "\Project_Parameters.mdb"
        'cn.Open()

        ADORecordset = getDBDataSet("SELECT * FROM Apexfiles WHERE Apexfile = 'Subarea' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")
        fs = CreateObject("Scripting.FileSystemObject")
        e = fs.CreateTextFile(Initial.Output_files & "\" & mvarfilesub)
        convertFormat = New Convertion

        'subarea.Open("SELECT * FROM subarea ", cn, ADOR.CursorTypeEnum.adOpenDynamic, ADOR.LockTypeEnum.adLockOptimistic)
        'subarea.AddNew()
        swat = mvarfileName
        apex = mvarfilesub

        With ADORecordset.Tables(0)
            current_line = .Rows(0).Item("Line")

            For i = 0 To .Rows.Count - 1
                If Not IsDBNull(.Rows(i).Item("SwatFile")) AndAlso .Rows(i).Item("SwatFile") <> "" Then
                    Select Case .Rows(i).Item("SwatFile")
                        Case "*.hru"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilehru)
                        Case "*.rte"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilerte)
                        Case "*.pnd"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilepnd)
                        Case "*.sub"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfileName)
                        Case "*.mgt"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilemgt)
                        Case "*.sol"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilesol)
                        Case "basins.bsn"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(Initial.BaseFile)
                        Case "Basins.fig"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(Initial.figsFile)
                        Case Else
                            TakeField.filename = Initial.Input_Files & "\" & Trim(.Rows(i).Item("SwatFile"))
                    End Select

                    TakeField.Leng = .Rows(i).Item("Leng")
                    TakeField.LineNum = .Rows(i).Item("Lines")
                    TakeField.Inicia = .Rows(i).Item("Inicia")
                    value = TakeField.value() '1/15/2014 Yang?
                Else
                    If .Rows(i).Item("Value") = "Blank" Then
                        value = " "
                    Else
                        value = .Rows(i).Item("Value")
                    End If
                End If

                temp = .Rows(i).Item("Line") & .Rows(i).Item("Field")
                Select Case temp
                    Case "00"
                        totalare = Val(value)
                    Case "11"
                        value = mvarnumber
                    Case "12"
                        value = mvarnumber
                    Case "23"
                        value = mvarnumber
                    Case "24"
                        value = mvarnumber
                    Case "25"
                        grazing = getDBDataTable("SELECT manureId, manureProduced, bioConsumed, herdsNumber, ownerID FROM grazing WHERE fileName = '" & mvarfilehru & "'")
                        If grazing.Rows.Count > 0 Then
                            value = grazing.Rows(0).Item("ownerID")
                        End If
                        grazing.Dispose()
                        grazing = Nothing
                    Case "28"   'commentarized because it is not needed any more. TODO I have to check for flood plain information lenght, width, and fraction.
                        If Val(value) > 0 Then value = 1
                        ' value = 0
                    Case "210"
                        value = Initial.File_Number
                    'value = Weather_Code(mvarfileName)
                    Case "414"
                        values = Split(value, "|")
                        value = System.Math.Round(CSng(values(0)) * totalare * 100, 4) 'calculate hru (field) fraction and convert km2 to ha.
                        wsa1 = value
                        area = wsa1
                        CHL = Format(System.Math.Sqrt((area * 0.01) * 2), "0000.000")
                        rchl = CHL * 0.9
                    Case "415"
                    'CHL = CHL
                    Case "419"
                        values = Split(value, "|")
                        value = values(0)
                        If Val(value) <= 0 Then value = "  0.0001"
                        'Case "420"
                        'values = Split(value, "|")
                        'value = values(0)
                        'Case "421"
                        'values = Split(value, "|")
                        'value = values(0)
                    Case "422"
                        value = 0.367 * Val(value) ^ 0.2967
                    Case "523"
                    'rchl = rchl
                    Case "532"
                        If Val(value) > 0 Then value = (wsa1 * 0.01) / (Val(value) * 0.001)
                    Case "634"
                        rsae = Val(value)
                    Case "635"
                        If rsae > 0 Then value = (Val(value) * 10000) / (rsae * 10)
                    Case "637"
                        rsap = Val(value)
                    Case "638"
                        If rsap > 0 Then value = (Val(value) * 10000) / (rsap * 10)
                    Case "74"
                        pcof = Val(value)
                    Case "962"
                        value = Val(value) / 24
                    Case "1173"
                        grazing = getDBDataTable("SELECT manureId, manureProduced, bioConsumed, herdsNumber, ownerID FROM grazing WHERE fileName = '" & mvarfilehru & "'")
                        If grazing.Rows.Count > 0 Then
                            value = grazing.Rows(0).Item("herdsNumber")
                        End If
                        grazing.Dispose()
                        grazing = Nothing
                End Select

                If Not IsDBNull(.Rows(i).Item("Format")) AndAlso .Rows(i).Item("Format") <> "" Then
                    lenFormat = Len(.Rows(i).Item("Format"))
                    roundformat = Right(Trim(.Rows(i).Item("Format")), 1)
                    NumberFormat = Left(Trim(.Rows(i).Item("Format")), lenFormat - 2)
                    If IsNumeric(value) Then
                        If temp = "414" And area > 9999 Then NumberFormat = "#####0.0"
                        value = convertFormat.Convert(System.Math.Round(CSng(value), roundformat), NumberFormat)
                    Else
                        'If value.Trim = "" Then value = "0" 'Val(value) 1/15/2014
                        If Val(value) = 0.0 Then value = "0"
                        value = convertFormat.Convert(value, NumberFormat)
                    End If
                End If

                If temp <> "00" Then

                    If i < .Rows.Count - 1 Then
                        If .Rows(i).Item("Line") <> 0 Then
                            If current_line = .Rows(i + 1).Item("Line") Then
                                e.Write(value)
                            Else
                                e.WriteLine(value)
                            End If
                        End If
                        current_line = .Rows(i + 1).Item("Line")
                    End If
                Else
                    current_line = .Rows(i + 1).Item("Line")
                End If
            Next

            modifyRecords("INSERT INTO Subarea (Swat,Apex,Area,chl,rchl) VALUES('" & swat & "','" & apex & "'," & area & "," & CHL & "," & rchl & ")")
            'subarea.Update()
        End With

        e.Close()
        'subarea.Close()
        'cn.Close()

        Exit Sub
goError:
        MsgBox(Err.Description & " General.Suba " & temp)

    End Sub

    Public Sub Updatehru()  'OGM Check what this is doing
        Dim area As Single = 0
        Dim temp As String
        Dim i As Integer
        Dim swFile As StreamWriter
        Dim srFile As StreamReader
        Dim hru_fr As Single
        Dim totalArea As Single

        On Error GoTo goError

        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & mvarfilehru))
        swFile = New StreamWriter(File.Create(Initial.New_Swat & "\" & mvarfilehru))
        convertFormat = New Convertion

        Dim adoConnection, HRUArea, ADORecordset As DataTable
        Dim TakeField As Convertion

        TakeField = New Convertion

        ADORecordset = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'Updatehru' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")
        With ADORecordset
            If Not IsDBNull(.Rows(0).Item("SwatFile")) AndAlso .Rows(0).Item("SwatFile") <> "" Then
                Select Case .Rows(0).Item("SwatFile")
                    Case "*.hru"
                        TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfileName)
                    Case "basins.bsn"
                        TakeField.filename = Initial.Input_Files & "\" & Trim(Initial.BaseFile)
                    Case "*.sub"
                        TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfileName)
                    Case Else
                        TakeField.filename = Initial.Input_Files & "\" & Trim(.Rows(0).Item("SwatFile"))
                End Select

                TakeField.Leng = .Rows(0).Item("Leng")
                TakeField.LineNum = .Rows(0).Item("Lines")
                TakeField.Inicia = .Rows(0).Item("Inicia")
                totalArea = TakeField.value()
            End If
            i = 1
            Do While srFile.EndOfStream <> True
                If .Rows(1).Item("Lines") = i Then
                    temp = srFile.ReadLine
                    totalArea = Val(totalArea)
                    hru_fr = area / (totalArea * 100)
                    If Initial.Version = "4.0.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then hru_fr = 1
                    swFile.Write(convertFormat.Convert(hru_fr, "#######0.000000" & " "))
                    swFile.WriteLine(Mid(temp, 17, 74))
                End If
                i = i + 1
                swFile.WriteLine(srFile.ReadLine)
            Loop
        End With
        srFile.Close()
        srFile.Dispose()
        srFile = Nothing

        swFile.Close()
        swFile.Dispose()
        swFile = Nothing

        Exit Sub
goError:
        MsgBox(Err.Description & " Geenral.Updatehru")

    End Sub

    Public Sub Soil()
        Dim Print_levels As Single = 0
        Dim soillevel As Single
        Dim i, j As Integer
        Dim Offset As Integer
        Dim current_line As Integer
        Dim hsg As Object
        Dim swFile As StreamWriter
        Dim srFile As StreamReader
        Dim fs As Object
        Dim ADORecordset As DataTable
        Dim TakeField As Convertion
        Dim value As String
        Dim temp As String
        Dim value1 As Single
        Dim total_Records As Integer
        Dim Soil_levels As Short
        Dim tempval(10) As Object

        On Error GoTo goError

        TakeField = New Convertion

        ADORecordset = New DataTable
        ADORecordset = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'Soil' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")

        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & Trim(mvarfilesol)))
        swFile = New StreamWriter(File.Create(Initial.Output_files & "\" & Trim(mvarfilesoi)))
        convertFormat = New Convertion

        hsg = 0
        With ADORecordset
            current_line = .Rows(0).Item("Line")
            For j = 0 To .Rows.Count - 1
                If Not IsDBNull(.Rows(j).Item("SwatFile")) AndAlso .Rows(j).Item("SwatFile") <> "" Then
                    Select Case .Rows(j).Item("SwatFile")
                        Case "*.sol"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilesol)
                        Case "*.chm"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilechm)
                    End Select

                    TakeField.Leng = .Rows(j).Item("Leng")
                    TakeField.LineNum = .Rows(j).Item("Lines")
                    TakeField.Inicia = .Rows(j).Item("Inicia")
                    value = TakeField.value()
                Else
                    If .Rows(j).Item("Value") = "Blank" Then
                        value = " "
                    Else
                        value = .Rows(j).Item("Value")
                    End If
                End If
                temp = .Rows(j).Item("Line") & .Rows(j).Item("Field")
                Select Case temp
                    Case "00"
                        Offset = 1
                        For i = 1 To 9
                            tempval(i) = Mid(value, Offset, 12)
                            Offset = Offset + 12
                            If tempval(i) = "" Then Exit For
                        Next
                        soillevel = i - 1
                    Case "21"
                    'If Val(value) < 0.1 Then value = "    0.10"
                    Case "22"
                        Select Case value
                            Case "A"
                                value = "    1.00"
                            Case "B"
                                value = "    2.00"
                            Case "C"
                                value = "    3.00"
                            Case "D"
                                value = "    4.00"
                        End Select
                    Case "31"
                        value = convertFormat.Convert(soillevel, "####0.00")
                        Print_levels = Val(value)
                        If soillevel <= 2 Then value = 0
                    Case Is > "100"
                        Soil_levels = CShort(Right(temp, 1))
                End Select

                If Val(temp) >= 41 And CDbl(Val(temp)) < 50 Then
                    value = convertFormat.Convert(Val(value) / 1000, "####0.00")
                ElseIf Val(temp) >= 51 And CDbl(Val(temp)) < 60 Then
                    value = convertFormat.Convert(Val(value), "####0.00")
                ElseIf Val(temp) >= 81 And CDbl(Val(temp)) < 99 Then
                    value = convertFormat.Convert(Val(value), "####0.00")
                ElseIf Val(temp) >= 131 And CDbl(Val(temp)) < 140 Then
                    value = convertFormat.Convert(Val(value), "####0.00")
                ElseIf Val(temp) >= 161 And CDbl(Val(temp)) < 170 Then
                    value = convertFormat.Convert(Val(value), "####0.00")
                ElseIf Val(temp) >= 221 And CDbl(Val(temp)) < 230 Then
                    value = convertFormat.Convert(Val(value), "####0.00")
                End If

                If Not IsDBNull(.Rows(j).Item("Format")) AndAlso .Rows(j).Item("Format") <> "" Then
                    lenFormat = Len(.Rows(j).Item("Format"))
                    roundformat = Right(Trim(.Rows(j).Item("Format")), 1)
                    NumberFormat = Left(Trim(.Rows(j).Item("Format")), lenFormat - 2)
                    value = convertFormat.Convert(System.Math.Round(Val(value), roundformat), NumberFormat)
                End If

                If temp <> "00" Then
                    If j < .Rows.Count - 1 Then
                        If Soil_levels > Print_levels And .Rows(j).Item("Line") > 3 Then value = "        "
                        If Soil_levels = 0 And .Rows(j).Item("Line") > 3 Then value = "        "

                        If current_line = .Rows(j + 1).Item("Line") Then
                            swFile.Write(value)
                        Else
                            swFile.WriteLine(value)
                        End If
                        current_line = .Rows(j + 1).Item("Line")
                    Else
                        swFile.WriteLine(value)
                    End If
                End If
            Next
        End With

        srFile.Close()
        srFile.Dispose()
        srFile = Nothing

        swFile.Close()
        swFile.Dispose()
        swFile = Nothing

        Exit Sub
goError:
        MsgBox(Err.Description & " " & mvarfilesol)

    End Sub

    Public Sub Operations()
        Dim j As Integer
        Dim year2, month2, day2 As String
        Dim pcod As Integer
        Dim POpv7, POpv6, POpv5, POpv4, POpv3, POpv2, POpv1, Pblank As String
        Dim pmat, pcrp, PTrac As String
        Dim Condition As String
        Dim day1, month1, year1 As String
        Dim xmtu As Integer
        Dim jx4 As Integer
        Dim lyr As String
        Dim Fert As Single
        Dim day_temp As Integer
        Dim POpv2_Planting As Single
        Dim tempopc As String
        Dim i As Integer
        Dim a As StreamReader
        Dim lun As Object = Nothing
        Dim totalare As Double
        Dim temp9 As String
        Dim current_line As Integer
        Dim b As StreamWriter
        'Dim fs As Object
        Dim one As Integer
        Dim currdir As String
        Dim TakeField As Object
        Dim cn As DataTable = Nothing
        Dim adoCrop, adoParm, Crop1310, grazing As DataTable
        Dim ADORecordset As DataSet
        Dim grazingCur As DataTable
        Dim temp, value, oper, default_Renamed As String
        Dim RecOper() As String
        Dim temp1, Line1 As Integer
        Dim Temp2 As Integer
        Dim opv7, opv1, wsa2 As Single
        Dim xtp1 As Integer
        Dim animals As Single
        Dim manureID As Single
        Dim bioConsumed As Single
        Dim manureProduced As Single
        Dim flag As Boolean
        Dim col2, col1, col3, col4 As Integer
        Dim limit As Integer
        Dim k, p, o As Integer
        Dim date1 As Date
        Dim tempHerds As Short
        Dim query As String
        Dim sqltemp As String
        Dim sqltemp1, sqltemp2 As String
        Dim herdsNumber, ownerID As Integer
        Dim tablename As String
        Dim created As Boolean = False

        On Error GoTo goError

        TakeField = New Convertion
        grazingCur = New DataTable
        grazing = New DataTable
        temp = String.Empty
        currdir = CurDir()
        one = 1
        ADORecordset = getDBDataSet("SELECT * FROM Apexfiles WHERE Apexfile = 'Operations' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")

        'fs = CreateObject("Scripting.FileSystemObject")
        b = New StreamWriter(Initial.Output_files & "\" & mvarfilesub, False) 'Name of Operation file
        'b = fs.CreateTextFile(Initial.Output_files & "\" & mvarfilesub) 'Name of Operation file
        convertFormat = New Convertion

        Select Case Initial.Version
            Case "1.0.0"
                col1 = 18
                limit = 2
                col2 = 29
                col3 = 21
                col4 = 7
            Case "1.1.0"
                one = -1
                col1 = 16
                limit = 30
                col2 = 20
                col3 = 33
                col4 = 5
            Case "1.2.0"
                one = -1
                col1 = 16
                limit = 30
                col2 = 20
                col3 = 33
                col4 = 5
            Case "1.3.0"
                one = -1
                col1 = 16
                limit = 30
                col2 = 20
                col3 = 33
                col4 = 5
            Case "2.0.0"
                col1 = 18
                limit = 2
                col2 = 29
                col3 = 21
                col4 = 7
            Case "2.1.0"
                col1 = 16
                limit = 30
                col2 = 20
                col3 = 33
                col4 = 5
            Case "2.3.0"
                col1 = 16
                limit = 30
                col2 = 20
                col3 = 33
                col4 = 5
            Case "3.0.0"
                col1 = 18
                limit = 2
                col2 = 29
                col3 = 21
                col4 = 7
            Case "3.1.0"
                col1 = 16
                limit = 30
                col2 = 20
                col3 = 33
                col4 = 4
            Case "4.0.0"
                one = -1
                col1 = 18
                limit = 2
                col2 = 29
                col3 = 21
                col4 = 7
            Case "4.1.0"
                one = -1
                col1 = 16
                limit = 30
                col2 = 20
                col3 = 33
                col4 = 5
            Case "4.2.0"
                one = -1
                col1 = 16
                limit = 30
                col2 = 20
                col3 = 33
                col4 = 5
            Case "4.3.0"
                one = -1
                col1 = 16
                limit = 30
                col2 = 20
                col3 = 33
                col4 = 5
        End Select

        k = 0
        ReDim RecOper(k)
        Line1 = 3
        flag = False
        opv7 = 0.0#
        Temp2 = 0
        temp1 = Nothing

        With ADORecordset.Tables(0)
            current_line = .Rows(0).Item("Line")

            For i = 0 To .Rows.Count - 1
                If Not IsDBNull(.Rows(i).Item("SwatFile")) AndAlso .Rows(i).Item("SwatFile") <> "" Then
                    Select Case .Rows(i).Item("SwatFile")
                        Case "*.mgt"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilemgt)
                        Case "*.sol"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilesol)
                        Case "*.sub"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfileName)
                        Case "*.hru"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilehru)
                    End Select

                    TakeField.Leng = .Rows(i).Item("Leng")
                    TakeField.LineNum = .Rows(i).Item("Lines")
                    TakeField.Inicia = .Rows(i).Item("Inicia")
                    value = TakeField.value()
                Else
                    value = .Rows(i).Item("Value")
                End If

                temp9 = .Rows(i).Item("Line") & .Rows(i).Item("Field")
                Select Case CInt(temp9)
                    Case 0
                        totalare = Val(value)
                    Case 21
                        Select Case value
                            Case "A"
                                value = "   1"
                                cn = getDBDataTable("SELECT HSG_A FROM Curve_Number WHERE Land_Use = 21")
                            Case "B"
                                value = "   2"
                                cn = getDBDataTable("SELECT HSG_B FROM Curve_Number WHERE Land_Use = 21")
                            Case "C"
                                value = "   3"
                                cn = getDBDataTable("SELECT HSG_C FROM Curve_Number WHERE Land_Use = 21")
                            Case "D"
                                value = "   4"
                                cn = getDBDataTable("SELECT HSG_D FROM Curve_Number WHERE Land_Use = 21")
                            Case Else
                                value = "   1"
                                cn = getDBDataTable("SELECT HSG_A FROM Curve_Number WHERE Land_Use = 21")
                        End Select
                    Case 31
                        value = Val(value)
                    Case 14
                        value = System.Math.Round(Val(value) * totalare * 100, 4)
                        wsa2 = value
                End Select

                temp = ""
                If Not IsDBNull(.Rows(i).Item("Condition")) Then temp = .Rows(i).Item("Condition")

                If Not IsDBNull(.Rows(i).Item("Format")) AndAlso .Rows(i).Item("Format") <> "" Then
                    lenFormat = Len(.Rows(i).Item("Format"))
                    roundformat = Right(Trim(.Rows(i).Item("Format")), 1)
                    NumberFormat = Left(Trim(.Rows(i).Item("Format")), lenFormat - 2)
                    value = convertFormat.Convert(System.Math.Round(Val(value), roundformat), NumberFormat) '1/15/2014 Yang
                End If

                If i < .Rows.Count - 1 Then
                    If current_line <> 0 Then
                        If current_line = .Rows(i + 1).Item("Line") Then
                            b.Write(value)
                        Else
                            b.WriteLine(value)
                        End If
                    End If
                    current_line = .Rows(i + 1).Item("Line")
                End If

                If (temp = "1") Then
                    lun = cn.Rows(0).Item(0)
                Else
                    lun = value
                End If
            Next
        End With

        a = New StreamReader(Initial.Input_Files & "\" & Trim(mvarfilemgt))
        'Wait_Form.Label1(5).Text = mvarfilemgt
        'Wait_Form.Show()
        'Wait_Form.Refresh()
        'a = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(mvarfilemgt))
        For i = 1 To limit
            a.ReadLine()
        Next

        year1 = 1
        tempopc = ""
        POpv2_Planting = 0

        Do While a.EndOfStream <> True
            'Wait_Form.Label1(6).Text = "Year " & year1
            'Wait_Form.Label1(7).Text = "Operation " & temp
            'Wait_Form.Show()
            'Wait_Form.Refresh()
            If tempopc = "  5" Then
                temp = Left(temp, col1 - 1) & "  8"
                If day_temp > 1 Then Mid(temp, 7, 2) = convertFormat.Convert(day_temp, "#0")
                tempopc = ""
            Else
                temp = a.ReadLine
                tempopc = Mid(temp, col1, 3)
            End If

            Fert = Val(Mid(temp, 36, 8))

            If Mid(temp, col1, 3) = "  5" Then
                temp = Left(temp, col1 - 1) & "  7"
                day_temp = Val(Mid(temp, 7, 2))
                If day_temp > 1 Then Mid(temp, 7, 2) = convertFormat.Convert(day_temp - 1, "#0")
            End If

            If (Mid(temp, col1, 3) = "  3" Or Mid(temp, col1, 3) = "  4" Or Mid(temp, col1, 3) = " 11") Then
                lyr = Mid(temp, col2, 4)
            Else
                lyr = "   0"
            End If

            If (Mid(temp, col1, 3).Trim = "" Or Mid(temp, col1, 3).Trim = "17" Or Mid(temp, col1, 3).Trim = "0") Then '1/15/2014 Yang
                year1 = year1 + 1
            Else
                If Trim(Mid(temp, col2, 4)) = "18" And Trim(Mid(temp, col1, 4)) = "1" Then
                    adoParm = getDBDataTable("SELECT * FROM Parmopc where code = " & "  0")
                    lun = 31
                Else
                    adoParm = getDBDataTable("SELECT * FROM Parmopc where code = " & Mid(temp, col1, 3))
                End If

                If adoParm.Rows.Count = 0 Then
                    default_Renamed = "   "
                Else
                    default_Renamed = convertFormat.Convert(adoParm.Rows(0).Item("Default"), "##0")
                End If

                jx4 = default_Renamed
                If Trim(Mid(temp, col2, 4)) = "18" And Trim(Mid(temp, col1, 3)) = "1" Then
                    jx4 = 9
                    Temp2 = 9
                End If
                If Not IsDBNull(adoParm.Rows(0).Item("File_Name")) AndAlso adoParm.Rows(0).Item("File_Name") <> "" Then
                    If adoParm.Rows(0).Item("File_Name") <> "Harvest" Then
                        adoCrop = getDBDataTable("SELECT * FROM " & adoParm.Rows(0).Item("File_Name") & " where swat_code = " & Mid(temp, col2, 4))
                        If adoCrop.Rows.Count = 0 Then
                            jx4 = "   "
                        Else
                            jx4 = convertFormat.Convert(adoCrop.Rows(0).Item("Apex_Code"), "###0")
                        End If
                    End If

                    Select Case Mid(temp, col1, 3)
                        Case "  1"      'planting
                            If Temp2 = 0 Then Temp2 = Val(jx4)
                            temp1 = jx4
                            Crop1310 = getDBDataTable("SELECT PPLP2,IDC From crop1310 WHERE [Numb]=" & temp1)
                            xtp1 = Int(Crop1310.Rows(0).Item("PPLP2"))
                            xmtu = Int(Crop1310.Rows(0).Item("idc"))
                            opv1 = convertFormat.Convert(System.Math.Round(Val(Mid(temp, col3, 9))), "#####0.0") '1/15/2014 Yang

                            If (opv1 = 0) Then
                                opv1 = 2000
                            End If
                        Case "  5", "  7"
                            If temp1 = "" Then
                                adoCrop = getDBDataTable("SELECT * FROM S_A_CROP where swat_code = " & Mid(temp, col2, 4))

                                If adoCrop.Rows.Count = 0 Then
                                    temp1 = ""
                                Else
                                    temp1 = convertFormat.Convert(adoCrop.Rows(0).Item("Apex_Code"), "###0")
                                End If
                            End If
                    End Select
                End If

                month1 = convertFormat.Convert(Val(Mid(temp, 1, 4)), "#0")
                day1 = convertFormat.Convert(Val(Mid(temp, col4, 2)), "#0")

                'Updated by Yang 0n 2/13/2014: For the *.mgt file which does not have the date, the *.opc should not have the date too.
                'Updates start
                If month1 = " 0" Then
                    'month1 = " 4"
                    'If Trim(adoParm.Rows(0).Item("Default")) = "623" Or Trim(adoParm.Rows(0).Item("Default")) = "451" Then month1 = " 9"
                    month1 = "  "
                End If

                If day1 = " 0" Then
                    'day1 = "15"
                    day1 = "  "
                End If
                'Updates end

                If Temp2 = 0 Then
                    Line1 = Line1 + 1
                End If

                Condition = Mid(temp, col1, 3)
                PTrac = "  0"
                pcrp = "  0"
                pmat = Space(4)

                If Initial.Version = "4.1.0" Or Initial.Version = "4.0.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Or
                   Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                    PTrac = "    0"
                    pcrp = "    0"
                    pmat = "    0"
                    default_Renamed = convertFormat.Convert(adoParm.Rows(0).Item("Default"), "####0")
                End If

                Pblank = Space(1)
                POpv1 = Space(8)
                POpv2 = Space(8)
                POpv3 = Space(8)
                POpv4 = Space(8)
                POpv5 = Space(8)
                POpv6 = Space(8)
                POpv7 = Space(8)
                Select Case Condition
                    Case "  1" 'Planting
                        'If Initial.Version = "4.0.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" And xmtu = 7 Or xmtu = 8 Then pmat = "    0"
                        If (Initial.Version = "4.0.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Or
                            Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0") _
                            And (xmtu = 7 Or xmtu = 8) Then pmat = "   50"
                        xmtu = 0
                        pcod = default_Renamed
                        pcrp = convertFormat.Convert(jx4, "##0")
                        If Val(pcrp) = 9 Then lun = 31
                        POpv1 = convertFormat.Convert(System.Math.Round(opv1, 0), "#####0.0")
                        POpv2 = convertFormat.Convert(System.Math.Round(lun * one, 0), "#####0.0")
                        POpv2_Planting = System.Math.Round(lun * one, 0)
                        POpv5 = convertFormat.Convert(System.Math.Round(xtp1, 0), "#####0.0")
                        'Updated by Yang on 2/13/2014: need to write the heat unite values in the *.opc files
                        POpv7 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 9, 6))), 3), "###0.000")
                    Case "  2", " 10" 'irrigation
                        pcod = default_Renamed
                        If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or
                            Initial.Version = "4.3.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                            POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 32, 12))), 0), "#####0.0")
                        Else
                            POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 33, 6))), 0), "#####0.0")
                        End If
                        'Updated by Yang on 2/13/2014: need to write the heat unit values in the *.opc files
                        POpv7 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 9, 6))), 3), "###0.000")
                    Case "  3"
                        If (lyr = "   0") Or (lyr = "    ") Then lyr = "  50"
                        pcod = convertFormat.Convert(jx4, "##0")
                        pcrp = convertFormat.Convert(temp1, "##0")
                        pmat = convertFormat.Convert(lyr, "###0")
                        If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "3.1.0" Or Initial.Version = "4.1.0" Or
                            Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                            POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 32, 12))), 0), "#####0.0")
                        Else
                            POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Mid(temp, 33, 6)), 0), "#####0.0")
                        End If
                        POpv2 = convertFormat.Convert(10, "######0.")
                        'Updated by Yang on 2/13/2014: need to write the heat unite values in the *.opc files
                        POpv7 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 9, 6))), 3), "###0.000")
                    Case "  4"
                        pcod = convertFormat.Convert(jx4, "##0")
                        pcrp = convertFormat.Convert(temp1, "##0")
                        pmat = convertFormat.Convert(lyr, "###0")
                        POpv2 = convertFormat.Convert(Fert, "#####0.0")
                        'Updated by Yang on 2/13/2014: need to write the heat unite values in the *.opc files
                        POpv7 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 9, 6))), 3), "###0.000")
                    Case "  5", "  7" '(5) Harvest and kill (7) Harvest Only
                        pcod = default_Renamed
                        Crop1310 = getDBDataTable("SELECT HarvestCode From crop1310 WHERE [Numb]=" & temp1)
                        If Crop1310.Rows.Count <> 0 Then
                            pcod = Int(Crop1310.Rows(0).Item("HarvestCode"))
                        End If

                        Crop1310.Dispose()
                        pcrp = convertFormat.Convert(temp1, "##0")
                        'Updated by Yang on 2/13/2014: need to write the heat unite values in the *.opc files
                        POpv7 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 9, 6))), 3), "###0.000")
                        'POpv7 = convertFormat.Convert(opv7, "####0.00")
                        POpv2 = convertFormat.Convert(POpv2_Planting, "#####0.0")
                    Case "  6" 'tillage OPERATION
                        pcod = Format(jx4, "##0")
                        pcrp = convertFormat.Convert(temp1, "##0")
                        POpv2 = Format(POpv2_Planting, "#####0.0")
                        If Trim(POpv2_Planting) = 0 Then POpv2 = convertFormat.Convert(lun * one, "#####0.0")
                        'Updated by Yang on 2/13/2014: need to write the heat unite values in the *.opc files
                        POpv7 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 9, 6))), 3), "###0.000")
                    Case "  8"
                        pcod = default_Renamed
                        pcrp = convertFormat.Convert(temp1, "##0")
                        POpv2 = convertFormat.Convert(POpv2_Planting, "#####0.0")
                        'Updated by Yang on 2/13/2014: need to write the heat unite values in the *.opc files
                        POpv7 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 9, 6))), 3), "###0.000")
                        'POpv7 = convertFormat.Convert(opv7, "####0.00")
                    Case "  9" 'Grazing
                        pcod = convertFormat.Convert(jx4, "##0")
                        pcrp = convertFormat.Convert(temp1, "##0")
                        'Updated by Yang on 2/13/2014: need to write the heat unite values in the *.opc files
                        POpv7 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 9, 6))), 3), "###0.000")
                        grazing = getDBDataTable("SELECT * FROM grazing")
                        With grazing
                            'calculate number of animals.
                            manureProduced = Mid(temp, 52, 11)
                            Select Case Mid(temp, 25, 3)
                                Case " 44"
                                    animals = manureProduced * wsa2 / 54.7
                                    manureProduced = 54.7
                                Case " 45"
                                    animals = manureProduced * wsa2 / 26.3
                                    manureProduced = 26.3
                                Case " 47"
                                    animals = manureProduced * wsa2 / 7.6
                                    manureProduced = 7.6
                                Case " 50"
                                    animals = manureProduced * wsa2 / 34.7
                                    manureProduced = 34.7
                                Case " 52"
                                    animals = manureProduced * wsa2 / 0.2
                                    manureProduced = 0.2
                                Case Else
                                    animals = manureProduced * wsa2 / 26.3
                                    manureProduced = 26.3
                            End Select
                            If created = False Then
                                bioConsumed = Mid(temp, 32, 12)
                                If animals > 0 Then bioConsumed = Math.Round(bioConsumed * wsa2 / animals, 2)
                                animals = Math.Round(animals)
                                If .Rows.Count > 0 Then
                                    If Trim(.Rows(.Rows.Count - 1).Item("fileName")) <> Trim(mvarfilehru) Then
                                        grazingCur = getDBDataTable("SELECT * FROM grazing WHERE manureID= " & Mid(temp, 25, 3) & " AND bioConsumed= " & bioConsumed & " AND manureProduced= " & manureProduced & " AND animals= " & animals)
                                        'grazingCur.open()
                                        If animals < 1 And Initial.owners > 0 Then
                                            animals = 1
                                        End If
                                        'Dim newRow As DataRow = .NewRow
                                        'newRow.Item("fileName") = mvarfilehru
                                        'newRow.Item("manureID") = Mid(temp, 25, 3)
                                        'newRow.Item("bioConsumed") = bioConsumed
                                        'newRow.Item("ManureProduced") = manureProduced
                                        'newRow.Item("ownerID") = grazingCur.Rows(0).Item("ownerID")
                                        'newRow.Item("herdsNumber") = grazingCur.Rows(0).Item("herdsNumber")
                                        If grazingCur.Rows.Count > 0 Then ' False Then   '
                                            'newRow.Item("ownerID") = grazingCur.Rows(0).Item("ownerID")
                                            'newRow.Item("herdsNumber") = grazingCur.Rows(0).Item("herdsNumber")
                                            ownerID = grazingCur.Rows(0).Item("ownerID")
                                            herdsNumber = grazingCur.Rows(0).Item("herdsNumber")
                                        Else
                                            Initial.owners = Initial.owners + 1
                                            tempHerds = Initial.owners
                                            o = 0
                                            Do While tempHerds > 0
                                                p = p + 1
                                                tempHerds = tempHerds - 10
                                                o = o + 1
                                            Loop
                                            ownerID = o 'newRow.Item("ownerID") = o
                                            Initial.herds = Initial.herds + 1
                                            If Initial.herds > 10 Then Initial.herds = 1
                                            herdsNumber = Initial.herds 'newRow.Item("herdsNumber") = Initial.herds
                                            Call printHerd(animals, Mid(temp, 25, 3), bioConsumed, manureProduced, o)
                                        End If
                                        'newRow.Item("animals") = animals
                                        'newRow.Item("herdsNumber") = grazingCur.Rows(0).Item("herdsNumber")
                                        sqltemp = "INSERT INTO grazing (fileName, manureID, bioConsumed, ManureProduced, ownerID, herdsNumber, Animals) VALUES ('" & mvarfilehru & "', " & Mid(temp, 25, 3) & "," & bioConsumed & ", " & manureProduced & "," & ownerID & ", " & herdsNumber & ", " & animals & ")"
                                        ' sqltemp1 = " FIELDS(fileName, manureID, bioConsumed, ManureProduced, ownerID, herdsNumber)"
                                        ' sqltemp2 = "VALUES('" & mvarfilehru & "', " & Mid(temp, 25, 3) & "," & bioConsumed & ", " & manureProduced & "," & ownerID & ", " & herdsNumber & " )"
                                        modifyRecords(sqltemp)
                                        grazingCur.Dispose()
                                    End If
                                Else
                                    Initial.owners = Initial.owners + 1
                                    ownerID = Initial.owners
                                    If animals < 1 And Initial.owners > 0 Then
                                        animals = 1
                                    End If
                                    If Initial.owners = 0 Then manureProduced = 0
                                    'Dim newRow As DataRow = .NewRow
                                    ' newRow.Item("fileName") = mvarfilehru
                                    ' newRow.Item("manureID") = Mid(temp, 25, 3)
                                    ' newRow.Item("bioConsumed") = bioConsumed
                                    ' newRow.Item("ManureProduced") = manureProduced
                                    'newRow.Item("ownerID") = Initial.owners

                                    Initial.herds = Initial.herds + 1
                                    herdsNumber = Initial.herds
                                    If Initial.herds > 10 Then Initial.herds = 1
                                    'newRow.Item("herdsNumber") = Initial.herds
                                    'newRow.Item("animals") = animals
                                    sqltemp = "INSERT INTO grazing (fileName, manureID, bioConsumed, ManureProduced, ownerID, herdsNumber, Animals) VALUES ('" & mvarfilehru & "', " & Mid(temp, 25, 3) & "," & bioConsumed & ", " & manureProduced & "," & ownerID & ", " & herdsNumber & ", " & animals & ")"
                                    'sqltemp1 = " Fields(fileName, manureID, bioConsumed, ManureProduced, ownerID, herdsNumber)"
                                    'sqltemp2 = "VALUES('" & mvarfilehru & "', " & Mid(temp, 25, 3) & "," & bioConsumed & ", " & manureProduced & "," & ownerID & ", " & herdsNumber & ", )"
                                    modifyRecords(sqltemp)
                                    grazingCur.Dispose()

                                    '.Rows.Add(newRow) '.Update()
                                    Call printHerd(animals, Mid(temp, 25, 3), bioConsumed, manureProduced, Initial.owners)
                                End If
                                created = True
                            End If

                        End With

                    Case " 11"
                        GoTo continue_Renamed
                        'Autofertilization operation - Pending
                    Case " 13"
                        GoTo continue_Renamed
                        'Release operation - Pending
                    Case Else
                        If (lyr = "   0") Or (lyr = "    ") Then lyr = "  50"

                        pcod = convertFormat.Convert(jx4, "##0")
                        pcrp = convertFormat.Convert(temp1, "##0")
                        pmat = convertFormat.Convert(lyr, "###0")
                        If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or
                            Initial.Version = "4.3.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                            POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 32, 12))), 0), "#####0.0") '1/15/2014 Yang
                        Else
                            POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 33, 6))), 0), "#####0.0") '1/15/2014 Yang
                        End If
                        'Updated by Yang on 2/13/2014: need to write the heat unite values in the *.opc files
                        POpv7 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 9, 6))), 3), "###0.000")
                End Select

                If flag = False Then
                    If CDbl(Val(POpv2)) < 0 Then '1/15/2014 Yang
                        flag = True
                    End If
                Else
                    POpv2 = Space(8)
                End If

                If Initial.Version = "4.1.0" Or Initial.Version = "4.0.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
                    Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "1.1.0" Then
                    If Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "1.1.0" Then
                        RecOper(k) = convertFormat.Convert(CInt(year1), "##0") & month1 & " " & day1 & " " & convertFormat.Convert(CInt(pcod), "####0") & convertFormat.Convert(PTrac, "####0") & convertFormat.Convert(CInt(pcrp), "####0") & convertFormat.Convert(pmat, "####0") & POpv1 & POpv2 & POpv3 & POpv4 & POpv5 & POpv6 & POpv7
                    Else
                        RecOper(k) = convertFormat.Convert(CInt(year1), "#0") & month1 & day1 & convertFormat.Convert(CInt(pcod), "####0") & convertFormat.Convert(PTrac, "####0") & convertFormat.Convert(CInt(pcrp), "####0") & convertFormat.Convert(pmat, "####0") & POpv1 & POpv2 & POpv3 & POpv4 & POpv5 & POpv6 & POpv7
                    End If
                    If Condition = "  9" Then 'if it is grazing add stop grazing operation calculating the date from days in SWAT
                        If Temp2 = 0 Then
                            Line1 = Line1 + 1
                        End If
                        date1 = System.DateTime.FromOADate(0)
                        If day1 = "  " Then day1 = 1
                        If month1 = "  " Then month1 = 1
                        date1 = DateAdd(Microsoft.VisualBasic.DateInterval.Day, day1 + 1, date1)
                        date1 = DateAdd(Microsoft.VisualBasic.DateInterval.Month, month1 - 1, date1)
                        date1 = DateAdd(Microsoft.VisualBasic.DateInterval.Year, year1 - 1, date1)
                        date1 = DateAdd(Microsoft.VisualBasic.DateInterval.Day, CDbl(Val(Mid(temp, 20, 4))), date1)
                        month2 = convertFormat.Convert(Month(date1), "#0")
                        day2 = convertFormat.Convert(VB.Day(date1), "#0")
                        year2 = Year(date1) - 1900 + 1
                        k = k + 1
                        ReDim Preserve RecOper(k)
                        If Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "1.1.0" Then
                            RecOper(k) = convertFormat.Convert(year2, "#0") & month2 & " " & day2 & " " & convertFormat.Convert(427, "####0") & convertFormat.Convert(PTrac, "####0") & convertFormat.Convert(pcrp, "####0") & convertFormat.Convert(pmat, "####0") & POpv1 & POpv2 & POpv3 & POpv4 & POpv5 & POpv6 & POpv7
                        Else
                            RecOper(k) = convertFormat.Convert(year2, "#0") & month2 & day2 & convertFormat.Convert(427, "####0") & convertFormat.Convert(PTrac, "####0") & convertFormat.Convert(pcrp, "####0") & convertFormat.Convert(pmat, "####0") & POpv1 & POpv2 & POpv3 & POpv4 & POpv5 & POpv6 & POpv7
                        End If
                    End If
                Else
                    RecOper(k) = convertFormat.Convert(year1, "#0") & month1 & day1 & pcod & PTrac & pcrp & pmat & Pblank & POpv1 & POpv2 & POpv3 & POpv4 & POpv5 & POpv6 & POpv7
                    If Condition = "  9" Then 'if it is grazing add stop grazing operation calculating the date from days in SWAT
                        date1 = year1 + month1 + day1
                        date1 = DateAdd(Microsoft.VisualBasic.DateInterval.Day, day1 + 1, date1)
                        date1 = DateAdd(Microsoft.VisualBasic.DateInterval.Month, month1 - 1, date1)
                        date1 = DateAdd(Microsoft.VisualBasic.DateInterval.Year, year1 - 1, date1)
                        date1 = DateAdd(Microsoft.VisualBasic.DateInterval.Day, CDbl(Val(Mid(temp, 20, 4))), date1)
                        month2 = Month(date1)
                        day2 = VB.Day(date1)
                        year2 = Year(date1) - 1900
                        k = k + 1
                        ReDim Preserve RecOper(k)
                        RecOper(k) = convertFormat.Convert(year2, "#0") & month2 & day2 & convertFormat.Convert(427, "####0") & convertFormat.Convert(PTrac, "####0") & convertFormat.Convert(pcrp, "####0") & convertFormat.Convert(pmat, "####0") & POpv1 & POpv2 & POpv3 & POpv4 & POpv5 & POpv6 & POpv7
                    End If
                End If

                k = k + 1
                ReDim Preserve RecOper(k)
            End If

continue_Renamed:

        Loop ' do loop ends here.
        RecOper(k) = " "
        If Not grazing Is Nothing Then grazing.Dispose()

        For k = 0 To Line1 - 3
            'If Initial.Version = "4.1.0" Or Initial.Version = "4.0.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
            '    Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
            '    value = Left(RecOper(k), 16) & convertFormat.Convert(Temp2, "####0") & Mid(RecOper(k), 22, 90)
            'Else
            '    value = Left(RecOper(k), 12) & convertFormat.Convert(Temp2, "##0") & Mid(RecOper(k), 16, 61)
            'End If
            Select Case Initial.Version
                Case "4.1.0", "4.0.0", "4.2.0", "4.3.0"
                    value = Left(RecOper(k), 16) & convertFormat.Convert(Temp2, "####0") & Mid(RecOper(k), 22, 90)
                Case "1.1.0", "1.2.0", "1.3.0"
                    value = Left(RecOper(k), 19) & convertFormat.Convert(Temp2, "####0") & Mid(RecOper(k), 25, 90)
                Case Else
                    value = Left(RecOper(k), 12) & convertFormat.Convert(Temp2, "##0") & Mid(RecOper(k), 16, 61)
            End Select
            b.WriteLine(value)
        Next

        For j = k To UBound(RecOper)
            b.WriteLine(RecOper(j))
        Next

        a.Close()
        b.Close()

        Exit Sub
goError:
        MsgBox(Err.Description & " File = " & Trim(mvarfilemgt))

    End Sub

    Sub printHerd(ByRef animals As Single, ByRef manureID As Single, ByRef bioConsumed As Single, ByRef manureProduced As Single, ByRef o As Short)
        'OGM CHECK LATER
        Dim swFile As StreamWriter

        swFile = New StreamWriter(File.Open(Initial.Output_files & "\" & Initial.herd, FileMode.Append))
        convertFormat = New Convertion

        swFile.Write(convertFormat.Convert(o, "###0"))
        swFile.Write(convertFormat.Convert(animals, "####0.00"))
        swFile.Write(convertFormat.Convert(manureID, "####0.00"))
        swFile.Write(convertFormat.Convert(0, "####0.00"))
        swFile.Write(convertFormat.Convert(bioConsumed, "####0.00"))
        swFile.Write(convertFormat.Convert(manureProduced, "####0.00"))
        swFile.WriteLine(convertFormat.Convert(0, "####0.00"))
        swFile.Close()
        swFile.Dispose()
    End Sub
    Public Sub Site()
        Dim sitpos As Object
        Dim totalare As Single
        Dim temp As Object
        Dim current_line As Integer
        Dim swFile As StreamWriter
        Dim srFile As StreamReader
        Dim ADORecordset As DataTable
        Dim TakeField As Convertion
        Dim i As Integer

        On Error GoTo goError

        TakeField = New Convertion
        ADORecordset = New DataTable
        ADORecordset = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'Site' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")

        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & mvarfilesub))
        swFile = New StreamWriter(File.Create(Initial.Output_files & "\" & mvarfilesit))
        convertFormat = New Convertion

        With ADORecordset
            current_line = .Rows(0).Item("Line")
            For i = 0 To .Rows.Count - 1
                If Not IsDBNull(.Rows(i).Item("SwatFile")) AndAlso .Rows(i).Item("SwatFile") <> "" Then
                    Select Case .Rows(i).Item("SwatFile")
                        Case "*.sub"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfileName)
                        Case "*.hru"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilehru)
                        Case "basins.bsn"
                            TakeField.filename = Initial.Input_Files & "\" & "basins.bsn"
                        Case "*.wgn"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilewgn)
                    End Select

                    TakeField.Leng = .Rows(i).Item("Leng")
                    TakeField.LineNum = .Rows(i).Item("Lines")
                    TakeField.Inicia = .Rows(i).Item("Inicia")
                    value = TakeField.value()
                Else
                    If .Rows(i).Item("Value") = "Blank" Then
                        value = " "
                    Else
                        value = .Rows(i).Item("Value")
                    End If
                End If

                temp = .Rows(i).Item("Line") & .Rows(i).Item("Field")

                Select Case temp
                    Case "00"
                        totalare = Val(value)
                End Select

                If .Rows(i).Item("Field") = 5 Then
                    value = convertFormat.Convert(System.Math.Round(Val(value), 2), "####0.00")
                End If

                If Not IsDBNull(.Rows(i).Item("Format")) AndAlso .Rows(i).Item("Format") <> "" Then
                    lenFormat = Len(.Rows(i).Item("Format"))
                    roundformat = Right(Trim(.Rows(i).Item("Format")), 1)
                    NumberFormat = Left(Trim(.Rows(i).Item("Format")), lenFormat - 2)
                    value = convertFormat.Convert(System.Math.Round(Val(value), roundformat), NumberFormat)
                End If

                If temp <> "00" Then
                    If i < .Rows.Count - 1 Then
                        If current_line = .Rows(i + 1).Item("Line") Then
                            swFile.Write(value)
                        Else
                            swFile.WriteLine(value)
                        End If
                        current_line = .Rows(i + 1).Item("Line")
                    End If
                Else
                    current_line = .Rows(i + 1).Item("Line")
                End If
            Next

            If Initial.Version <> "4.0.0" And Initial.Version <> "4.1.0" And Initial.Version <> "4.2.0" And Initial.Version <> "4.3.0" And Initial.Version <> "1.1.0" And Initial.Version <> "1.2.0" And Initial.Version <> "1.3.0" Then
                sitpos = InStr(1, mvarfilesit, ".")
                value = Left(mvarfilesit, sitpos - 1) & ".wth"
                If CDbl(Val(Variables.pcpgages)) = 0 Then value = ""
                swFile.WriteLine(value)
                value = " "
                swFile.WriteLine(value)
                swFile.WriteLine(value)
            Else
                value = " "
                swFile.WriteLine(value)
                swFile.WriteLine(value)
            End If
        End With

        srFile.Close()
        srFile.Dispose()
        srFile = Nothing

        swFile.Close()
        swFile.Dispose()
        swFile = Nothing

        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub

    Public Sub Weather()
        Dim bis As Object
        Dim tmp1 As Object
        Dim j, i As Integer
        Dim c As Object = Nothing
        Dim b As Object = Nothing
        Dim a As Object = Nothing
        Dim fs As Object = Nothing
        'Dim adoConnection As ADODB.Connection
        Dim ADORecordset As DataTable
        Dim year1, k, day_Renamed As Object
        Dim month_Renamed As Short
        Dim tmp As Object
        Dim pcp As String
        Dim Months, monthj As Object
        Dim daya As Object

        On Error GoTo goError

        ADORecordset = New DataTable
        'adoConnection = New ADODB.Connection
        'adoConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir1 & "\Project_Parameters.mdb"
        'adoConnection.Open()
        convertFormat = New Convertion

        ADORecordset = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'Weather' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")
        fs = CreateObject("Scripting.FileSystemObject")

        a = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(Initial.prpfiles(1)))
        If Initial.ngn > 1 Then b = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(Initial.temfiles(1)))

        c = fs.CreateTextFile(Initial.Output_files & "\" & mvarfilewht)
        Months = New Object() {0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365}
        monthj = New Object() {0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366}

        With ADORecordset
            '.MoveFirst()
            For i = 0 To .Rows.Count - 1
                'Do While .EOF <> True
                If (.Rows(i).Item("SwatFile") <> "") Then
                    Select Case .Rows(i).Item("SwatFile")
                        Case "*.sub"
                            convertFormat.filename = Initial.Input_Files & "\" & mvarfileName
                    End Select
                    convertFormat.Leng = .Rows(i).Item("Leng")
                    convertFormat.LineNum = .Rows(i).Item("Lines")
                    convertFormat.Inicia = .Rows(i).Item("Inicia")
                    value = convertFormat.value()
                Else
                    If .Rows(i).Item("Value") = "Blank" Then
                        value = " "
                    Else
                        value = .Rows(i).Item("Value")
                    End If
                End If
                If Not IsDBNull(.Rows(i).Item("Format")) AndAlso .Rows(i).Item("Format") <> "" Then
                    lenFormat = Len(.Rows(i).Item("Format"))
                    roundformat = Right(Trim(.Rows(i).Item("Format")), 1)
                    NumberFormat = Left(Trim(.Rows(i).Item("Format")), lenFormat - 2)
                    value = convertFormat.Convert(System.Math.Round(Val(value), roundformat), NumberFormat)
                End If

                Select Case .Rows(i).Item("Line")
                    Case 98
                        Last = value
                    Case 99
                        Lasttmp = value
                End Select
                '.MoveNext()
            Next
        End With

        a.ReadLine()
        a.ReadLine()
        a.ReadLine()
        a.ReadLine()
        If Initial.ngn > 1 Then
            b.ReadLine()
            b.ReadLine()
            b.ReadLine()
            b.ReadLine()
        End If

        Do While Val(Last) > Val(Variables.pcpgages)
            Last = Last - Variables.pcpgages
        Loop
        i = 5 * (Last - 1) + 8

        Do While Val(Lasttmp) > Val(Variables.tmpgages)
            Lasttmp = Lasttmp - Variables.tmpgages
        Loop

        j = 10 * (Lasttmp - 1) + 8
        If j <= 0 Then j = 1
        Do While a.atEndOfStream <> True
            If Initial.ngn > 1 Then
                tmp = b.ReadLine
            Else
                tmp = "000000000"
            End If
            tmp1 = a.ReadLine
            year1 = Left(tmp1, 4)
            daya = Mid(tmp1, 5, 3)
            day_Renamed = Int(CDbl(Val(daya)))
            bis = year1 Mod 4

            If (bis <> 0) Then
                For k = 1 To 12
                    If (day_Renamed <= Months(k)) Then
                        month_Renamed = k
                        If (k > 1) Then
                            day_Renamed = day_Renamed - Months(k - 1)
                        End If
                        Exit For
                    End If
                Next
            Else
                For k = 1 To 12
                    If (day_Renamed <= monthj(k)) Then
                        month_Renamed = k
                        If (k > 1) Then
                            day_Renamed = day_Renamed - monthj(k - 1)
                        End If
                        Exit For
                    End If
                Next
            End If
            c.Write("  ")
            c.Write(year1)
            c.Write(convertFormat.Convert(month_Renamed, "###0"))
            c.Write(convertFormat.Convert(day_Renamed, "###0"))
            c.Write("     0")
            c.Write(convertFormat.Convert(Mid(tmp, j, 5), "##0.00"))
            c.Write(convertFormat.Convert(Mid(tmp, j + 5, 5), "##0.00"))
            c.WriteLine(convertFormat.Convert(Mid(tmp1, i, 5), "##0.00"))
        Loop

        a.Close()
        If Initial.ngn > 0 Then b.Close()
        c.Close()
        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub

    Public Sub Weather1()
        Dim bis As Object
        Dim e As Object = Nothing
        Dim d As Object = Nothing
        Dim c As Object = Nothing
        Dim z As Object = Nothing
        Dim X As Object = Nothing
        Dim b As Object = Nothing
        Dim a As Object = Nothing
        Dim fs As Object = Nothing
        Dim ADORecordset As DataTable
        Dim year1, h, k, i, j, r, w, day_Renamed As Integer
        Dim l As Short
        Dim month_Renamed As Short
        Dim tmp As Object
        Dim pcp As String
        Dim Months, monthj As UShort()
        Dim daya As Object
        Dim ngnMod As Object
        Dim ngnTemp As Short
        Dim slr, pcp1, wnd As Integer
        Dim tmp1 As String
        Dim hmd As Short
        Dim slr1, wnd1 As Object
        Dim hmd1 As String

        Try
            fs = CreateObject("Scripting.FileSystemObject")
            ADORecordset = New DataTable

            convertFormat = New Convertion

            ADORecordset = getDBDataTable("SELECT * FROM sub_included ORDER BY Subbasin")

            Months = New UShort() {0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365}
            monthj = New UShort() {0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366}
            a = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(Initial.prpfiles(1)))
            If Initial.ngn > 1 Then b = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(Initial.temfiles(1)))
            X = fs.CreateTextFile(Initial.Output_files & "\Wdlstcom.dat")
            pcp = CStr(0)
            tmp = 0
            slr = 0
            wnd = 0
            hmd = 0
            ngnTemp = ngn

            Do While ngnTemp > 0
                ngnMod = ngnTemp Mod 10
                ngnTemp = ngnTemp / 10
                Select Case ngnMod
                    Case 1
                        pcp1 = 1
                    Case 2
                        tmp1 = 1
                    Case 3
                        slr = 1
                    Case 4
                        wnd = 1
                    Case 5
                        hmd = 1
                End Select
            Loop

            With ADORecordset
                Initial.File_Number = 0
                For l = 0 To .Rows.Count - 1
                    'Wait_Form.Pbar_Scenarios.Value = (1 - (.Rows.Count - l) / .Rows.Count) * 100
                    'If .Rows(l).Item("File_Number") > 0 Then Initial.File_Number = .Rows(l).Item("File_Number")
                    Initial.File_Number += 1
                    If .Rows(l).Item("File_Number") <> 0 Then
                        z = fs.CreateTextFile(Initial.Output_files & "\" & Left(.Rows(l).Item("Subbasin"), 10) & "wth")
                        X.WriteLine(convertFormat.Convert(Initial.File_Number, "###0") & "  " & Left(.Rows(l).Item("Subbasin"), 10) & "wth")
                        a = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(Initial.prpfiles(1)))
                        If Initial.ngn > 1 Then b = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(Initial.temfiles(1)))
                        If slr = 1 Then c = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(Initial.slrfiles))
                        If hmd = 1 Then d = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(Initial.hmdfiles))
                        If wnd = 1 Then e = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(Initial.wndfiles))

                        a.ReadLine()
                        a.ReadLine()
                        a.ReadLine()
                        a.ReadLine()

                        If Initial.ngn > 1 Then
                            b.ReadLine()
                            b.ReadLine()
                            b.ReadLine()
                            b.ReadLine()
                        End If

                        If slr = 1 Then c.ReadLine()
                        If hmd = 1 Then d.ReadLine()
                        If wnd = 1 Then e.ReadLine()

                        j = 5 * (.Rows(l).Item("pcpNumber") - 1) + 8
                        i = 10 * (.Rows(l).Item("tmpNumber") - 1) + 8
                        r = 8 * (.Rows(l).Item("slrNumber") - 1) + 8
                        If r <= 0 Then r = 1
                        h = 8 * (.Rows(l).Item("hmdNumber") - 1) + 8
                        If h <= 0 Then h = 1
                        w = 8 * (.Rows(l).Item("wndNumber") - 1) + 8
                        If w <= 0 Then w = 1

                        Do While a.atEndOfStream <> True
                            slr1 = "  -99.00"
                            wnd1 = "  -99.00"
                            hmd1 = "  -99.00"
                            If slr = 1 Then
                                If c.atEndOfStream <> True Then slr1 = Mid(c.ReadLine, r, 8)
                            End If
                            If hmd = 1 Then
                                If d.atEndOfStream <> True Then
                                    hmd1 = Mid(d.ReadLine, h, 8)
                                    If hmd1 > 0 Then hmd1 = hmd1 / 100
                                End If
                            End If
                            If wnd = 1 Then
                                If e.atEndOfStream <> True Then wnd1 = Mid(e.ReadLine, w, 8)
                            End If
                            If Initial.ngn > 1 Then
                                If b.atEndOfStream <> True Then tmp = Mid(b.ReadLine, i, 10)
                            Else
                                tmp = "  -99.00"
                            End If
                            tmp1 = a.ReadLine
                            If tmp1.Trim = "" Then
                                Exit Do
                            End If
                            year1 = Left(tmp1, 4)
                            daya = Mid(tmp1, 5, 3)
                            day_Renamed = Int(CDbl(Val(daya)))
                            bis = year1 Mod 4

                            If (bis <> 0) Then
                                For k = 1 To 12
                                    If (day_Renamed <= Months(k)) Then
                                        month_Renamed = k
                                        If (k > 1) Then
                                            day_Renamed = day_Renamed - Months(k - 1)
                                        End If
                                        Exit For
                                    End If
                                Next
                            Else
                                For k = 1 To 12
                                    If (day_Renamed <= monthj(k)) Then
                                        month_Renamed = k
                                        If (k > 1) Then
                                            day_Renamed = day_Renamed - monthj(k - 1)
                                        End If
                                        Exit For
                                    End If
                                Next
                            End If
                            z.Write("  ")
                            z.Write(year1)
                            z.Write(convertFormat.Convert(month_Renamed, "###0"))
                            z.Write(convertFormat.Convert(day_Renamed, "###0"))
                            z.Write(convertFormat.Convert(slr1, "##0.00"))
                            z.Write(convertFormat.Convert(Mid(tmp, 1, 5), "##0.00"))
                            z.Write(convertFormat.Convert(Mid(tmp, 1 + 5, 5), "##0.00"))
                            z.Write(convertFormat.Convert(Mid(tmp1, j, 5), "##0.00"))
                            z.Write(convertFormat.Convert(hmd1, "##0.00"))
                            z.WriteLine(convertFormat.Convert(wnd1, "##0.00"))
                        Loop

                        a.Close()
                        If Initial.ngn > 1 Then b.Close()
                        z.Close()
                    Else
                        X.WriteLine(convertFormat.Convert(Initial.File_Number + 1, "###0") & "  " & Left(.Rows(l).Item("Subbasin"), 10) & "wth")
                    End If
                Next
            End With

            X.Close()
            Exit Sub

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            If Not X Is Nothing Then
                X.Close()
                X = Nothing
            End If
            If Not e Is Nothing Then
                e.Close()
                e = Nothing
            End If
            If Not d Is Nothing Then
                d.Close()
                d = Nothing
            End If
            If Not c Is Nothing Then
                c.Close()
                c = Nothing
            End If
            If Not z Is Nothing Then
                z.Close()
                z = Nothing
            End If
            If Not b Is Nothing Then
                b.Close()
                b = Nothing
            End If
            If Not a Is Nothing Then
                a.Close()
                a = Nothing
            End If
        End Try
    End Sub
    Function lookatparm(ByRef oper As String) As String

        'Dim adoConnection As ADODB.Connection
        Dim ADORecordset As DataTable

        On Error GoTo goError

        'adoConnection = New ADODB.Connection
        'adoConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir1 & "\Project_Parameters.mdb"
        'adoConnection.Open()

        ADORecordset = getDBDataTable("SELECT * FROM Parmopc where code = " & oper)
        'ADORecordset.MoveFirst()
        If ADORecordset.Rows.Count = 0 Then
            lookatparm = "    "
        Else
            lookatparm = convertFormat.Convert(ADORecordset.Rows(0).Item("Field"), "###0")
        End If

        ADORecordset.Dispose()
        ADORecordset = Nothing

        Exit Function
goError:
        MsgBox(Err.Description)

    End Function

    Function lookatfile(ByRef id As String, ByRef oper As String) As String

        'Dim adoConnection As ADODB.Connection
        Dim ADORecordset As DataTable
        Dim adoRecordSet1 As DataTable

        On Error GoTo goError

        'adoConnection = New ADODB.Connection
        'adoConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Dir1 & "\Project_Parameters.mdb"
        '		adoConnection.Open()

        ADORecordset = getDBDataTable("SELECT * FROM Parmopc where code = " & oper)
        'ADORecordset.MoveFirst()
        adoRecordSet1 = getDBDataTable("SELECT * FROM " & ADORecordset.Rows(0).Item("Line") & " where swat_code = " & id)
        If adoRecordSet1.Rows.Count = 0 Then
            lookatfile = "    "
        Else
            lookatfile = convertFormat.Convert(adoRecordSet1.Rows(0).Item("Line"), "###0")
        End If

        ADORecordset.Dispose()
        adoRecordSet1.Dispose()
        ADORecordset = Nothing
        adoRecordSet1 = Nothing
        Exit Function
goError:
        MsgBox(Err.Description)
    End Function
    Public Sub Swat()
        Dim temp As Object
        Dim Temp2 As Object
        Dim temp1 As Object
        Dim d As Object
        Dim c As Object
        Dim filenm As Object
        Dim File1 As Object
        Dim b As Object
        Dim fs As Object
        'Dim adoConnection As ADODB.Connection
        Dim ADORecordset As DataTable
        Dim filenum(2) As Object
        Dim swatf() As String
        Dim foundit As Object
        Dim flag As Boolean
        Dim i, j As Short

        On Error GoTo goError
        'adoConnection = New ADODB.Connection
        'adoConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Initial.Dir1 & "\Project_Parameters.mdb"
        'adoConnection.Open()
        convertFormat = New Convertion

        ADORecordset = New DataTable
        ADORecordset = getDBDataTable("Output_Swt")

        fs = CreateObject("Scripting.FileSystemObject")
        b = fs.OpenTextFile(mvarfilewht & "\" & mvarfileName)
        b.ReadLine()
        b.ReadLine()
        b.ReadLine()
        b.ReadLine()
        b.ReadLine()
        value = b.ReadLine
        b.Close()
        filenum(1) = Mid(value, 13, 1)
        filenum(2) = Mid(value, 14, 1)

        If filenum(1) = "0" Then
            File1 = " " & " " & filenum(2) & "P"
        Else
            File1 = " " & filenum(1) & filenum(2) & "P"
        End If

        foundit = False

        For j = 0 To ADORecordset.Rows.Count - 1
            'Do While ADORecordset.EOF <> True
            If File1 = ADORecordset.Rows(j).Item(0).Value Then
                foundit = True
                Exit For
            Else
                foundit = False
            End If
            'ADORecordset.MoveNext()
        Next

        i = i + 1
        filenm = File1 & ".dat"

        If foundit = True Then
            b = fs.OpenTextFile(mvarfilewht & "\" & mvarfileName)
            c = fs.OpenTextFile(mvarfilewht & "\" & File1)
            d = fs.CreateTextFile(mvarfilewht & "\" & "temp")

            For i = 1 To 6
                d.WriteLine(c.ReadLine)
            Next

            For i = 1 To 9
                b.ReadLine()
            Next

            Do While c.atEndOfStream <> True Or b.atEndOfStream <> True
                temp1 = b.ReadLine
                Temp2 = c.ReadLine
                d.Write(Mid(Temp2, 1, 10))
                For i = 11 To 95 Step 17
                    d.Write(convertFormat.Convert(Val(Mid(temp1, i, 17)) + Val(Mid(Temp2, i, 17)), " 0.0000000000E+00"))
                Next
                d.WriteLine(convertFormat.Convert(Val(Mid(temp1, 96, 17)) + Val(Mid(Temp2, 96, 17)), " 0.0000000000E+00"))
            Loop
            b.Close()
            c.Close()
            d.Close()
            c = fs.CreateTextFile(mvarfilewht & "\" & File1)
            d = fs.OpenTextFile(mvarfilewht & "\" & "temp")
            Do While d.atEndOfStream <> True
                c.WriteLine(d.ReadLine)
            Loop
        Else
            modifyRecords("INSERT INTO Output_Swt (Name) VALUES('" & File1 & "')")
            'ADORecordset.AddNew()
            'ADORecordset.Fields(0).Value = File1
            'ADORecordset.Update()
            b = fs.OpenTextFile(mvarfilewht & "\" & mvarfileName)
            c = fs.CreateTextFile(mvarfilewht & "\" & File1)
            c.WriteLine(b.ReadLine) 'line 1
            c.WriteLine(b.ReadLine) 'line 2
            c.WriteLine(b.ReadLine) 'line 3
            b.ReadLine()
            b.ReadLine()
            c.WriteLine(b.ReadLine) ' line 4
            b.ReadLine()
            c.WriteLine(b.ReadLine) ' line 5
            c.WriteLine(b.ReadLine) ' line 6
            temp = Mid(b.ReadLine, 10, 103)
            c.WriteLine("   0    0" & temp) ' line 7
            Do While b.atEndOfStream <> True
                c.WriteLine(b.ReadLine) ' others lines
            Loop
        End If

        b.Close()
        c.Close()
        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub

    Public Sub wmpfiles()
        Dim current_line As Integer
        Dim swFile As StreamWriter
        Dim srFile As StreamReader
        Dim ADORecordset As DataTable
        Dim TakeField As Convertion
        Dim i As Integer

        On Error GoTo goError
        ADORecordset = New DataTable
        TakeField = New Convertion

        ADORecordset = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'Wpmfiles' AND version =" & "'" & Initial.Version & "'" & " ORDER BY line, field")
        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & Trim(mvarfilewgn)))
        swFile = New StreamWriter(File.Create(Initial.Output_files & "\" & Trim(mvarfilewp1)))
        convertFormat = New Convertion

        With ADORecordset
            current_line = .Rows(0).Item("Line")
            For i = 0 To .Rows.Count - 1
                If Not IsDBNull(.Rows(i).Item("SwatFile")) AndAlso .Rows(i).Item("SwatFile") <> "" Then
                    Select Case .Rows(i).Item("SwatFile")
                        Case "*.wgn"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(mvarfilewgn)
                    End Select
                    TakeField.Leng = .Rows(i).Item("Leng")
                    TakeField.LineNum = .Rows(i).Item("Lines")
                    TakeField.Inicia = .Rows(i).Item("Inicia")
                    value = TakeField.value()
                Else
                    If .Rows(i).Item("Value") = "Blank" Then
                        value = " "
                    Else
                        value = .Rows(i).Item("Value")
                    End If
                End If

                If Not IsDBNull(.Rows(i).Item("Format")) AndAlso .Rows(i).Item("Format") <> "" Then
                    lenFormat = Len(.Rows(i).Item("Format"))
                    roundformat = Right(Trim(.Rows(i).Item("Format")), 1)
                    NumberFormat = Left(Trim(.Rows(i).Item("Format")), lenFormat - 2)
                    value = convertFormat.Convert(System.Math.Round(CSng(value), roundformat), NumberFormat)
                End If

                If i < .Rows.Count - 1 Then
                    If current_line = .Rows(i + 1).Item("Line") Then
                        swFile.Write(value)
                    Else
                        swFile.WriteLine(value)
                    End If
                    current_line = .Rows(i + 1).Item("Line")
                Else
                    swFile.WriteLine(value)
                End If
            Next
        End With

        srFile.Close()
        srFile.Dispose()
        srFile = Nothing

        swFile.Close()
        swFile.Dispose()
        swFile = Nothing

        Exit Sub
goError:
        MsgBox(Err.Description)
    End Sub

    Public Sub SWATSubbasins(ByRef SubbasinNum As String, rotation As Integer, myConnection As OleDb.OleDbConnection)
        Dim tillName, cropName As String
        Dim Pmat1 As Object
        Dim lun As Single = 0
        Dim pcod As Object
        Dim item9, item8, item7, item6, item5, item4, item3, item2, item1 As String
        Dim POpv7, POpv6, POpv5, POpv4, POpv3, POpv2, POpv1, Pblank As String
        Dim pmat As Object
        Dim pcrp As Object
        Dim PTrac As Object
        Dim Condition As Object
        Dim day1 As Integer
        Dim month1 As Integer
        Dim opv1 As Single
        Dim xtp1 As Integer = 0
        Dim Crop1310 As DataTable
        Dim jx4 As String
        Dim lyr As Object
        Dim Fert As Object
        Dim tempopc As Object
        Dim i As Integer
        'Dim current_line As Object
        Dim temp1 As Object
        Dim Temp2 As Object
        Dim opv7 As Object
        Dim Line1 As Object
        'Dim fs As Object
        Dim currdir, FEMScenario As String
        Dim adoParm, Fertilizer, adoCrop As DataTable
        'Dim FEM As DataTable
        'Dim adotill As DataTable
        Dim col2, col1, col3, col4 As Integer
        Dim limit As Short
        Dim temp As String = String.Empty
        Dim default_Renamed As String = ""
        Dim year_Renamed As Short
        Dim k As Short
        Dim name As String = String.Empty
        Dim sSql As String()
        Dim a As StreamReader = New StreamReader(Initial.Input_Files & "\" & Trim(mvarfilemgt))
        Dim lat1, lon1 As String
        Try
            tillName = String.Empty
            cropName = String.Empty
            currdir = CurDir()
            opv1 = 0

            'TakeField = New Convertion
            convertFormat = New Convertion

            'FEM = New DataTable
            Fertilizer = New DataTable
            adoParm = New DataTable
            adoCrop = New DataTable

            Select Case Initial.Version
                Case "1.0.0"
                    col1 = 18
                    limit = 2
                    col2 = 29
                    col3 = 21
                    col4 = 7
                Case "1.1.0"
                    col1 = 16
                    limit = 30
                    col2 = 20
                    col3 = 33
                    col4 = 5
                Case "1.2.0"
                    col1 = 16
                    limit = 30
                    col2 = 20
                    col3 = 33
                    col4 = 5
                Case "1.3.0"
                    col1 = 16
                    limit = 30
                    col2 = 20
                    col3 = 33
                    col4 = 5
                Case "2.0.0"
                    col1 = 18
                    limit = 2
                    col2 = 29
                    col3 = 21
                    col4 = 7
                Case "2.1.0"
                    col1 = 16
                    limit = 30
                    col2 = 20
                    col3 = 33
                    col4 = 5
                Case "2.3.0"
                    col1 = 16
                    limit = 30
                    col2 = 20
                    col3 = 33
                    col4 = 5
                Case "3.0.0"
                    col1 = 18
                    limit = 2
                    col2 = 29
                    col3 = 21
                    col4 = 7
                Case "3.1.0"
                    col1 = 16
                    limit = 30
                    col2 = 20
                    col3 = 33
                    col4 = 4
                Case "4.0.0"
                    col1 = 18
                    limit = 2
                    col2 = 29
                    col3 = 21
                    col4 = 7
                Case "4.1.0"
                    col1 = 16
                    limit = 30
                    col2 = 20
                    col3 = 33
                    col4 = 5
                Case "4.2.0"
                    col1 = 16
                    limit = 30
                    col2 = 20
                    col3 = 33
                    col4 = 5
                Case "4.3.0"
                    col1 = 16
                    limit = 30
                    col2 = 20
                    col3 = 33
                    col4 = 5
            End Select

            k = 0
            Dim RecOper(k) As Object
            Line1 = 3
            flag = False
            opv7 = 0.0#
            Temp2 = Nothing
            temp1 = Nothing


            'FEM = New DataTable
            'a = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(mvarfilemgt))
            For i = 1 To limit
                a.ReadLine()
            Next

            year_Renamed = 1
            tempopc = ""

            '/* delete records added before regarding this project and operation file */
            'FEM = getLocalDataTable("SELECT * From fem", Initial.Output_files)
            modifyLocalRecords("DELETE * FROM fem WHERE Composite = '" & mvarfilemgt & "'", Initial.Output_files)
            i = 0
            'adoParm = getDBDataTable("SELECT * FROM Parmopc")
            adoParm = getDBDataTableNoCon("SELECT * FROM Parmopc", myConnection)
            Dim j As UShort
            Dim adoParm_file As String = String.Empty
            Do While a.EndOfStream <> True
                If tempopc = "  5" Then
                    temp = Left(temp, col1 - 1) & "  8"
                    tempopc = ""
                Else
                    temp = a.ReadLine
                    tempopc = Mid(temp, col1, 3)
                End If
                Fert = Val(Mid(temp, 33, 8))

                If Mid(temp, col1, 3) = "  5" Then temp = Left(temp, col1 - 1) & "  7"

                If (Mid(temp, col1, 3) = "  3" Or Mid(temp, col1, 3) = "  4" Or Mid(temp, col1, 3) = " 11") Then
                    lyr = Mid(temp, col2, 4)
                Else
                    lyr = "   0"
                End If

                If (Trim(Mid(temp, col1, 3)) = "" Or (Mid(temp, col1, 3)).Trim = "0" Or (Mid(temp, col1, 3)).Trim = "17") Then
                    year_Renamed = year_Renamed + 1
                Else
                    'adoParm = getDBDataTable("SELECT * FROM Parmopc where code = " & Mid(temp, col1, 3))
                    default_Renamed = "   "
                    adoParm_file = String.Empty
                    For j = 0 To adoParm.Rows.Count - 1
                        If adoParm.Rows(j).Item("code") = Mid(temp, col1, 3) Then
                            default_Renamed = convertFormat.Convert(adoParm.Rows(j).Item("Default"), "##0")
                            If Not adoParm.Rows(j).Item("File_Name") Is DBNull.Value Then
                                adoParm_file = adoParm.Rows(j).Item("File_Name")
                            End If
                            Exit For
                        End If
                    Next

                    jx4 = default_Renamed
                    If adoParm_file <> String.Empty Then
                        adoCrop = getDBDataTableNoCon("SELECT * FROM " & adoParm_file & " where swat_code = " & Mid(temp, col2, 4), myConnection)
                        If adoCrop.Rows.Count > 0 Then
                            jx4 = convertFormat.Convert(adoCrop.Rows(0).Item("Apex_Code"), "###0")
                        Else
                            jx4 = "   "
                        End If

                        If (Mid(temp, col1, 3) = "  1") Then
                            If (IsNothing(Temp2)) Then Temp2 = jx4
                            temp1 = jx4
                            'Crop1310 = getDBDataTable("SELECT PPLP2 From crop1310 WHERE [Numb]=" & temp1)
                            Crop1310 = getDBDataTableNoCon("SELECT PPLP2 From crop1310 WHERE [Numb]=" & temp1, myConnection)
                            xtp1 = Int(Crop1310.Rows(0).Item("PPLP2"))
                            opv1 = convertFormat.Convert(System.Math.Round(Val(Mid(temp, col3, 6))), "#####0.0")

                            If (opv1 = 0) Then
                                opv1 = 2000
                            End If
                        End If
                    End If

                    month1 = convertFormat.Convert(Val(Mid(temp, 1, 4)), "#0")
                    day1 = convertFormat.Convert(Val(Mid(temp, col4, 2)), "#0")

                    If month1 = " 0" Then
                        month1 = " 4"
                    End If

                    If day1 = " 0" Then
                        day1 = "15"
                    End If

                    If IsNothing(Temp2) Or Temp2 Is System.DBNull.Value Or Temp2 = 0 Then
                        Line1 = Line1 + 1
                    End If

                    Condition = Mid(temp, col1, 3)
                    PTrac = Space(3) 'Tractor ID
                    pcrp = Space(3) 'Crop ID Number
                    pmat = Space(4) 'Id Numbers
                    Pblank = Space(1)
                    POpv1 = Space(8)
                    POpv2 = Space(8)
                    POpv3 = Space(8)
                    POpv4 = Space(8)
                    POpv5 = Space(8)
                    POpv6 = Space(8)
                    POpv7 = Space(8)
                    item1 = Space(20)
                    item2 = Space(20)
                    item3 = Space(20)
                    item4 = Space(20)
                    item5 = Space(20)
                    item6 = Space(20)
                    item7 = Space(20)
                    item8 = "LATITUD"
                    item9 = "LONGITUDE"
                    If temp1 = Nothing Then temp1 = 0
                    Select Case Condition
                        Case "  1"
                            pcod = default_Renamed
                            pcrp = convertFormat.Convert(jx4, "##0")
                            POpv1 = Format(opv1, "######0.")
                            POpv2 = Format(lun, "######0.")
                            POpv5 = Format(xtp1, "######0.")
                            item1 = "Heat Units"
                            item2 = "Curve Number"
                            item5 = "Plant Population"
                        Case "  2", " 10"
                            pcod = default_Renamed
                            If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
                                 Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                                POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 32, 12))), 0), "######0.")
                            Else
                                POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 33, 6))), 0), "######0.")
                            End If
                            item1 = "Irrigation"
                        Case "  3", " 11"  'fertilizer
                            If (lyr = "   0") Or (lyr = "    ") Then lyr = "  50"
                            pcod = convertFormat.Convert(jx4, "##0")
                            pcrp = convertFormat.Convert(temp1, "##0")
                            pmat = convertFormat.Convert(lyr, "###0")
                            If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "3.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
                                 Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                                POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 32, 12))), 0), "######0.")
                            Else
                                POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 33, 6))), 0), "######0.")
                            End If
                            POpv2 = convertFormat.Convert(10, "######0.")
                            Pmat1 = Val(Trim(pmat))
                            Fertilizer = getDBDataTableNoCon("SELECT Name FROM Fertilizer WHERE code = " & Pmat1, myConnection)
                            If Fertilizer.Rows.Count > 0 Then
                                item1 = Fertilizer.Rows(0).Item("Name")
                            Else
                                item1 = "Fertilizer"
                            End If

                            Fertilizer.Dispose()
                            Fertilizer = Nothing
                            item2 = "Depth"
                        Case "  4"  'pesticide
                            pcod = convertFormat.Convert(jx4, "##0")
                            pmat = convertFormat.Convert(108, "###0")
                            item1 = "Control Factor"
                            item2 = "Pesticide"
                        Case "  5"
                            pcod = default_Renamed
                            pcrp = convertFormat.Convert(temp1, "##0")
                            POpv7 = convertFormat.Convert(opv7, "####0.00")
                            item1 = "Time of Operation"
                        Case "  6"  'tillage
                            pcod = convertFormat.Convert(jx4, "##0")
                            pcrp = convertFormat.Convert(temp1, "##0")
                            item1 = "Tillage"
                        Case "  7"   'harvest
                            pcod = default_Renamed
                            pcrp = convertFormat.Convert(temp1, "##0")
                            POpv2 = convertFormat.Convert(System.Math.Round(lun, 0), "######0.")
                            POpv7 = convertFormat.Convert(opv7, "####0.00")
                            item2 = "Curve Number"
                        Case "  8"
                            pcod = default_Renamed
                            pcrp = convertFormat.Convert(temp1, "##0")
                            POpv2 = convertFormat.Convert(System.Math.Round(lun, 0), "######0.")
                            POpv7 = convertFormat.Convert(opv7, "####0.00")
                            item1 = "Curve Number"
                            item2 = "Time of Operation"
                        Case Else
                            If (lyr = "   0") Or (lyr = "    ") Then lyr = "  50"
                            pcod = convertFormat.Convert(jx4, "##0")
                            pcrp = convertFormat.Convert(temp1, "##0")
                            pmat = convertFormat.Convert(lyr, "###0")
                            If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
                                 Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                                POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 32, 12))), 0), "######0.")
                            Else
                                POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 33, 6))), 0), "######0.")
                            End If
                    End Select

                    'adoCrop = getDBDataTable("SELECT description FROM S_A_CROP WHERE APEX_code = " & Val(pcrp))
                    'If adoCrop.Rows.Count > 0 Then cropName = adoCrop.Rows(0).Item("Description")
                    'adotill = getDBDataTable("SELECT description FROM S_A_TILL WHERE APEX_code = " & Val(pcod))
                    'If adotill.Rows.Count > 0 Then tillName = adotill.Rows(0).Item("Description")
                    name = "fem"
                    FEMScenario = "Baseline"
                    If Initial.Scenario <> "Baseline" Then FEMScenario = "Scenario"
                    ReDim Preserve sSql(i)
                    If mvarnumber > 0 Then lat1 = Initial.lat1(mvarnumber - 1) : lon1 = Initial.lon1(mvarnumber - 1) Else lat1 = "" : lon1 = ""

                    sSql(i) = "INSERT INTO FEM (Composite,[Applies To],[year],[month],[day],[APEX Operation Code],operation,[Apex Crop Code]," &
                        "Crop,[Year in rotation],[Rotation Length],frequency,item1,value1,item2,value2,item3,value3,item4,value4,item5,value5" &
                        ",item6,value6,item7,value7,item8,value8,item9,value9) VALUES('" &
                        Initial.Scenario & " " & mvarfilemgt & "_" & SubbasinNum & "','" &
                        FEMScenario & "'," & year_Renamed & "," & month1 & "," & day1 & ",'" &
                        pcod & "','" & getData("SELECT description FROM S_A_TILL WHERE APEX_code = " & Val(pcod), myConnection) & "','" & Val(Temp2) & "','" & getData("SELECT description FROM S_A_CROP WHERE APEX_code = " & Val(pcrp), myConnection) & "'," &
                        year_Renamed & "," & rotation & ",1,'" &
                        item1 & "','" & Val(POpv1) & "','" &
                        item2 & "','" & Val(POpv2) & "','" &
                        item3 & "','" & Val(POpv3) & "','" &
                        item4 & "','" & Val(POpv4) & "','" &
                        item5 & "','" & Val(POpv5) & "','" &
                        item6 & "','" & Val(POpv6) & "','" &
                        item7 & "','" & Val(POpv7) & "','" &
                        item8 & "','" & lat1 & "','" &
                        item9 & "','" & lon1 & "')"
                    i += 1
                    'getLocalDataSet(sSql, Initial.Output_files)
                    name = "none"
                End If
            Loop

            UpdateStringArray(sSql, Initial.Output_files)
            'UpdateStringArray1(sSql, Initial.Output_files)
            a.Close()
            'FEM.Dispose()
            'FEM = Nothing
            adoCrop.Dispose()
            adoCrop = Nothing
            adoParm.Dispose()
            adoParm = Nothing

        Catch ex As Exception
            MsgBox(ex.Message & " SWATSubbasins " & name)
        End Try
    End Sub

    Public Sub APEXSubareas(ByRef SubbasinNum As String, myConnection As OleDb.OleDbConnection)
        Dim Rotation As Integer
        Dim tillName As String = String.Empty
        Dim cropName As String = String.Empty
        Dim Pmat1 As Object
        Dim pmat As Object
        Dim lun As Single = 0
        Dim pcrp As Object
        Dim item9, item8, item7, item6, item5, item4, item3, item2, item1 As String
        Dim POpv7, POpv6, POpv5, POpv4, POpv3, POpv2, POpv1, Pblank As String
        Dim opv1 As Object
        Dim xtp1 As Object
        Dim jx4 As String
        Dim lyr As Object
        Dim Condition As String
        Dim day1 As Object
        Dim month1 As Object
        Dim pcod As Object
        Dim Fert As Object
        Dim tempopc As Object
        Dim i As UShort
        Dim a As StreamReader = Nothing
        Dim temp1 As Object
        Dim Temp2 As String
        Dim opv7 As Object
        Dim Line1 As Object
        'Dim fs As Object
        Dim TakeField As Object
        Dim currdir As String
        Dim Crop1310, FEM, Fertilizer, adoCrop, adoTill As DataTable
        Dim col2, col3, col4 As Integer
        Dim limit As Short
        Dim temp As String
        Dim year_Renamed As Short
        Dim k As Short
        Dim sqlcmd As String()

        currdir = CurDir()

        On Error GoTo goError

        TakeField = New Convertion
        'fs = CreateObject("Scripting.FileSystemObject")
        convertFormat = New Convertion
        Fertilizer = New DataTable
        Crop1310 = New DataTable

        Select Case Initial.Version
            Case "1.0.0"
                limit = 2
                col2 = 30
                col3 = 21
                col4 = 7
            Case "1.1.0"
                limit = 2
                col2 = 30
                col3 = 33
                col4 = 5
            Case "1.2.0"
                limit = 2
                col2 = 30
                col3 = 35
                col4 = 43
            Case "1.3.0"
                limit = 2
                col2 = 30
                col3 = 35
                col4 = 43
            Case "2.0.0"
                limit = 2
                col2 = 29
                col3 = 21
                col4 = 7
            Case "2.1.0"
                limit = 2
                col2 = 20
                col3 = 33
                col4 = 5
            Case "2.3.0"
                limit = 2
                col2 = 20
                col3 = 33
                col4 = 5
            Case "3.0.0"
                limit = 2
                col2 = 29
                col3 = 21
                col4 = 7
            Case "3.1.0"
                limit = 2
                col2 = 20
                col3 = 33
                col4 = 4
            Case "4.0.0"
                limit = 2
                col2 = 27
                col3 = 35
                col4 = 43
            Case "4.1.0"
                limit = 2
                col2 = 27
                col3 = 35
                col4 = 43
            Case "4.2.0"
                limit = 2
                col2 = 27
                col3 = 35
                col4 = 43
            Case "4.3.0"
                limit = 2
                col2 = 27
                col3 = 35
                col4 = 43
        End Select

        k = 0
        Dim RecOper(k) As Object
        Line1 = 3
        flag = False
        opv7 = 0.0#
        Temp2 = ""
        temp1 = 0
        a = New StreamReader(Initial.Output_files & "\" & Trim(mvarfilemgt))
        For i = 1 To limit
            a.ReadLine()
        Next

        year_Renamed = 1
        tempopc = ""

        '/* delete records added before regarding this project and operation file */
        Fert = 0
        i = 0
        Do While a.EndOfStream <> True
            temp = a.ReadLine
            If Trim(temp) = "" Then Exit Sub
            pcod = Mid(temp, Initial.col2, 3)
            year_Renamed = Val(Left(temp, Initial.Year_Col))
            month1 = Val(Mid(temp, Initial.Year_Col + 1, Initial.Year_Col))
            day1 = Val(Mid(temp, Initial.Year_Col + Initial.Year_Col + 1, Initial.Year_Col))
            Temp2 = Val(Mid(temp, Initial.col7, 4)) 'select crop code

            'Select Case tempopc
            Select Case pcod
                Case "580"
                    Fert = Val(Mid(temp, col2, 5))
                    Condition = "  3"
                    lyr = Mid(temp, 16, 4)
                    lyr = 0
                Case "136"
                    Condition = "  1"
                    lyr = 0
                Case "621"
                    Condition = "  4"
                    lyr = Mid(temp, 21, 8)
                Case "570"
                    Condition = "  5"
                    lyr = 0
                Case "623"
                    Condition = "  7"
                    lyr = 0
                Case "451"
                    Condition = "  8"
                    lyr = 0
                Case "501"
                    Condition = "  2"
                    lyr = 0
                Case "426"
                    Condition = "  9"
                    lyr = 0
                Case Else
                    Condition = "  6"
                    lyr = 0
            End Select

            jx4 = pcod
            Crop1310 = getDBDataTableNoCon("SELECT PPLP2 From crop1310 WHERE [Numb]=" & Temp2, myConnection)
            xtp1 = Int(Crop1310.Rows(0).Item("PPLP2"))
            opv1 = convertFormat.Convert(System.Math.Round(Val(Mid(temp, col2, 8))), "#####0.0")
            If (opv1 = 0) Then
                opv1 = 2000
            End If

            POpv1 = Space(8)
            POpv2 = Space(8)
            POpv3 = Space(8)
            POpv4 = Space(8)
            POpv5 = Space(8)
            POpv6 = Space(8)
            POpv7 = Space(8)
            item1 = Space(20)
            item2 = Space(20)
            item3 = Space(20)
            item4 = Space(20)
            item5 = Space(20)
            item6 = Space(20)
            item7 = Space(20)
            item8 = "LATITUD"
            item9 = "LONGITUDE"
            lyr = lyr.ToString()
            Select Case Condition
                Case "  1"
                    pcrp = convertFormat.Convert(jx4, "##0")
                    POpv1 = convertFormat.Convert(opv1, "######0.")
                    POpv2 = convertFormat.Convert(System.Math.Round(Val(Mid(temp, col3, 8))), 0)
                    POpv5 = convertFormat.Convert(System.Math.Round(xtp1, 0), "######0.")
                    item1 = "Heat Units"
                    item2 = "Curve Number"
                    item5 = "Plant Population"
                Case "  2", " 10"
                    If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
                         Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                        POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, col3, 12))), 0), "######0.")
                    Else
                        POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 33, 6))), 0), "######0.")
                    End If
                    item1 = "Irrigation"
                Case "  3", " 11"
                    If (lyr = "   0") Or (lyr = "    ") Then lyr = "  50"
                    pcod = convertFormat.Convert(jx4, "##0")
                    pcrp = convertFormat.Convert(Temp2, "##0")
                    pmat = convertFormat.Convert(lyr, "###0")
                    If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "3.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
                         Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                        If Mid(temp, col2, 8) = "        " Then
                            POpv1 = Mid(temp, col2, 8)
                        Else
                            POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, col2, 8))), 0), "######0.")
                        End If

                    Else
                        POpv1 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, 21, 8))), 0), "######0.")
                    End If

                    If Mid(temp, col3, 8) = "        " Then
                        POpv2 = Mid(temp, col3, 8)
                    Else
                        POpv2 = convertFormat.Convert(System.Math.Round(CDbl(Val(Mid(temp, col3, 8))), 0), "######0.")
                    End If
                    Pmat1 = Val(Trim(pmat))
                    Fertilizer = getDBDataTableNoCon("SELECT Name FROM Fertilizer WHERE code = " & Pmat1, myConnection)
                    If Fertilizer.Rows.Count = 0 Then
                        item1 = "Fertilizer"
                    Else
                        item1 = Fertilizer.Rows(0).Item("Name")
                    End If

                    Fertilizer.Dispose()
                    Fertilizer = Nothing
                    item2 = "Depth"
                Case "  4"
                    pcod = convertFormat.Convert(jx4, "##0")
                    pmat = convertFormat.Convert(108, "###0")
                    item1 = "Control Factor"
                    item2 = "Pesticide"
                Case "  5"
                    pcrp = convertFormat.Convert(temp1, "##0")
                    POpv7 = convertFormat.Convert(opv7, "####0.00")
                    item1 = "Time of Operation"
                Case "  6"
                    pcod = convertFormat.Convert(jx4, "##0")
                    'pcrp = convertFormat.Convert(temp1, "##0")
                    pcrp = Format(temp1, "##0")
                    item1 = "Tillage"
                Case "  7"
                    pcrp = convertFormat.Convert(temp1, "##0")
                    POpv2 = convertFormat.Convert(System.Math.Round(lun, 0), "######0.")
                    POpv7 = convertFormat.Convert(opv7, "####0.00")
                    item2 = "Curve Number"
                Case "  8"  'Kill
                    pcrp = convertFormat.Convert(temp1, "##0")
                    POpv2 = convertFormat.Convert(System.Math.Round(lun, 0), "######0.")
                    POpv7 = convertFormat.Convert(opv7, "####0.00")
                    item1 = "Curve Number"
                    item2 = "Time of Operation"
                Case Else
                    If (lyr = "   0") Or (lyr = "    ") Then lyr = "  50"
                    pcod = convertFormat.Convert(jx4, "##0")
                    pcrp = convertFormat.Convert(temp1, "##0")
                    pmat = convertFormat.Convert(lyr, "###0")
                    If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
                         Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                        If Mid(temp, 32, 12) = "            " Then
                            POpv1 = Mid(temp, 32, 12)
                        Else
                            POpv1 = convertFormat.Convert(Math.Round(CDbl(Val(Mid(temp, 32, 12))), 0), "######0.")
                        End If

                    Else
                        If Mid(temp, 33, 6) = "      " Then
                            POpv1 = Mid(temp, 33, 6)
                        Else
                            POpv1 = convertFormat.Convert(Math.Round(CDbl(Val(Mid(temp, 33, 6))), 0), "######0.")
                        End If
                    End If
            End Select

            adoCrop = getDBDataTableNoCon("SELECT description FROM S_A_CROP WHERE APEX_code = " & Val(Temp2), myConnection)
            If adoCrop.Rows.Count > 0 Then cropName = adoCrop.Rows(0).Item("Description")
            adoCrop.Dispose()
            adoCrop = Nothing
            adoTill = getDBDataTableNoCon("SELECT description FROM S_A_TILL WHERE APEX_code = " & Val(pcod), myConnection)
            If adoTill.Rows.Count > 0 Then tillName = adoTill.Rows(0).Item("Description")
            adoTill.Dispose()
            adoTill = Nothing

            'With FEM
            Dim composite As String
            Dim applies_to As String
            Dim Op_code, operation, crop_code, crop, year_rotation, rotion_length, freq, value1, value2, value3, value4, value5, value6, value7 As String
            Dim value8, value9 As String
            Dim year, month, day As Integer

            composite = Initial.Scenario & " " & mvarfilemgt & "_" & SubbasinNum
            applies_to = "Baseline"
            If Initial.Scenario <> "Baseline" Then applies_to = "scenario"
            year = year_Renamed
            month = month1
            day = day1
            Op_code = pcod
            operation = tillName
            crop_code = Val(Temp2)
            crop = cropName
            year_rotation = year_Renamed
            rotion_length = Rotation
            freq = 1
            '.Fields("item1").Value = item1
            value1 = Val(POpv1)
            '.Fields("item2").Value = item2
            value2 = Val(POpv2)
            '.Fields("item3").Value = item3
            value3 = Val(POpv3)
            '.Fields("item4").Value = item4
            value4 = Val(POpv4)
            '.Fields("item5").Value = item5
            value5 = Val(POpv5)
            '.Fields("item6").Value = item6
            value6 = Val(POpv7)
            '.Fields("item7").Value = item7
            value7 = Val(POpv7)
            '.Fields("item8").Value = item8
            If mvarnumber > 0 Then value8 = Initial.lat1(mvarnumber - 1) Else value8 = 0
            '.Fields("item9").Value = item9
            If mvarnumber > 0 Then value9 = Initial.lon1(mvarnumber - 1) Else value9 = 0
            'item2 = 0
            'item3 = 0
            'item4 = 0
            'item5 = 0
            'item6 = 0
            'item7 = 0
            ReDim Preserve sqlcmd(i)

            sqlcmd(i) = "INSERT INTO FEM " &
                "(Composite,[Applies To],[Year],[Month],[Day],[APEX Operation Code],Operation,[APEX Crop Code]," &
                "Crop,[Year in rotation], [Rotation Length],Frequency, Item1, Value1, Item2, Value2, Item3, " &
                "Value3, Item4, Value4, Item5, Value5, Item6, Value6, Item7, Value7, Item8, Value8, Item9, Value9) " &
                "VALUES ('" & composite & "', '" & applies_to & "', " & year & ", " & month & "," & day & ",'" & Op_code & "', '" & operation & "'," & crop_code & ", '" & crop & "'," & year_rotation & ", " & rotion_length & ", '" & freq & "', '" & item1 & "', '" & value1 & "', '" & item2 & "', '" & value2 & "', '" & item3 & "', '" & value3 & "', '" & item4 & "', '" & value4 & "', '" & item5 & "', '" & value5 & "', '" & item6 & "', '" & value6 & "', '" & item7 & "', '" & value7 & "', '" & item8 & "', '" & value8 & "', '" & item9 & "', '" & value9 & "')"

            temp = a.ReadLine
            i += 1
        Loop
        UpdateStringArray(sqlcmd, Initial.Output_files)

        a.Close()

        Exit Sub
goError:
        MsgBox(Err.Description & "Fucntion-->APEXSubareas, File-->" & mvarfilemgt)

    End Sub

    'Function Weather_Code(ByRef fileSub As String) As Object
    '    Dim rec As DataTable

    '    rec = New DataTable

    '    rec = getDBDataTable("SELECT pcpNumber, File_Number FROM Sub_Included WHERE Subbasin=" & "'" & fileSub & "'")
    '    With rec
    '        Weather_Code = 1
    '        If .Rows.Count > 0 Then
    '            Weather_Code = Initial.File_Number
    '                getDBDataSet("UPDATE Subbasins SET File_Number = " & Initial.File_Number & " WHERE Subbasin=" & "'" & fileSub & "'")
    '            End If
    '    End With
    '    rec.Dispose()
    '    rec = Nothing

    'End Function

    Public Sub herds(ByRef ownerID As Short)
        'print blanks at the end of the file
        If ownerID = 0 Then
            Dim swFile As StreamWriter

            swFile = New StreamWriter(File.Open(Initial.Output_files & "\" & Initial.herd, FileMode.Append))
            swFile.WriteLine(" ")
            swFile.Close()
            swFile.Dispose()
            Exit Sub
        End If

    End Sub

    Public Sub copyAllFiles(ByVal orgPath As String, ByVal targetPath As String, ByVal searchPattern As String, overwrite As Boolean)
        Dim files() As String
        Dim fileo, filet As String
        Dim lastIndexOf As Integer
        'searchPattern = "SWAT*.bat"
        Dim i As UShort
        files = Directory.GetFiles(orgPath, searchPattern)
        For i = 0 To files.Length - 1
            lastIndexOf = files(i).LastIndexOf("\")
            filet = targetPath & "\" & files(i).Substring(lastIndexOf + 1)
            On Error Resume Next
            File.Copy(files(i), filet, overwrite)
        Next
    End Sub


End Class