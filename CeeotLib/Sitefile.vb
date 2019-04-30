Option Strict Off
Option Explicit On

Imports System.IO

Module Sitefiles
    Dim Last As Object
    Dim linex(200) As Object
    Dim line1(200) As Object
    Dim filename, Title, others, hruname As Object
    Dim temp As String
    Dim m, soilnum, opt, l, z As Object
    Dim numSub As Short
    Dim Code, char1, luse1, reg As Object
    Dim luse(4) As Object
    Dim newName As String
    Dim Dataclass As General
    Dim flag As Boolean
    Dim TakeField As Convertion
    Dim convertFormat As Convertion
    Dim u As Object
    Dim totalArea As Double

    Public Function SiteFiles(ByRef index As String) As Object
        Dim pos As Integer
        Dim r, r1 As String
        Dim new_area1 As Object
        Dim opcname As String
        Dim numsol As Short = 0
        Dim k, l, m, flag1, o As Integer
        Dim hrudes As String
        Dim hrunum As Integer
        Dim a, c As StreamReader
        Dim Subbasin, temp1 As String
        Dim F As StreamWriter = Nothing
        Dim h, g, e, d, b As StreamWriter
        Dim fs As Object
        Dim fs1 As StreamWriter
        Dim sub_area As Single
        Dim new_area As Single
        Dim Exclude As DataTable
        Dim ADORecordset As DataSet
        Dim i, j As Integer
        Dim hrutot As Short
        Dim hrufiles() As String = Nothing
        Dim value As String
        Dim last_filename As String = String.Empty
        Dim Subfile() As String = Nothing
        Dim filename As String
        Dim X As Object

        On Error GoTo goError

        TakeField = New Convertion
        Exclude = New DataTable
        'cn = New ADODB.Connection
        fs = CreateObject("Scripting.FileSystemObject")

        If CDbl(index) = 3 Then
            modifyRecords("DELETE * FROM grazing")
            Initial.herds = 0
            Initial.owners = 0
        End If

        numSub = 0
        filename = Right(last_filename, 13)
        ADORecordset = getDBDataSet("SELECT * FROM Apexfiles WHERE Apexfile = 'Subasins' AND version =" & "'" & Initial.Version & "'" & " ORDER BY line, field")

        Dataclass = New General
        d = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.suba))
        e = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Soil))
        g = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Opcs))
        h = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Site))

        If CDbl(Create_Files_Form.Tag) = 2 Then F = New StreamWriter(File.OpenWrite(Initial.Output_files & Initial.RunFile))
        X = fs.OpenTextFile(Initial.Output_files & Initial.cntrl1)
        u = fs.createtextfile(Initial.Output_files & "\X" & Initial.suba)
        convertFormat = New Convertion

        Subbasin = "subbasin"
        For i = 1 To Initial.limit
            X.ReadLine()
        Next

        Initial.File_Number = 0
        Do While X.AtEndOfStream <> True
            Initial.File_Number = Initial.File_Number + 1
            reg = Mid(X.ReadLine, 1, 50)
            filename = Mid(reg, Initial.col3, 13)
            Wait_Form.Label1(4).Text = filename
            Wait_Form.Show()
            Wait_Form.Refresh()

            Code = Mid(reg, Initial.col4, 4)
            temp1 = X.ReadLine
            Subbasin = Mid(temp1, 1, 8)
            a = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & Trim(filename)))
            i = 0
            j = 0

            Do While a.EndOfStream <> True
                i = i + 1
                ReDim Preserve Subfile(i)
                Subfile(i) = a.ReadLine

                If (Mid(Subfile(i), 10, 4) = ".hru") Then
                    j = j + 1
                    ReDim Preserve hrufiles(j)
                    hrufiles(j) = Subfile(i)
                End If
            Loop

            soilnum = 0
            i = 0

            With ADORecordset.Tables(0)
                For i = 0 To .Rows.Count - 1
                    If (.Rows(i).Item("SwatFile") <> "") Then
                        TakeField.filename = Initial.Input_Files & "\" & Trim(filename)
                        TakeField.Leng = .Rows(i).Item("leng")
                        TakeField.LineNum = .Rows(i).Item("Lines")
                        TakeField.Inicia = .Rows(i).Item("Inicia")
                        If .Rows(i).Item("Condition") Is DBNull.Value Then
                            TakeField.Condition = ""
                        Else
                            TakeField.Condition = .Rows(i).Item("Condition")
                        End If
                        TakeField.Description = .Rows(i).Item("Description")
                        TakeField.col1 = .Rows(i).Item("Col1")
                        TakeField.col2 = .Rows(i).Item("Col2")
                        value = TakeField.value()
                    Else
                        value = .Rows(i).Item("Value")
                    End If

                    If .Rows(i).Item("Apex_Field") = "HRUTOT" Then
                        hrunum = Val(value)
                        hrudes = " " & .Rows(i).Item("Description")
                    End If
                Next
            End With

            l = 0
            m = 0
            k = 0

            For z = 1 To UBound(hrufiles)
                hruname = Trim(Mid(hrufiles(z), 14, 13))
                If (hruname <> "") Then
                    c = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & Trim(hruname)))
                    Title = c.ReadLine
                    i = InStr(1, Title, "Luse:")
                    temp = ""
                    For j = i + 5 To i + 8
                        temp = temp & Mid(Title, j, 1)
                    Next
                    temp = Trim(temp)

                    On Error Resume Next
                    Exclude = getDBDataTable("SELECT * FROM Exclude WHERE code = " & "'" & temp & "' AND Project= '" & Initial.Project & "'")

                    If Exclude.Rows.Count > 0 Then
                        flag1 = 1
                    Else
                        flag1 = 0
                        k = k + 1
                        m = m + 1
                        linex(k) = hrufiles(z)
                    End If
                    Exclude.Dispose()
                    Exclude = Nothing
                    If flag1 = 1 Then
                        l = l + 1
                        line1(l) = hrufiles(z)
                    End If

                    c.Close()
                End If
            Next

            flag = True
            If (m = 0) Then
                m = m + 1
                k = k + 1
                linex(k) = line1(1)
                flag = False
            End If

            hrunum = m
            a.Close()

            getDBDataSet("DELETE * FROM Subarea")

            For i = 1 To l 'modified
                numsol = numsol + 1
                Dataclass.filesol = Trim(Mid(line1(i), 27, 9)) & ".sol"
                Dataclass.fileSub = Trim(Mid(line1(i), 2, 8)) & ".sba"
                Dataclass.filesoi = Trim(Mid(line1(i), 28, 8)) & ".soi"
                Dataclass.filehru = Trim(Left(line1(i), 13))
                Dataclass.filemgt = Trim(Mid(line1(i), 14, 13)) '& ".mgt"

                If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Then
                    Dataclass.filerte = Trim(Mid(filename, 1, 9)) & ".rte"
                    Dataclass.filepnd = Trim(Mid(filename, 1, 9)) & ".pnd"
                Else
                    Dataclass.filerte = Trim(Mid(reg, 21, 14))
                    Dataclass.filepnd = Trim(Mid(reg, 21, 14))
                End If

                Dataclass.filesit = Trim(Mid(filename, 2, 8)) & ".sit"
                Dataclass.filechm = Trim(Mid(line1(i), 40, 9)) & ".chm"
                Dataclass.filewht = Trim(Mid(line1(i), 2, 8)) & ".wth"

                If Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "3.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Then
                    Dataclass.filewp1 = Trim(Mid(filename, 2, 8)) & ".wp1"
                    Dataclass.filewgn = Trim(Mid(filename, 1, 9)) & ".wgn"
                Else
                    Dataclass.filewp1 = Trim(Mid(temp1, 8, 8)) & ".wp1"
                    Dataclass.filewgn = Trim(Mid(filename, 1, 9)) & ".wgn"
                End If

                Dataclass.filename = filename
                Dataclass.number = numsol
                Dataclass.flag = flag

                Select Case Initial.Version
                    Case "2.0.0", "2.1.0", "2.3.0"
                        Espace = " "
                    Case "3.0.0", "3.1.0"
                        Espace = "  "
                    Case "4.0.0", "4.1.0", "4.2.0", "4.3.0", "1.1.0", "1.2.0", "1.3.0"
                        Espace = " "
                End Select

                Select Case Create_Files_Form.Tag
                    Case "2"  'Suarea files
                        d.WriteLine(convertFormat.Convert(numsol, "####0") & Initial.Espace & Trim(Mid(line1(i), 2, 8)) & ".sba")
                        Dataclass.suba((i))
                        If flag = False Then
                            Dataclass.flag = flag
                            Dataclass.Updatehru()
                            flag = True
                        End If
                    Case "3"   'Soil files
                        e.WriteLine(convertFormat.Convert(numsol, "####0") & Initial.Espace & Mid(line1(i), 28, 8) & ".soi" & "     " & Trim(Mid(line1(i), 29, 8)))
                        Dataclass.Soil()
                    Case "4"   'Operation files
                        Dataclass.fileSub = Trim(Mid(line1(i), 15, 8)) & ".opc"
                        opcname = Trim(Mid(line1(i), 15, 8)) & ".opc"
                        g.WriteLine(convertFormat.Convert(CInt(numsol), "####0") & Initial.Espace & opcname.PadLeft(12) & "     " & Trim(Mid(line1(i), 15, 8)))
                        Dataclass.Operations()
                    Case "5"   'Site files
                        Dataclass.fileSub = filename
                        Dataclass.Last = Left(last_filename, 4)
                        h.WriteLine(convertFormat.Convert(numsol, "####0") & Initial.Espace & Trim(Mid(filename, 2, 8)) & ".sit")
                        Dataclass.Site()
                        fs.CopyFile(Initial.Output_files & "\" & Initial.parm, Initial.Output_files & "\" & Trim(Mid(filename, 2, 8)) & ".prm")
                    Case "6"   'Weather files
                        Dataclass.Last = Mid(temp1, 49, 4)
                        Dataclass.Lasttmp = Mid(temp1, 53, 4)
                        Dataclass.Weather()
                    Case "7"   'wpm files
                        Dataclass.Last = Left(last_filename, 4)
                        Dataclass.wmpfiles()
                        Dataclass.number = numsol
                        Dataclass.wpm11310()
                End Select

                If Create_Files_Form.Tag = "2" Then
                    If Initial.subareafile = 1 Or Initial.Version = "3.0.0" Or Initial.Version = "3.1.0" Then
                        F.Write(Mid(line1(i), 2, 8))
                        F.Write(convertFormat.Convert(numsol, "###0"))
                        F.Write(convertFormat.Convert(1, "###0"))           'always 1. China.wnd. Before it was numsol
                        If Initial.Version = "3.0.0" Or Initial.Version = "3.1.0" Then
                            F.Write("    ") 'iwp5 Epic does not need
                        End If
                        F.Write("   1")
                        F.WriteLine(convertFormat.Convert(numsol, "###0") & convertFormat.Convert(numsol, "###0") & convertFormat.Convert("0", "###0"))
                    End If
                End If
            Next

            If Create_Files_Form.Tag = "2" Then
                If (CDbl(Create_Files_Form.Tag) = 2 And l > 0) Or (Initial.subareafile = 2 And l > 0) Then
                    numSub = numSub + 1
                    Call Subarea_Calculation(filename)

                    b = New StreamWriter(File.Create(Initial.New_Swat & "\" & filename))
                    o = 1
                    Do While Mid(Subfile(o), 10, 4) <> ".hru"
                        If o = 2 Then
                            sub_area = Val(Left(Subfile(o), 16))
                            new_area = sub_area * 100 - totalArea
                            If new_area <= 0.001 Then
                                new_area1 = 0.000001
                            Else
                                new_area1 = new_area / 100
                            End If
                            new_area1 = convertFormat.Convert(new_area1, "#####0.000000")
                            r = InStr(1, Subfile(o), "|")
                            r1 = Mid(Subfile(o), 1, r - 1)
                            r1 = Trim(r1)
                            Subfile(o) = Replace(Subfile(o), r1, new_area1)
                        End If
                        hrutot = InStr(1, Subfile(o), "HRUTOT")
                        If hrutot > 0 Then
                            Subfile(o) = (convertFormat.Convert(hrunum, "##############0 ") & "| " & Mid(Subfile(o), hrutot, 73))
                        End If
                        b.WriteLine(Subfile(o))
                        o = o + 1
                    Loop

                    For o = 1 To k
                        b.WriteLine(linex(o))
                        pos = InStr(1, linex(o), ".hru")
                        If Initial.Version = "4.0.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
                            Or Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then Call newHRUCalculation(sub_area, new_area, Left(linex(o), pos + 3))
                    Next
                    b.Close()

                End If

                If Initial.subareafile = 2 And l > 0 Then
                    If Initial.Version <> "3.0.0" And Initial.Version <> "3.1.0" Then
                        F.Write((Mid(filename, 2, 7)) & "0")
                        If Initial.Version <> "1.1.0" Or Initial.Version <> "1.2.0" Or Initial.Version <> "1.3.0" Then F.Write(" ")
                        F.Write(convertFormat.Convert(numsol, "##0"))
                        If Initial.Version <> "1.1.0" Or Initial.Version <> "1.2.0" Or Initial.Version <> "1.3.0" Then F.Write(" ")
                        F.Write(convertFormat.Convert(1, "###0"))       'always 1. China.wnd. Before it was numsol
                        If Initial.Version <> "1.1.0" Or Initial.Version <> "1.2.0" Or Initial.Version <> "1.3.0" Then F.Write(" ")
                        F.Write("   1")
                        If Initial.Version <> "1.1.0" Or Initial.Version <> "1.2.0" Or Initial.Version <> "1.3.0" Then F.Write(" ")
                        F.Write(convertFormat.Convert(numSub, "###0") & convertFormat.Convert("0", "###0"))
                        If Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                            F.WriteLine("   0")
                        Else
                            F.WriteLine("")
                        End If
                    End If
                End If
            End If

        Loop

        u.Close()
        u.Dispose()
        u = Nothing
        d.Close()
        d.Dispose()
        d = Nothing
        e.Close()
        e.Dispose()
        e = Nothing
        g.Close()
        g.Dispose()
        g = Nothing
        h.Close()
        h.Dispose()
        h = Nothing
        X.Close()
        X.Dispose()
        X = Nothing
        F.Close()
        F.Dispose()
        F = Nothing
        If Initial.subareafile = 2 And CDbl(Create_Files_Form.Tag) = 2 Then
            fs.CopyFile(Initial.Output_files & "\X" & Initial.suba, Initial.Output_files & "\" & Initial.suba)
        End If

        Call Dataclass.herds(0)
        Call Dataclass.herds(0)
        Call Dataclass.herds(0)
        Call Dataclass.herds(0)

        Exit Function
goError:
        MsgBox(Err.Description & "SiteFiles")

    End Function

    Public Sub Subarea_Calculation(ByRef filename As String)
        Dim line5 As String
        Dim line4 As String
        Dim rchl As String = String.Empty
        Dim CHL As String = String.Empty
        Dim srFile As StreamReader
        Dim swFile As StreamWriter = Nothing
        Dim subarea, SubIncluded As DataTable
        Dim hrus As Short
        Dim i, k, j, n As Integer
        Dim recNum As Integer
        Dim numberFormat As String = ""

        Try
            subarea = New DataTable
            SubIncluded = New DataTable

            If Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "4.1.0" Or Initial.Version = "3.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Then
                subarea = getDBDataTable("SELECT Swat, Sum(Area) AS [TotalArea] FROM Subarea WHERE Swat = " & "'" & filename & "' GROUP BY Swat")
            Else
                subarea = getDBDataTable("SELECT Swat, Sum(Area) AS [TotalArea] FROM Subarea GROUP BY Swat")
            End If

            If subarea.Rows.Count > 0 Then
                totalArea = subarea.Rows(0).Item("TotalArea")
                subarea = getDBDataTable("Select * FROM subarea ORDER BY area")
            End If

            If Initial.subareafile = 2 Then
                swFile = New StreamWriter(File.Create(Initial.Output_files & "\SubATemp.sub"))
            End If

            hrus = 1 ' control number of HRUs in the same subarea file to define sign of area and chl and rchl values
            recNum = subarea.Rows.Count - 1

            k = 0
            j = 0
            With subarea
                For n = 0 To subarea.Rows.Count - 1
                    srFile = New StreamReader(File.OpenRead(Initial.Output_files & "\" & .Rows(n).Item("APEX")))
                    If Initial.subareafile = 1 Then
                        swFile = New StreamWriter(File.Create(Initial.Output_files & "\SubATemp.sub"))
                    End If

                    For k = 1 To 3
                        swFile.WriteLine(srFile.ReadLine)
                    Next

                    Select Case Create_Files_Form.ReachLenght
                        Case 0, 1
                            CHL = Format(.Rows(n).Item("chl"), "0000.000")
                            rchl = Format(.Rows(n).Item("rchl"), "0000.000")
                        Case 2
                            CHL = Format(.Rows(n).Item("Area") / totalArea * .Rows(n).Item("chl"), "00000.00")
                            rchl = Format(.Rows(n).Item("rchl"), "0000.000")
                        Case 3
                            CHL = Format(0, "00000.00")
                            rchl = CHL
                    End Select

                    If j = 0 Then rchl = CHL

                    If j + 1 = recNum Then hrus = 0

                    j = j + 1
                    line4 = srFile.ReadLine
                    line5 = srFile.ReadLine
                    numberFormat = "###0.000"
                    If Val(Left(line4, 8)) > 9999 Then numberFormat = "###000.0"

                    Select Case hrus
                        Case 0
                            line4 = convertFormat.Convert(Val(Left(line4, 8)), numberFormat) & convertFormat.Convert(CHL, "###0.000") & Mid(line4, 17, 64)
                            line5 = convertFormat.Convert(rchl, "###0.000") & Mid(line5, 9, 72)
                        Case 1
                            line4 = Left(line4, 8) & convertFormat.Convert(CHL, "###0.000") & Mid(line4, 17, 64)
                            line5 = convertFormat.Convert(rchl, "###0.000") & Mid(line5, 9, 72)
                        Case 2 ' negative is used just to add one field to another, positive is used to route one field to another.
                            line4 = convertFormat.Convert(Val(Left(line4, 8)), numberFormat) & convertFormat.Convert(CHL, "###0.000") & Mid(line4, 17, 64)
                            line5 = convertFormat.Convert(rchl, "###0.000") & Mid(line5, 9, 72)
                        Case Else ' negative is used just to add one field to another, positive is used to route one field to another.
                            line4 = convertFormat.Convert(Val(Left(line4, 8)), numberFormat) & convertFormat.Convert(CHL, "###0.000") & Mid(line4, 17, 64)
                            line5 = convertFormat.Convert(rchl, "###0.000") & Mid(line5, 9, 72)
                    End Select

                    swFile.WriteLine(line4)
                    swFile.WriteLine(line5)

                    For i = 1 To Initial.sublines - 5
                        swFile.WriteLine(srFile.ReadLine)
                    Next

                    If Initial.subareafile = 1 Then
                        swFile.WriteLine()
                    End If

                    srFile.Close()
                    srFile.Dispose()
                    srFile = Nothing

                    If Initial.subareafile = 1 Then
                        swFile.Close()
                        File.Copy(Initial.Output_files & "\SubATemp.sub", Initial.Output_files & "\" & .Rows(n).Item("Apex"), True)
                    End If

                    If Initial.subareafile = 2 Then
                        File.Delete(Initial.Output_files & "\" & .Rows(n).Item("Apex"))
                        hrus = hrus + 1
                    End If

                    newName = Left(.Rows(n).Item("APEX"), 4) & "0000.sba"
                    getDBDataSet("DELETE * FROM subarea WHERE apex='" & .Rows(n).Item("Apex"))
                Next
            End With
            swFile.WriteLine(" ")

            If Initial.subareafile = 2 Then
                swFile.WriteLine(" ")
                swFile.Close()
                swFile.Dispose()
                swFile = Nothing

                File.Copy(Initial.Output_files & "\SubATemp.sub", Initial.Output_files & "\" & newName, True)
                newName = convertFormat.Convert(numSub, "####0") & Initial.Espace & newName
                u.WriteLine(newName)
                newName = ""
            End If
            subarea.Dispose()
            subarea = Nothing
            getDBDataSet("UPDATE Sub_Included SET Area= " & totalArea & " WHERE Subbasin= '" & filename & "'")

        Catch ex As Exception
            MsgBox(ex.Message, , "SiteFiles - Subrutine: Percentages")
        End Try

    End Sub

    Public Sub hrus()
        Dim hrunum As Object
        Dim flag1 As Object
        Dim c As StreamReader
        Dim j As Object
        Dim temp1 As String
        Dim a As StreamReader
        Dim Subbasin As Object
        Dim X As StreamReader
        Dim name1 As Object
        Dim area As Object
        Dim totalArea As Double
        Dim Exclude As DataTable
        Dim i As Integer
        Dim k As Integer
        Dim hrufiles() As String = Nothing
        Dim value As String
        Dim last_filename As String = String.Empty
        Dim Subfile() As String
        Dim filename As String

        On Error GoTo goError

        TakeField = New Convertion
        Exclude = New DataTable
        name1 = "HRUs"
        numSub = 0
        filename = Right(last_filename, 13)
        Dataclass = New General
        X = New StreamReader(File.OpenRead(Initial.Output_files & Initial.cntrl1))

        convertFormat = New Convertion
        Subbasin = "subbasin"

        For i = 1 To Initial.limit
            temp1 = X.ReadLine
        Next

        Do While X.EndOfStream <> True
            If Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Then
                If Subbasin <> "subbasin" Then Exit Do
            End If

            reg = Mid(X.ReadLine, 1, 30)
            filename = Mid(reg, Initial.col3, 13)
            Code = Mid(reg, Initial.col4, 4)
            temp1 = X.ReadLine
            Subbasin = Mid(temp1, 1, 8)
            a = New StreamReader(Initial.Input_Files & "\" & Trim(filename))
            i = 0
            j = 0

            Do While a.EndOfStream <> True
                i = i + 1
                ReDim Preserve Subfile(i)
                Subfile(i) = a.ReadLine

                If (Mid(Subfile(i), 10, 4) = ".hru") Then
                    j = j + 1
                    ReDim Preserve hrufiles(j)
                    hrufiles(j) = Subfile(i)
                End If
            Loop

            soilnum = 0
            i = 0
            l = 0
            m = 0
            k = 0

            For z = 1 To UBound(hrufiles)
                hruname = Mid(hrufiles(z), 14, 13)
                If (hruname <> "") Then
                    c = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & Trim(hruname)))
                    Title = c.ReadLine
                    i = 17
                    char1 = "X"

                    Do While (char1 <> ":")
                        char1 = Mid(Title, i, 1)
                        i = i + 1
                    Loop

                    temp = ""
                    For j = i To i + 4
                        temp = temp & Mid(Title, j, 1)
                    Next

                    On Error Resume Next
                    Exclude = getDBDataTable(("SELECT * FROM Exclude WHERE code = " & "'" & temp & "'"))

                    If Exclude.Rows.Count > 0 Then
                        flag1 = 1
                    Else
                        flag1 = 0
                        k = k + 1
                        m = m + 1
                        linex(k) = hrufiles(z)
                    End If

                    Exclude.Dispose()
                    Exclude = Nothing
                    If flag1 = 1 Then
                        l = l + 1
                        line1(l) = hrufiles(z)
                    End If

                    c.Close()
                End If
            Next

            flag = True
            If (m = 0) Then
                m = m + 1
                k = k + 1
                linex(k) = line1(1)
                flag = False
            End If

            hrunum = m
            a.Close()

        Loop

        Exit Sub

goError:
        MsgBox(Err.Description, , "HRUs - Subrutine: hrus")

    End Sub

    Public Sub hrus1()
        Dim NumberFormat As Object
        Dim roundformat As Object
        Dim lenFormat As Object
        Dim totalare As Object
        Dim current_line As Object
        Dim fs As Object
        Dim ADORecordset As DataTable
        Dim HRUset As DataTable
        Dim TakeField As Convertion
        Dim value As String = String.Empty
        Dim temp As String
        Dim totalArea, Fraction As Single
        Dim i, j As Integer

        On Error GoTo goError

        fs = CreateObject("Scripting.FileSystemObject")
        convertFormat = New Convertion
        TakeField = New Convertion
        HRUset = New DataTable

        ADORecordset = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'HRUs' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")
        HRUset = getDBDataTable(("SELECT * FROM HRUs WHERE Percentage = 0"))

        With ADORecordset
            current_line = .Rows(0).Item("Line")
            For i = 0 To .Rows.Count - 1
                If (.Rows(i).Item("SwatFile") <> "") Then
                    Select Case .Rows(i).Item("SwatFile")
                        Case "*.hru"
                            For j = 0 To HRUset.Rows.Count - 1
                                TakeField.filename = Initial.Input_Files & "\" & HRUset.Rows(j).Item("HRU")
                                TakeField.Leng = .Rows(i).Item("Leng")
                                TakeField.LineNum = .Rows(i).Item("Lines")
                                TakeField.Inicia = .Rows(i).Item("Inicia")
                                Fraction = TakeField.value()
                                modifyRecords("UPDATE HRUs SET area=" & totalArea * Fraction & ",Percentage=" & Fraction)
                            Next
                        Case "bsn"
                            TakeField.filename = Initial.Input_Files & Create_Files_Form.bsnFile
                            TakeField.Leng = .Rows(i).Item("Leng")
                            TakeField.LineNum = .Rows(i).Item("Lines")
                            TakeField.Inicia = .Rows(i).Item("Inicia")
                            totalArea = TakeField.value()
                    End Select
                End If

                temp = .Rows(i).Item("Line") & .Rows(i).Item("Field")
                Select Case temp
                    Case "00"
                        totalare = Val(value)
                    Case "11"
                        Fraction = Val(value)
                End Select

                If Not IsDBNull(.Rows(i).Item("Format").Value) AndAlso .Rows(i).Item("Format").Value <> "" Then
                    lenFormat = Len(.Rows(i).Item("Format").Value)
                    roundformat = Right(Trim(.Rows(i).Item("Format").Value), 1)
                    NumberFormat = Left(Trim(.Rows(i).Item("Format").Value), lenFormat - 2)
                    value = convertFormat.Convert(System.Math.Round(Val(value), roundformat), NumberFormat)
                End If
                '.MoveNext()
            Next

        End With
        Exit Sub

goError:
        MsgBox(Err.Description, , "HRUs - Subrutine: hrus1")

    End Sub

    Public Sub FEMFiles(ByRef index As String)
        Dim flag1 As Object
        Dim Temp_Soil As Object
        Dim Pos_Soil As Object
        Dim c As Object
        Dim j As Object
        Dim Subbasin As Object
        Dim X As Object
        Dim a As Object
        Dim temp1 As String = String.Empty
        Dim fs As Object
        Dim Exclude As DataTable
        Dim Included As DataTable
        Dim i, k As Short
        Dim hrufiles() As String = Nothing
        Dim value As String : Dim last_filename As String = String.Empty
        Dim Subfile() As String
        Dim filename As String
        Dim apex_swat As String

        On Error GoTo goError

        Exclude = New DataTable
        Included = New DataTable
        fs = CreateObject("Scripting.FileSystemObject")

        numSub = 0
        a = Val(Mid(temp1, 49, 4))
        modifyLocalRecords("DELETE * FROM APEXHRUs", Initial.Output_files)
        filename = Right(last_filename, 13)
        X = fs.OpenTextFile(Initial.Output_files & Initial.cntrl4)
        Subbasin = "subbasin"

        For i = 1 To Initial.limit
            X.ReadLine()
        Next

        Do While X.AtEndOfStream <> True
            reg = Mid(X.ReadLine, 1, 50)
            filename = Mid(reg, Initial.col3, 13)
            a = Initial.Version
            Code = Mid(reg, Initial.col4, 4)
            temp1 = X.ReadLine
            a = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(filename))
            i = 0
            j = 0

            Do While a.AtEndOfStream <> True
                i = i + 1
                ReDim Preserve Subfile(i)
                Subfile(i) = a.ReadLine

                If (Mid(Subfile(i), 10, 4) = ".hru") Then
                    j = j + 1
                    ReDim Preserve hrufiles(j)
                    hrufiles(j) = Subfile(i)
                End If
            Loop

            soilnum = 0
            i = 0
            l = 0
            m = 0
            k = 0

            For z = 1 To UBound(hrufiles)
                hruname = Mid(hrufiles(z), 27, 13)
                If (hruname <> "") Then
                    c = fs.OpenTextFile(Initial.Input_Files & "\" & Trim(hruname))
                    Title = c.ReadLine
                    i = InStr(1, Title, "Luse:")
                    temp = ""

                    For j = i + 5 To i + 9
                        temp1 = Mid(Title, j, 1)
                        If temp1 = " " Then Exit For
                        temp = temp & temp1
                    Next

                    temp = Trim(temp)
                    Pos_Soil = InStr(Title, "Soil:")
                    Temp_Soil = Trim(Mid(Title, Pos_Soil + 5, 20))
                    Pos_Soil = InStr(Temp_Soil, " ")
                    Temp_Soil = Left(Temp_Soil, Pos_Soil - 1)
                    Exclude = getDBDataTable("SELECT * FROM Exclude WHERE code = " & "'" & Trim(temp) & "' AND Project= '" & Initial.Project & "'")

                    If Exclude.Rows.Count > 0 Then
                        flag1 = 1
                    Else
                        flag1 = 0
                        k = k + 1
                        m = m + 1
                        linex(k) = hrufiles(z)
                        modifyLocalRecords("INSERT INTO APEXHRUs (Project,Subarea,HRU,Apex_Swat,pcpGage,Managment,Soil,Chemical,Groundwater,SoilName,LandUse) " & _
                                         "VALUES('" & Initial.Project & "','" & filename & "','" & Left(linex(k), 13) & "','SWAT'," & _
                                         Val(Mid(temp1, 49, 4)) & ",'" & Trim(Mid(linex(k), 14, 13)) & "','" & _
                                         Trim(Mid(linex(k), 27, 13)) & "','" & Trim(Mid(linex(k), 40, 13)) & "','" & _
                                         Trim(Mid(linex(k), 53, 13)) & "','" & Temp_Soil & "','" & temp & "')", Initial.Output_files)
                    End If

                    If flag1 = 1 Then
                        l = l + 1
                        line1(l) = hrufiles(z)
                        Included = getDBDataTable(("SELECT * FROM Sub_Included WHERE folder = " & "'" & Initial.Dir1 & "'" & " AND Project= " & "'" & Initial.Project & "'" & " AND Subbasin = " & "'" & Trim(filename) & "'"))
                        apex_swat = "APEX"
                        If Included.Rows.Count = 0 Then apex_swat = "SWAT"
                        Included.Dispose()
                        Included = Nothing
                        modifyLocalRecords("INSERT INTO APEXHRUs (Project,Subarea,HRU,Apex_Swat,pcpGage,Managment,Soil,Chemical,Groundwater,SoilName,LandUse) " & _
                                         "VALUES('" & Initial.Project & "','" & filename & "','" & Left(line1(l), 13) & "','" & apex_swat & "'," & _
                                         Val(Mid(temp1, 49, 4)) & ",'" & Trim(Mid(line1(l), 14, 13)) & "','" & _
                                         Trim(Mid(line1(l), 27, 13)) & "','" & Trim(Mid(line1(l), 40, 13)) & "','" & _
                                         Trim(Mid(line1(l), 53, 13)) & "','" & Temp_Soil & "','" & temp & "')", Initial.Output_files)
                    End If
                    Exclude.Dispose()

                    c.Close()
                End If
            Next
        Loop

        Exit Sub
goError:
        MsgBox(Err.Description & "SiteFiles")

    End Sub

    ' Private Sub Create_mgt_Database()
    '  Dim temp_rec As Object
    '  Dim mgt1 As DataTable
    '  'Dim mgt2 As DataTable
    '  Dim APEXFiles As DataTable
    '  'Dim cn As ADODB.Connection
    '  Dim k As Short
    '  Dim fs, a As Object
    '  Dim value As String

    '  On Error GoTo goError

    '  mgt1 = New DataTable
    '  'mgt2 = New DataTable
    '  APEXFiles = DataTable
    '  'cn = New ADODB.Connection
    '  fs = CreateObject("Scripting.FileSystemObject")
    '  TakeField = New Convertion

    '  'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & Initial.Output_files & "\local.mdb"
    '  'cn.Open()
    '  APEXFiles = getDBDataTable(("SELECT * FROM Apexfiles WHERE Apexfile = 'mgt1' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field"))
    '  mgt1 = getLocalDataTable("SELECT * FROM mgt1", Initial.Output_files)

    '  With APEXFiles
    '   For k = 1 To m
    '    TakeField.filename = linex(m)
    '    TakeField.LineNum = .Rows(0).Item("Line")
    '    TakeField.Inicia = .Rows(0).Item("Field")
    '    value = TakeField.value
    '   Next
    '  End With

    'mgt2000:
    '  temp_rec = a.ReadLine
    '  Do
    '   With mgt1
    '    .AddNew()
    '   End With
    '  Loop
    '  Return

    'mgt2005:
    '  Return
    '  Exit Sub
    'goError:
    '  MsgBox(Err.Description & "Create_mgt_Database")

    ' End Sub

    Sub newHRUCalculation(ByVal sub_area As Single, ByVal new_area As Single, ByVal hruFile As String)
        Dim r1 As Object
        Dim r As Object
        Dim fs, a As Object
        Dim temp() As String
        Dim i As Short
        Dim fractionHRU As Object
        Dim areaHRU As Single

        On Error GoTo goError

        fs = CreateObject("Scripting.FileSystemObject")
        TakeField = New Convertion

        a = fs.OpenTextFile(Initial.Input_Files & "\" & hruFile)
        ReDim temp(2)
        temp(0) = a.ReadLine
        temp(1) = a.ReadLine
        i = 2

        Do While a.AtEndOfStream <> True
            temp(i) = a.ReadLine
            i = i + 1
            ReDim Preserve temp(i)
        Loop

        a.Close()
        i = InStr(1, temp(1), "|")
        fractionHRU = Left(temp(1), i - 1)
        areaHRU = fractionHRU * sub_area

        If new_area <= 0.001 Then
            fractionHRU = TakeField.Convert(1, "0.0000000")
        Else
            fractionHRU = TakeField.Convert(areaHRU / new_area * 100, "0.0000000")
        End If

        r = InStr(1, temp(1), "|")
        r1 = Mid(temp(1), 1, r - 1)
        r1 = Trim(r1)
        temp(1) = Replace(temp(1), r1, fractionHRU)
        a = fs.createtextfile(Initial.New_Swat & "\" & hruFile)

        For i = 0 To UBound(temp)
            a.WriteLine(temp(i))
        Next

        a.Close()
        Exit Sub

goError:
        MsgBox(Err.Description & "Create_mgt_Database")

    End Sub
End Module