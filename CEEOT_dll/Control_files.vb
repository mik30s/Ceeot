Option Strict Off
Option Explicit On

Imports System.IO

Public Module Control

    Public Sub Apexcont(ngnCreation As Short)
        Dim ngnFlag As Boolean
        Dim NumberFormat As String = String.Empty
        Dim roundformat, lenFormat, current_line As Integer
        Dim ADORecordset As New DataSet
        Dim TakeField As Convertion
        Dim value As String
        Dim swFile As StreamWriter = Nothing
        Dim swFile1 As StreamWriter = Nothing
        Dim i As Integer

        On Error GoTo goError

        TakeField = New Convertion

        ADORecordset = AccessDB.getDBDataSet("SELECT * FROM Apexfiles WHERE Version = " & "'" & Initial.Version & "'" & " AND Apexfile = 'Apexcont.dat' ORDER BY line, field")

        If ngnCreation = 0 Then
            swFile = New StreamWriter(File.Create(Initial.Output_files & "\" & Initial.herd))
            swFile.Close()
            swFile.Dispose()
            swFile1 = New StreamWriter(File.Create(Initial.Output_files & "\" & Initial.cont))
        End If

        Initial.ngn = 0

        With ADORecordset.Tables(0)
            current_line = .Rows(i).Item("Line")

            For i = 0 To .Rows.Count - 1
                If Not IsDBNull(.Rows(i).Item("SwatFile")) AndAlso .Rows(i).Item("SwatFile") <> "" Then
                    TakeField.filename = Initial.Input_Files & "\" & Trim(.Rows(i).Item("SwatFile"))
                    Select Case LCase(Trim(.Rows(i).Item("SwatFile")))
                        Case "basins.cod"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(Initial.CodFile)
                        Case "basins.fig"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(Initial.figsFile)
                        Case "basins.bsn"
                            TakeField.filename = Initial.Input_Files & "\" & Trim(Initial.BaseFile)
                        Case "file.cio"
                            TakeField.filename = Initial.Input_Files & "\" & "file.cio"
                    End Select

                    TakeField.Leng = .Rows(i).Item("Leng")
                    TakeField.LineNum = .Rows(i).Item("Lines")
                    TakeField.Inicia = .Rows(i).Item("Inicia")
                    value = TakeField.value()
                Else
                    value = .Rows(i).Item("Value")
                End If

                If Not IsDBNull(.Rows(i).Item("Format")) AndAlso .Rows(i).Item("Format") <> "" Then
                    lenFormat = Len(.Rows(i).Item("Format"))
                    roundformat = Right(Trim(.Rows(i).Item("Format")), 1)
                    NumberFormat = Left(Trim(.Rows(i).Item("Format")), lenFormat - 2)
                    value = TakeField.Convert(System.Math.Round(Val(value), roundformat), NumberFormat)
                End If

                ngnFlag = True
                Select Case .Rows(i).Item("Field") & .Rows(i).Item("Line")
                    Case CStr(61)
                        If value = "1" Then Initial.ngn = 2 'Temperature
                        ngnFlag = False
                    Case CStr(71)
                        If value = "1" Then Initial.ngn = Initial.ngn * 10 + 3 'Solar Radiation
                        ngnFlag = False
                    Case CStr(81)
                        If value = "1" Then Initial.ngn = Initial.ngn * 10 + 4 'Wind Speed
                        ngnFlag = False
                    Case CStr(91)
                        ngnFlag = False
                        If value = "1" Then Initial.ngn = Initial.ngn * 10 + 5 'Humidity simulation
                    Case CStr(101)
                        If Initial.ngn < 2 Then If Trim(value) = "1" Then Initial.ngn = 1 'Rainfall
                        'NumberFormat = "0"
                        value = TakeField.Convert(System.Math.Round(Initial.ngn, roundformat), NumberFormat)
                        ngnFlag = True
                    Case CStr(141)
                        Select Case Trim(value)
                            Case "0"
                                value = "   3"
                            Case "1"
                                value = "   1"
                            Case "2"
                                value = "   4"
                            Case "3"
                                value = "   4"
                        End Select
                End Select

                If i < .Rows.Count - 1 And ngnCreation = 0 Then
                    If ngnFlag = True Then
                        If current_line = .Rows(i + 1).Item("Line") Then
                            swFile1.Write(value)
                        Else
                            swFile1.WriteLine(value)
                        End If
                        current_line = .Rows(i + 1).Item("Line")
                    End If
                End If
            Next
        End With

        If ngnCreation = 0 Then
            swFile1.WriteLine()
            swFile1.WriteLine()

            swFile.Close()
            swFile.Dispose()
            swFile1.Close()
            swFile1.Dispose()
        End If

        Exit Sub
goError:
        MsgBox(Err.Description, , "Control Function")

    End Sub

    Public Sub YearSimulation()

        Dim TakeField As Convertion

        On Error GoTo goError

        TakeField = New Convertion

        Select Case Initial.Version
            Case "1.0.0", "2.0.0", "3.0.0", "4.0.0"
                TakeField.filename = Initial.Input_Files & "\" & Trim(Initial.CodFile)
                TakeField.LineNum = 2
            Case "1.1.0", "1.2.0", "1.3.0", "2.1.0", "2.3.0", "3.1.0", "4.1.0", "4.2.0", "4.3.0"
                TakeField.filename = Initial.Input_Files & "\" & "file.cio"
                TakeField.LineNum = 8
        End Select
        TakeField.Leng = 16
        TakeField.Inicia = 1
        Initial.YearSim = Val(TakeField.value())
        Exit Sub
goError:
        MsgBox(Err.Description, , "Control Function")

    End Sub

    Public Sub Pesticide()
        Dim PKOC As String
        Dim PWOF As String
        Dim PHLF As String
        Dim PHLS As String
        Dim PSOL As String
        Dim temp_Record As Object
        Dim z As StreamWriter

        Dim TakeField As Convertion
        Dim value As String
        Dim a As StreamReader

        On Error GoTo goError

        TakeField = New Convertion
        'fs = CreateObject("Scripting.FileSystemObject")
        z = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Pesticide))
        a = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & Initial.Pest))

        Do While a.EndOfStream <> True
            temp_Record = a.ReadLine
            z.Write("  ")
            z.Write(Left(temp_Record, 3))
            z.Write(" ")
            z.Write(Mid(temp_Record, 4, 16))
            PSOL = Mid(temp_Record, 57, 11)
            value = TakeField.Convert(Val(PSOL), "########.0")
            z.Write(value)
            PHLS = Mid(temp_Record, 44, 8)
            value = TakeField.Convert(Val(PHLS), "######.0")
            z.Write(value)
            PHLF = Mid(temp_Record, 36, 8)
            value = TakeField.Convert(Val(PHLF), "######.0")
            z.Write(value)
            PWOF = Mid(temp_Record, 31, 5)
            value = TakeField.Convert(Val(PWOF), "######.0")
            z.Write(value)
            PKOC = Mid(temp_Record, 21, 10)
            value = TakeField.Convert(Val(PKOC), "########.")
            z.WriteLine(value)
        Loop

        a.Close()
        a.Dispose()
        a = Nothing
        z.Close()
        z.Dispose()
        z = Nothing

        Exit Sub
goError:
        MsgBox(Err.Description, , "Control Function")

    End Sub

    Public Sub Fertilizer()
        Dim i As Object
        Dim Len1 As Object
        Dim value1 As Object
        Dim z As Object
        'create the APEX fertilizer file with the same information from SWAT fertilizer file.
        Dim TakeField As Convertion
        Dim value, temp As Object
        Dim name1 As String
        Dim fs, a As Object
        Dim valuesSplited() As String
        Dim code1 As Short

        On Error GoTo goError

        TakeField = New Convertion
        fs = CreateObject("Scripting.FileSystemObject")
        z = fs.CreateTextFile(Initial.Output_files & "\" & Initial.Fertilizer)
        a = fs.OpenTextFile(Initial.Input_Files & "\" & "Fert.dat")
        Do While a.atEndOfStream <> True
            temp = a.ReadLine
            ReDim valuesSplited(10)
            valuesSplited = Split(temp, "   ")
            code1 = CShort(Trim(Left(temp, 4)))
            name1 = Trim(Mid(temp, 6, 8))
            value1 = Mid(temp, 14, 16) & "   0.000" & Mid(temp, 30, 24) & "   0.350" & Mid(temp, 30, 16)
            z.Write(TakeField.Convert(code1, "####0")) 'Fertilizer code
            z.Write(" ")
            Len1 = Len(name1)
            Len1 = 8 - Len1
            z.Write(name1) 'Fertilizer Name
            For i = 1 To Len1
                z.Write(" ")
            Next
            z.WriteLine(value1)
        Loop
        Do While code1 <= 72
            code1 = code1 + 1

            z.Write(TakeField.Convert(code1, "####0"))
            z.Write(" ")
            Select Case code1
                Case 64
                    value1 = "Elem-N     1.000   0.000   0.000   0.000   0.000   0.000   0.000   0.000   0.000"
                Case 68
                    value1 = "82-00-00   0.820   0.000   0.000   0.000   0.000   0.000   0.000   0.000   0.447"
                Case 69
                    value1 = "DAP        0.180   0.460   0.000   0.000   0.000   0.990   0.350   0.000   0.000"
                Case 70
                    value1 = "DSMGAS     0.002   0.007   0.000   0.022   0.004   0.972   0.350   0.000   0.000"
                Case 71
                    value1 = "D-SO-MNU   0.002   0.007   0.000   0.022   0.004   0.972   0.350   0.000   0.000"
                Case 72
                    value1 = "D-LIQ      0.003   0.001   0.000   0.001   0.001   0.987   0.350   0.000   0.000"
                Case Else
                    value1 = " "
            End Select

            z.WriteLine(value1)
        Loop
        a.Close()
        z.Close()
        Exit Sub
goError:
        MsgBox(Err.Description, , "Control Function")

    End Sub
End Module