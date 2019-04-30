Option Strict Off

Imports System.Data
Imports System.Data.SqlClient
'Imports Microsoft.ApplicationBlocks.Data

Public Module AccessDB
    Private sDBConnectionDefault As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & SWAPP.Initial.Dir1 & "\Project_Parameters.mdb;"
    Private dbConnectString As String

    Public Function getDBDataSet(ByRef query As String) As Data.DataSet
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim myConnection1 As New OleDb.OleDbConnection
        Dim myConnection2 As New OleDb.OleDbConnection()
        Dim myConnection3 As New OleDb.OleDbConnection()
        Dim myConnection4 As New OleDb.OleDbConnection()
        Dim myConnection5 As New OleDb.OleDbConnection()
        Dim myConnection6 As New OleDb.OleDbConnection()
        Dim myConnection7 As New OleDb.OleDbConnection()

        Try
            dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & SWAPP.Initial.Dir1 & "\Project_Parameters.mdb;"
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection1.ConnectionString = dbConnectString
            If myConnection1.State = Data.ConnectionState.Closed Then
                myConnection1.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection1)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection2.ConnectionString = dbConnectString
            If myConnection2.State = Data.ConnectionState.Closed Then
                myConnection2.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection2)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection3.ConnectionString = dbConnectString
            If myConnection3.State = Data.ConnectionState.Closed Then
                myConnection3.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection3)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection4.ConnectionString = dbConnectString
            If myConnection4.State = Data.ConnectionState.Closed Then
                myConnection4.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection4)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection5.ConnectionString = dbConnectString
            If myConnection5.State = Data.ConnectionState.Closed Then
                myConnection5.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection5)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection6.ConnectionString = dbConnectString
            If myConnection6.State = Data.ConnectionState.Closed Then
                myConnection6.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection6)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection7.ConnectionString = dbConnectString
            If myConnection7.State = Data.ConnectionState.Closed Then
                myConnection7.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection7)
                ad.Fill(dsBas)
                Exit Try
            End If
        Catch ex As Exception
        End Try

        If myConnection1.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection2.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection3.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection4.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection5.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection6.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection7.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose() : myConnection1.Dispose() : myConnection2.Dispose() : myConnection3.Dispose()
        myConnection4.Dispose() : myConnection5.Dispose() : myConnection6.Dispose() : myConnection7.Dispose()

        myConnection = Nothing : myConnection1 = Nothing : myConnection2 = Nothing : myConnection3 = Nothing
        myConnection4 = Nothing : myConnection5 = Nothing : myConnection6 = Nothing : myConnection7 = Nothing

        Return dsBas

    End Function

    Public Function getData(ByRef query As String, myConnection As OleDb.OleDbConnection) As String
        Dim dsBas As New Data.DataSet()
        Try
            Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
            ad.Fill(dsBas)
            Exit Try
        Catch ex As Exception
        End Try
        If dsBas.Tables.Count <= 0 Then Return ""
        If dsBas.Tables(0).Rows.Count > 0 Then
            Return dsBas.Tables(0).Rows(0).Item(0)
        End If

        Return ""
    End Function

    Public Function getDBDataTable(ByRef query As String) As Data.DataTable
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim myConnection1 As New OleDb.OleDbConnection
        Dim myConnection2 As New OleDb.OleDbConnection()
        Dim myConnection3 As New OleDb.OleDbConnection()
        Dim myConnection4 As New OleDb.OleDbConnection()
        Dim myConnection5 As New OleDb.OleDbConnection()
        Dim myConnection6 As New OleDb.OleDbConnection()
        Dim myConnection7 As New OleDb.OleDbConnection()
        'Dim ad As New OleDb.OleDbDataAdapter

        Try
            dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & SWAPP.Initial.Dir1 & "\Project_Parameters.mdb;"
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection1.ConnectionString = dbConnectString
            If myConnection1.State = Data.ConnectionState.Closed Then
                myConnection1.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection1)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection2.ConnectionString = dbConnectString
            If myConnection2.State = Data.ConnectionState.Closed Then
                myConnection2.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection2)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection3.ConnectionString = dbConnectString
            If myConnection3.State = Data.ConnectionState.Closed Then
                myConnection3.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection3)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection4.ConnectionString = dbConnectString
            If myConnection4.State = Data.ConnectionState.Closed Then
                myConnection4.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection4)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection5.ConnectionString = dbConnectString
            If myConnection5.State = Data.ConnectionState.Closed Then
                myConnection5.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection5)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection6.ConnectionString = dbConnectString
            If myConnection6.State = Data.ConnectionState.Closed Then
                myConnection6.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection6)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection7.ConnectionString = dbConnectString
            If myConnection7.State = Data.ConnectionState.Closed Then
                myConnection7.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection7)
                ad.Fill(dsBas)
                Exit Try
            End If
        Catch ex As Exception
        End Try

        If myConnection1.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection2.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection3.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection4.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection5.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection6.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection7.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose() : myConnection1.Dispose() : myConnection2.Dispose() : myConnection3.Dispose()
        myConnection4.Dispose() : myConnection5.Dispose() : myConnection6.Dispose() : myConnection7.Dispose()

        myConnection = Nothing : myConnection1 = Nothing : myConnection2 = Nothing : myConnection3 = Nothing
        myConnection4 = Nothing : myConnection5 = Nothing : myConnection6 = Nothing : myConnection7 = Nothing

        If dsBas.Tables.Count > 0 Then
            Return dsBas.Tables(0)
        Else
            Return Nothing
        End If

    End Function

    Public Function getDBDataTableNoCon(ByRef query As String, myConnection As OleDb.OleDbConnection) As Data.DataTable
        Dim dsBas As New Data.DataSet()
        
        Try
            Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
            ad.Fill(dsBas)
            Return dsBas.Tables(0)

        Catch ex As Exception
        End Try


    End Function

    Public Function getLocalDataTable(ByRef query As String, ByVal path As String) As Data.DataTable
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim myConnection1 As New OleDb.OleDbConnection
        Dim myConnection2 As New OleDb.OleDbConnection()
        Dim myConnection3 As New OleDb.OleDbConnection()
        Dim myConnection4 As New OleDb.OleDbConnection()
        Dim myConnection5 As New OleDb.OleDbConnection()
        Dim myConnection6 As New OleDb.OleDbConnection()
        Dim myConnection7 As New OleDb.OleDbConnection()
        'Dim ad As New OleDb.OleDbDataAdapter

        Dim dbConnectString As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & path & "\Local.mdb;"

        Try
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection1.ConnectionString = dbConnectString
            If myConnection1.State = Data.ConnectionState.Closed Then
                myConnection1.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection1)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection2.ConnectionString = dbConnectString
            If myConnection2.State = Data.ConnectionState.Closed Then
                myConnection2.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection2)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection3.ConnectionString = dbConnectString
            If myConnection3.State = Data.ConnectionState.Closed Then
                myConnection3.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection3)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection4.ConnectionString = dbConnectString
            If myConnection4.State = Data.ConnectionState.Closed Then
                myConnection4.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection4)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection5.ConnectionString = dbConnectString
            If myConnection5.State = Data.ConnectionState.Closed Then
                myConnection5.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection5)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection6.ConnectionString = dbConnectString
            If myConnection6.State = Data.ConnectionState.Closed Then
                myConnection6.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection6)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection7.ConnectionString = dbConnectString
            If myConnection7.State = Data.ConnectionState.Closed Then
                myConnection7.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection7)
                ad.Fill(dsBas)
                Exit Try
            End If
        Catch ex As Exception
        End Try

        If myConnection1.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection2.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection3.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection4.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection5.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection6.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection7.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose() : myConnection1.Dispose() : myConnection2.Dispose() : myConnection3.Dispose()
        myConnection4.Dispose() : myConnection5.Dispose() : myConnection6.Dispose() : myConnection7.Dispose()

        myConnection = Nothing : myConnection1 = Nothing : myConnection2 = Nothing : myConnection3 = Nothing
        myConnection4 = Nothing : myConnection5 = Nothing : myConnection6 = Nothing : myConnection7 = Nothing

        If dsBas.Tables.Count > 0 Then
            Return dsBas.Tables(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetTableRecords(ByRef query As String, ByVal path As String) As UInteger
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim myConnection1 As New OleDb.OleDbConnection
        Dim myConnection2 As New OleDb.OleDbConnection()
        Dim myConnection3 As New OleDb.OleDbConnection()
        Dim myConnection4 As New OleDb.OleDbConnection()
        Dim myConnection5 As New OleDb.OleDbConnection()
        Dim myConnection6 As New OleDb.OleDbConnection()
        Dim myConnection7 As New OleDb.OleDbConnection()

        Dim dbConnectString As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & path & "\Local.mdb;"

        Try
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection1.ConnectionString = dbConnectString
            If myConnection1.State = Data.ConnectionState.Closed Then
                myConnection1.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection1)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection2.ConnectionString = dbConnectString
            If myConnection2.State = Data.ConnectionState.Closed Then
                myConnection2.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection2)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection3.ConnectionString = dbConnectString
            If myConnection3.State = Data.ConnectionState.Closed Then
                myConnection3.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection3)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection4.ConnectionString = dbConnectString
            If myConnection4.State = Data.ConnectionState.Closed Then
                myConnection4.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection4)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection5.ConnectionString = dbConnectString
            If myConnection5.State = Data.ConnectionState.Closed Then
                myConnection5.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection5)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection6.ConnectionString = dbConnectString
            If myConnection6.State = Data.ConnectionState.Closed Then
                myConnection6.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection6)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection7.ConnectionString = dbConnectString
            If myConnection7.State = Data.ConnectionState.Closed Then
                myConnection7.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection7)
                ad.Fill(dsBas)
                Exit Try
            End If
        Catch ex As Exception
        End Try

        If myConnection1.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection2.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection3.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection4.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection5.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection6.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection7.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose() : myConnection1.Dispose() : myConnection2.Dispose() : myConnection3.Dispose()
        myConnection4.Dispose() : myConnection5.Dispose() : myConnection6.Dispose() : myConnection7.Dispose()

        myConnection = Nothing : myConnection1 = Nothing : myConnection2 = Nothing : myConnection3 = Nothing
        myConnection4 = Nothing : myConnection5 = Nothing : myConnection6 = Nothing : myConnection7 = Nothing

        If dsBas.Tables.Count > 0 Then
            Return dsBas.Tables.Count
        Else
            Return 0
        End If
    End Function

    Public Function UpdateStringArray(ByRef query As String(), ByVal path As String) As String
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim i As UShort
        'Dim ad As New OleDb.OleDbDataAdapter

        Dim dbConnectString As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & path & "\Local.mdb;"

        Try
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                For i = 0 To query.Length - 1
                    Dim ad As New OleDb.OleDbDataAdapter(query(i), myConnection)
                    ad.Fill(dsBas)
                Next
                Exit Try
            End If

        Catch ex As Exception
        End Try

        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose()
        myConnection = Nothing
        Return "OK"

    End Function

    Public Function UpdateStringArray1(ByRef query As String(), ByVal path As String) As String
        Dim dsBas As New Data.DataSet()
        'Dim myConnection As New OleDb.OleDbConnection
        Dim i As UShort
        Dim con As New SqlConnection
        Dim da As New SqlDataAdapter
        'Dim ad As New OleDb.OleDbDataAdapter

        Dim dbConnectString As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & path & "\Local.mdb;"
        dbConnectString = "Data Source=" & path & "\Local.mdb;"

        Try
            'myConnection.ConnectionString = dbConnectString
            con.ConnectionString = dbConnectString
            If con.State = Data.ConnectionState.Closed Then
                con.Open()
                For i = 0 To query.Length - 1
                    Dim cm As New SqlCommand(query(i), con)
                    da.InsertCommand = cm
                    cm.ExecuteNonQuery()
                    'Dim ad As New OleDb.OleDbDataAdapter(query(i), myConnection)
                    'ad.Fill(dsBas)
                Next
                Exit Try
            End If

        Catch ex As Exception
        End Try

        If con.State = Data.ConnectionState.Open Then con.Close()

        con.Dispose()
        con = Nothing
        Return "OK"

    End Function

    Public Function getLocalDataSet(ByRef query As String, ByVal path As String) As Data.DataSet
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim myConnection1 As New OleDb.OleDbConnection
        Dim myConnection2 As New OleDb.OleDbConnection()
        Dim myConnection3 As New OleDb.OleDbConnection()
        Dim myConnection4 As New OleDb.OleDbConnection()
        Dim myConnection5 As New OleDb.OleDbConnection()
        Dim myConnection6 As New OleDb.OleDbConnection()
        Dim myConnection7 As New OleDb.OleDbConnection()
        'Dim ad As New OleDb.OleDbDataAdapter

        Dim dbConnectString As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & path & "\Local.mdb;"

        Try
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection1.ConnectionString = dbConnectString
            If myConnection1.State = Data.ConnectionState.Closed Then
                myConnection1.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection1)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection2.ConnectionString = dbConnectString
            If myConnection2.State = Data.ConnectionState.Closed Then
                myConnection2.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection2)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection3.ConnectionString = dbConnectString
            If myConnection3.State = Data.ConnectionState.Closed Then
                myConnection3.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection3)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection4.ConnectionString = dbConnectString
            If myConnection4.State = Data.ConnectionState.Closed Then
                myConnection4.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection4)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection5.ConnectionString = dbConnectString
            If myConnection5.State = Data.ConnectionState.Closed Then
                myConnection5.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection5)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection6.ConnectionString = dbConnectString
            If myConnection6.State = Data.ConnectionState.Closed Then
                myConnection6.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection6)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection7.ConnectionString = dbConnectString
            If myConnection7.State = Data.ConnectionState.Closed Then
                myConnection7.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection7)
                ad.Fill(dsBas)
                Exit Try
            End If
        Catch ex As Exception
        End Try

        If myConnection1.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection2.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection3.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection4.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection5.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection6.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection7.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose() : myConnection1.Dispose() : myConnection2.Dispose() : myConnection3.Dispose()
        myConnection4.Dispose() : myConnection5.Dispose() : myConnection6.Dispose() : myConnection7.Dispose()

        myConnection = Nothing : myConnection1 = Nothing : myConnection2 = Nothing : myConnection3 = Nothing
        myConnection4 = Nothing : myConnection5 = Nothing : myConnection6 = Nothing : myConnection7 = Nothing

        Return dsBas

    End Function

    Public Sub modifyRecords(ByRef query As String)
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim myConnection1 As New OleDb.OleDbConnection
        Dim myConnection2 As New OleDb.OleDbConnection()
        Dim myConnection3 As New OleDb.OleDbConnection()
        Dim myConnection4 As New OleDb.OleDbConnection()
        Dim myConnection5 As New OleDb.OleDbConnection()
        Dim myConnection6 As New OleDb.OleDbConnection()
        Dim myConnection7 As New OleDb.OleDbConnection()
        'Dim ad As New OleDb.OleDbDataAdapter

        Try
            dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & Initial.Dir1 & "\Project_Parameters.mdb;"
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection1.ConnectionString = dbConnectString
            If myConnection1.State = Data.ConnectionState.Closed Then
                myConnection1.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection1)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection2.ConnectionString = dbConnectString
            If myConnection2.State = Data.ConnectionState.Closed Then
                myConnection2.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection2)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection3.ConnectionString = dbConnectString
            If myConnection3.State = Data.ConnectionState.Closed Then
                myConnection3.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection3)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection4.ConnectionString = dbConnectString
            If myConnection4.State = Data.ConnectionState.Closed Then
                myConnection4.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection4)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection5.ConnectionString = dbConnectString
            If myConnection5.State = Data.ConnectionState.Closed Then
                myConnection5.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection5)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection6.ConnectionString = dbConnectString
            If myConnection6.State = Data.ConnectionState.Closed Then
                myConnection6.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection6)
                ad.Fill(dsBas)
                Exit Try
            End If
            myConnection7.ConnectionString = dbConnectString
            If myConnection7.State = Data.ConnectionState.Closed Then
                myConnection7.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection7)
                ad.Fill(dsBas)
                Exit Try
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " - " & query, MsgBoxStyle.OkOnly, "Modified Table/ModifyRecords")
        End Try

        If myConnection1.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection2.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection3.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection4.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection5.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection6.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection7.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose() : myConnection1.Dispose() : myConnection2.Dispose() : myConnection3.Dispose()
        myConnection4.Dispose() : myConnection5.Dispose() : myConnection6.Dispose() : myConnection7.Dispose()

        myConnection = Nothing : myConnection1 = Nothing : myConnection2 = Nothing : myConnection3 = Nothing
        myConnection4 = Nothing : myConnection5 = Nothing : myConnection6 = Nothing : myConnection7 = Nothing

    End Sub

    Public Sub modifyLocalRecords(ByRef query As String, ByVal path As String)
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim myConnection1 As New OleDb.OleDbConnection
        Dim myConnection2 As New OleDb.OleDbConnection()
        Dim myConnection3 As New OleDb.OleDbConnection()
        Dim myConnection4 As New OleDb.OleDbConnection()
        Dim myConnection5 As New OleDb.OleDbConnection()
        Dim myConnection6 As New OleDb.OleDbConnection()
        Dim myConnection7 As New OleDb.OleDbConnection()
        'Dim ad As New OleDb.OleDbDataAdapter

        Dim dbConnectString As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & path & "\Local.mdb;"

        Try
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection1.ConnectionString = dbConnectString
            If myConnection1.State = Data.ConnectionState.Closed Then
                myConnection1.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection1)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection2.ConnectionString = dbConnectString
            If myConnection2.State = Data.ConnectionState.Closed Then
                myConnection2.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection2)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection3.ConnectionString = dbConnectString
            If myConnection3.State = Data.ConnectionState.Closed Then
                myConnection3.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection3)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection4.ConnectionString = dbConnectString
            If myConnection4.State = Data.ConnectionState.Closed Then
                myConnection4.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection4)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection5.ConnectionString = dbConnectString
            If myConnection5.State = Data.ConnectionState.Closed Then
                myConnection5.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection5)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection6.ConnectionString = dbConnectString
            If myConnection6.State = Data.ConnectionState.Closed Then
                myConnection6.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection6)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection7.ConnectionString = dbConnectString
            If myConnection7.State = Data.ConnectionState.Closed Then
                myConnection7.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection7)
                ad.Fill(dsBas)
                Exit Try
            End If
        Catch ex As Exception
        End Try

        If myConnection1.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection2.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection3.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection4.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection5.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection6.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection7.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose() : myConnection1.Dispose() : myConnection2.Dispose() : myConnection3.Dispose()
        myConnection4.Dispose() : myConnection5.Dispose() : myConnection6.Dispose() : myConnection7.Dispose()

        myConnection = Nothing : myConnection1 = Nothing : myConnection2 = Nothing : myConnection3 = Nothing
        myConnection4 = Nothing : myConnection5 = Nothing : myConnection6 = Nothing : myConnection7 = Nothing
    End Sub

    Public Sub modifyFEMRecords(ByRef query As String, ByVal path As String)
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim myConnection1 As New OleDb.OleDbConnection
        Dim myConnection2 As New OleDb.OleDbConnection()
        Dim myConnection3 As New OleDb.OleDbConnection()
        Dim myConnection4 As New OleDb.OleDbConnection()
        Dim myConnection5 As New OleDb.OleDbConnection()
        Dim myConnection6 As New OleDb.OleDbConnection()
        Dim myConnection7 As New OleDb.OleDbConnection()
        'Dim ad As New OleDb.OleDbDataAdapter

        Dim dbConnectString As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & path & "\SWAPPFEMOut.mdb"

        Try
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection1.ConnectionString = dbConnectString
            If myConnection1.State = Data.ConnectionState.Closed Then
                myConnection1.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection1)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection2.ConnectionString = dbConnectString
            If myConnection2.State = Data.ConnectionState.Closed Then
                myConnection2.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection2)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection3.ConnectionString = dbConnectString
            If myConnection3.State = Data.ConnectionState.Closed Then
                myConnection3.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection3)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection4.ConnectionString = dbConnectString
            If myConnection4.State = Data.ConnectionState.Closed Then
                myConnection4.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection4)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection5.ConnectionString = dbConnectString
            If myConnection5.State = Data.ConnectionState.Closed Then
                myConnection5.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection5)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection6.ConnectionString = dbConnectString
            If myConnection6.State = Data.ConnectionState.Closed Then
                myConnection6.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection6)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection7.ConnectionString = dbConnectString
            If myConnection7.State = Data.ConnectionState.Closed Then
                myConnection7.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection7)
                ad.Fill(dsBas)
                Exit Try
            End If
        Catch ex As Exception
        End Try

        If myConnection1.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection2.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection3.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection4.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection5.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection6.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection7.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose() : myConnection1.Dispose() : myConnection2.Dispose() : myConnection3.Dispose()
        myConnection4.Dispose() : myConnection5.Dispose() : myConnection6.Dispose() : myConnection7.Dispose()

        myConnection = Nothing : myConnection1 = Nothing : myConnection2 = Nothing : myConnection3 = Nothing
        myConnection4 = Nothing : myConnection5 = Nothing : myConnection6 = Nothing : myConnection7 = Nothing
    End Sub

    Public Function ParmDBName(ByRef SQLString As String) As String
        Dim parmDB As DataTable

        parmDB = New DataTable
        parmDB = getDBDataTable(SQLString)
        ParmDBName = ""
        If parmDB.Rows.Count > 0 Then
            If Not parmDB.Rows(0).Item(0) Is System.DBNull.Value Then ParmDBName = parmDB.Rows(0).Item(0)
        End If

        parmDB.Dispose()
        parmDB = Nothing

        Return ParmDBName
    End Function
    Public Function getFEMDataTable(ByRef query As String, ByVal path As String) As Data.DataTable
        Dim dsBas As New Data.DataSet()
        Dim myConnection As New OleDb.OleDbConnection
        Dim myConnection1 As New OleDb.OleDbConnection
        Dim myConnection2 As New OleDb.OleDbConnection()
        Dim myConnection3 As New OleDb.OleDbConnection()
        Dim myConnection4 As New OleDb.OleDbConnection()
        Dim myConnection5 As New OleDb.OleDbConnection()
        Dim myConnection6 As New OleDb.OleDbConnection()
        Dim myConnection7 As New OleDb.OleDbConnection()
        'Dim ad As New OleDb.OleDbDataAdapter

        Dim dbConnectString As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & path & "\SWAPPFEMOut.mdb;"

        Try
            myConnection.ConnectionString = dbConnectString
            If myConnection.State = Data.ConnectionState.Closed Then
                myConnection.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection1.ConnectionString = dbConnectString
            If myConnection1.State = Data.ConnectionState.Closed Then
                myConnection1.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection1)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection2.ConnectionString = dbConnectString
            If myConnection2.State = Data.ConnectionState.Closed Then
                myConnection2.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection2)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection3.ConnectionString = dbConnectString
            If myConnection3.State = Data.ConnectionState.Closed Then
                myConnection3.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection3)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection4.ConnectionString = dbConnectString
            If myConnection4.State = Data.ConnectionState.Closed Then
                myConnection4.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection4)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection5.ConnectionString = dbConnectString
            If myConnection5.State = Data.ConnectionState.Closed Then
                myConnection5.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection5)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection6.ConnectionString = dbConnectString
            If myConnection6.State = Data.ConnectionState.Closed Then
                myConnection6.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection6)
                ad.Fill(dsBas)
                Exit Try
            End If

            myConnection7.ConnectionString = dbConnectString
            If myConnection7.State = Data.ConnectionState.Closed Then
                myConnection7.Open()
                Dim ad As New OleDb.OleDbDataAdapter(query, myConnection7)
                ad.Fill(dsBas)
                Exit Try
            End If
        Catch ex As Exception
        End Try

        If myConnection1.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection2.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection3.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection4.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection5.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection6.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection7.State = Data.ConnectionState.Open Then myConnection.Close()
        If myConnection.State = Data.ConnectionState.Open Then myConnection.Close()

        myConnection.Dispose() : myConnection1.Dispose() : myConnection2.Dispose() : myConnection3.Dispose()
        myConnection4.Dispose() : myConnection5.Dispose() : myConnection6.Dispose() : myConnection7.Dispose()

        myConnection = Nothing : myConnection1 = Nothing : myConnection2 = Nothing : myConnection3 = Nothing
        myConnection4 = Nothing : myConnection5 = Nothing : myConnection6 = Nothing : myConnection7 = Nothing

        If dsBas.Tables.Count > 0 Then
            Return dsBas.Tables(0)
        Else
            Return Nothing
        End If
    End Function
End Module
