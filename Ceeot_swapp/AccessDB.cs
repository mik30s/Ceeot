using System;
using Microsoft.VisualBasic;
using System.Data;
using System.Data.SqlClient;
// Imports Microsoft.ApplicationBlocks.Data
using MsgBox = System.Windows.Forms.MessageBox;

public static class AccessDB
{
    private static string sDBConnectionDefault = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= resources/databases/Project_Parameters.mdb;";
    private static string dbConnectString;

    public static System.Data.DataSet getDBDataSet(string query)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection1 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection2 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection3 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection4 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection5 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection6 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection7 = new System.Data.OleDb.OleDbConnection();

        try
        {
            dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= resources/databases/Project_Parameters.mdb;";
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
                ad.Fill(dsBas);
                ;
            }

            myConnection1.ConnectionString = dbConnectString;
            if (myConnection1.State == System.Data.ConnectionState.Closed)
            {
                myConnection1.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection1);
                ad.Fill(dsBas);
                ;
            }

            myConnection2.ConnectionString = dbConnectString;
            if (myConnection2.State == System.Data.ConnectionState.Closed)
            {
                myConnection2.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection2);
                ad.Fill(dsBas);
                ;
            }

            myConnection3.ConnectionString = dbConnectString;
            if (myConnection3.State == System.Data.ConnectionState.Closed)
            {
                myConnection3.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection3);
                ad.Fill(dsBas);
                ;
            }

            myConnection4.ConnectionString = dbConnectString;
            if (myConnection4.State == System.Data.ConnectionState.Closed)
            {
                myConnection4.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection4);
                ad.Fill(dsBas);
                ;
            }

            myConnection5.ConnectionString = dbConnectString;
            if (myConnection5.State == System.Data.ConnectionState.Closed)
            {
                myConnection5.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection5);
                ad.Fill(dsBas);
                ;
            }

            myConnection6.ConnectionString = dbConnectString;
            if (myConnection6.State == System.Data.ConnectionState.Closed)
            {
                myConnection6.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection6);
                ad.Fill(dsBas);
                ;
            }

            myConnection7.ConnectionString = dbConnectString;
            if (myConnection7.State == System.Data.ConnectionState.Closed)
            {
                myConnection7.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection7);
                ad.Fill(dsBas);
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (myConnection1.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection2.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection3.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection4.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection5.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection6.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection7.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose(); myConnection1.Dispose(); myConnection2.Dispose(); myConnection3.Dispose();
        myConnection4.Dispose(); myConnection5.Dispose(); myConnection6.Dispose(); myConnection7.Dispose();

        myConnection = null; myConnection1 = null; myConnection2 = null; myConnection3 = null;
        myConnection4 = null; myConnection5 = null; myConnection6 = null; myConnection7 = null;

        return dsBas;
    }

    public static string getData(ref string query, System.Data.OleDb.OleDbConnection myConnection)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        try
        {
            System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
            ad.Fill(dsBas);
           
        }
        catch (Exception ex)
        {
        }
        if (dsBas.Tables.Count <= 0)
            return "";
        if (dsBas.Tables[0].Rows.Count > 0)
            return (string)dsBas.Tables[0].Rows[0].ItemArray[0];

        return "";
    }

    public static System.Data.DataTable getDBDataTable(string query)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection1 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection2 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection3 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection4 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection5 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection6 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection7 = new System.Data.OleDb.OleDbConnection();
        // Dim ad As New OleDb.OleDbDataAdapter

        try
        {
            dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= resources/databases/Project_Parameters.mdb;";
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
                ad.Fill(dsBas);
                ;
            }

            myConnection1.ConnectionString = dbConnectString;
            if (myConnection1.State == System.Data.ConnectionState.Closed)
            {
                myConnection1.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection1);
                ad.Fill(dsBas);
                ;
            }

            myConnection2.ConnectionString = dbConnectString;
            if (myConnection2.State == System.Data.ConnectionState.Closed)
            {
                myConnection2.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection2);
                ad.Fill(dsBas);
                ;
            }

            myConnection3.ConnectionString = dbConnectString;
            if (myConnection3.State == System.Data.ConnectionState.Closed)
            {
                myConnection3.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection3);
                ad.Fill(dsBas);
                ;
            }

            myConnection4.ConnectionString = dbConnectString;
            if (myConnection4.State == System.Data.ConnectionState.Closed)
            {
                myConnection4.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection4);
                ad.Fill(dsBas);
                ;
            }

            myConnection5.ConnectionString = dbConnectString;
            if (myConnection5.State == System.Data.ConnectionState.Closed)
            {
                myConnection5.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection5);
                ad.Fill(dsBas);
                ;
            }

            myConnection6.ConnectionString = dbConnectString;
            if (myConnection6.State == System.Data.ConnectionState.Closed)
            {
                myConnection6.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection6);
                ad.Fill(dsBas);
                ;
            }

            myConnection7.ConnectionString = dbConnectString;
            if (myConnection7.State == System.Data.ConnectionState.Closed)
            {
                myConnection7.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection7);
                ad.Fill(dsBas);
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (myConnection1.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection2.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection3.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection4.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection5.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection6.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection7.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose(); myConnection1.Dispose(); myConnection2.Dispose(); myConnection3.Dispose();
        myConnection4.Dispose(); myConnection5.Dispose(); myConnection6.Dispose(); myConnection7.Dispose();

        myConnection = null; myConnection1 = null; myConnection2 = null; myConnection3 = null;
        myConnection4 = null; myConnection5 = null; myConnection6 = null; myConnection7 = null;

        if (dsBas.Tables.Count > 0)
            return dsBas.Tables[0];
        else
            return null;
    }

    public static System.Data.DataTable getDBDataTableNoCon(ref string query, System.Data.OleDb.OleDbConnection myConnection)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();

        try
        {
            System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
            ad.Fill(dsBas);
        }
        catch (Exception ex)
        {
        }
        return dsBas.Tables[0];
    }

    public static System.Data.DataTable getLocalDataTable(ref string query, string path)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection1 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection2 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection3 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection4 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection5 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection6 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection7 = new System.Data.OleDb.OleDbConnection();
        // Dim ad As New OleDb.OleDbDataAdapter

        string dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " + path + @"\Local.mdb;";

        try
        {
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
                ad.Fill(dsBas);
                ;
            }

            myConnection1.ConnectionString = dbConnectString;
            if (myConnection1.State == System.Data.ConnectionState.Closed)
            {
                myConnection1.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection1);
                ad.Fill(dsBas);
                ;
            }

            myConnection2.ConnectionString = dbConnectString;
            if (myConnection2.State == System.Data.ConnectionState.Closed)
            {
                myConnection2.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection2);
                ad.Fill(dsBas);
                ;
            }

            myConnection3.ConnectionString = dbConnectString;
            if (myConnection3.State == System.Data.ConnectionState.Closed)
            {
                myConnection3.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection3);
                ad.Fill(dsBas);
                ;
            }

            myConnection4.ConnectionString = dbConnectString;
            if (myConnection4.State == System.Data.ConnectionState.Closed)
            {
                myConnection4.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection4);
                ad.Fill(dsBas);
                ;
            }

            myConnection5.ConnectionString = dbConnectString;
            if (myConnection5.State == System.Data.ConnectionState.Closed)
            {
                myConnection5.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection5);
                ad.Fill(dsBas);
                ;
            }

            myConnection6.ConnectionString = dbConnectString;
            if (myConnection6.State == System.Data.ConnectionState.Closed)
            {
                myConnection6.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection6);
                ad.Fill(dsBas);
                ;
            }

            myConnection7.ConnectionString = dbConnectString;
            if (myConnection7.State == System.Data.ConnectionState.Closed)
            {
                myConnection7.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection7);
                ad.Fill(dsBas);
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (myConnection1.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection2.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection3.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection4.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection5.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection6.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection7.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose(); myConnection1.Dispose(); myConnection2.Dispose(); myConnection3.Dispose();
        myConnection4.Dispose(); myConnection5.Dispose(); myConnection6.Dispose(); myConnection7.Dispose();

        myConnection = null; myConnection1 = null; myConnection2 = null; myConnection3 = null;
        myConnection4 = null; myConnection5 = null; myConnection6 = null; myConnection7 = null;

        if (dsBas.Tables.Count > 0)
            return dsBas.Tables[0];
        else
            return null;
    }

    public static uint GetTableRecords(ref string query, string path)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection1 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection2 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection3 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection4 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection5 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection6 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection7 = new System.Data.OleDb.OleDbConnection();

        string dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " + path + @"\Local.mdb;";

        try
        {
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
                ad.Fill(dsBas);
                ;
            }

            myConnection1.ConnectionString = dbConnectString;
            if (myConnection1.State == System.Data.ConnectionState.Closed)
            {
                myConnection1.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection1);
                ad.Fill(dsBas);
                ;
            }

            myConnection2.ConnectionString = dbConnectString;
            if (myConnection2.State == System.Data.ConnectionState.Closed)
            {
                myConnection2.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection2);
                ad.Fill(dsBas);
                ;
            }

            myConnection3.ConnectionString = dbConnectString;
            if (myConnection3.State == System.Data.ConnectionState.Closed)
            {
                myConnection3.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection3);
                ad.Fill(dsBas);
                ;
            }

            myConnection4.ConnectionString = dbConnectString;
            if (myConnection4.State == System.Data.ConnectionState.Closed)
            {
                myConnection4.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection4);
                ad.Fill(dsBas);
                ;
            }

            myConnection5.ConnectionString = dbConnectString;
            if (myConnection5.State == System.Data.ConnectionState.Closed)
            {
                myConnection5.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection5);
                ad.Fill(dsBas);
                ;
            }

            myConnection6.ConnectionString = dbConnectString;
            if (myConnection6.State == System.Data.ConnectionState.Closed)
            {
                myConnection6.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection6);
                ad.Fill(dsBas);
                ;
            }

            myConnection7.ConnectionString = dbConnectString;
            if (myConnection7.State == System.Data.ConnectionState.Closed)
            {
                myConnection7.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection7);
                ad.Fill(dsBas);
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (myConnection1.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection2.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection3.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection4.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection5.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection6.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection7.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose(); myConnection1.Dispose(); myConnection2.Dispose(); myConnection3.Dispose();
        myConnection4.Dispose(); myConnection5.Dispose(); myConnection6.Dispose(); myConnection7.Dispose();

        myConnection = null; myConnection1 = null; myConnection2 = null; myConnection3 = null;
        myConnection4 = null; myConnection5 = null; myConnection6 = null; myConnection7 = null;

        if (dsBas.Tables.Count > 0)
            return (uint)dsBas.Tables.Count;
        else
            return 0;
    }

    public static string UpdateStringArray(ref string[] query, string path)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        ushort i;
        // Dim ad As New OleDb.OleDbDataAdapter

        string dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " + path + @"\Local.mdb;";

        try
        {
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                for (i = 0; i <= query.Length - 1; i++)
                {
                    System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query[i], myConnection);
                    ad.Fill(dsBas);
                }
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose();
        myConnection = null;
        return "OK";
    }

    public static string UpdateStringArray1(ref string[] query, string path)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        // Dim myConnection As New OleDb.OleDbConnection
        ushort i;
        SqlConnection con = new SqlConnection();
        SqlDataAdapter da = new SqlDataAdapter();
        // Dim ad As New OleDb.OleDbDataAdapter

        string dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " + path + @"\Local.mdb;";
        dbConnectString = "Data Source=" + path + @"\Local.mdb;";

        try
        {
            // myConnection.ConnectionString = dbConnectString
            con.ConnectionString = dbConnectString;
            if (con.State == System.Data.ConnectionState.Closed)
            {
                con.Open();
                for (i = 0; i <= query.Length - 1; i++)
                {
                    SqlCommand cm = new SqlCommand(query[i], con);
                    da.InsertCommand = cm;
                    cm.ExecuteNonQuery();
                }
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (con.State == System.Data.ConnectionState.Open)
            con.Close();

        con.Dispose();
        con = null;
        return "OK";
    }

    public static System.Data.DataSet getLocalDataSet(ref string query, string path)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection1 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection2 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection3 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection4 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection5 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection6 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection7 = new System.Data.OleDb.OleDbConnection();
        // Dim ad As New OleDb.OleDbDataAdapter

        string dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " + path + @"\Local.mdb;";

        try
        {
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
                ad.Fill(dsBas);
                ;
            }

            myConnection1.ConnectionString = dbConnectString;
            if (myConnection1.State == System.Data.ConnectionState.Closed)
            {
                myConnection1.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection1);
                ad.Fill(dsBas);
                ;
            }

            myConnection2.ConnectionString = dbConnectString;
            if (myConnection2.State == System.Data.ConnectionState.Closed)
            {
                myConnection2.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection2);
                ad.Fill(dsBas);
                ;
            }

            myConnection3.ConnectionString = dbConnectString;
            if (myConnection3.State == System.Data.ConnectionState.Closed)
            {
                myConnection3.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection3);
                ad.Fill(dsBas);
                ;
            }

            myConnection4.ConnectionString = dbConnectString;
            if (myConnection4.State == System.Data.ConnectionState.Closed)
            {
                myConnection4.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection4);
                ad.Fill(dsBas);
                ;
            }

            myConnection5.ConnectionString = dbConnectString;
            if (myConnection5.State == System.Data.ConnectionState.Closed)
            {
                myConnection5.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection5);
                ad.Fill(dsBas);
                ;
            }

            myConnection6.ConnectionString = dbConnectString;
            if (myConnection6.State == System.Data.ConnectionState.Closed)
            {
                myConnection6.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection6);
                ad.Fill(dsBas);
                ;
            }

            myConnection7.ConnectionString = dbConnectString;
            if (myConnection7.State == System.Data.ConnectionState.Closed)
            {
                myConnection7.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection7);
                ad.Fill(dsBas);
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (myConnection1.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection2.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection3.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection4.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection5.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection6.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection7.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose(); myConnection1.Dispose(); myConnection2.Dispose(); myConnection3.Dispose();
        myConnection4.Dispose(); myConnection5.Dispose(); myConnection6.Dispose(); myConnection7.Dispose();

        myConnection = null; myConnection1 = null; myConnection2 = null; myConnection3 = null;
        myConnection4 = null; myConnection5 = null; myConnection6 = null; myConnection7 = null;

        return dsBas;
    }

    public static void modifyRecords(ref string query)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection1 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection2 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection3 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection4 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection5 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection6 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection7 = new System.Data.OleDb.OleDbConnection();
        // Dim ad As New OleDb.OleDbDataAdapter

        try
        {
            dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= resources/databases/Project_Parameters.mdb;";
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
                ad.Fill(dsBas);
                ;
            }

            myConnection1.ConnectionString = dbConnectString;
            if (myConnection1.State == System.Data.ConnectionState.Closed)
            {
                myConnection1.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection1);
                ad.Fill(dsBas);
                ;
            }

            myConnection2.ConnectionString = dbConnectString;
            if (myConnection2.State == System.Data.ConnectionState.Closed)
            {
                myConnection2.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection2);
                ad.Fill(dsBas);
                ;
            }

            myConnection3.ConnectionString = dbConnectString;
            if (myConnection3.State == System.Data.ConnectionState.Closed)
            {
                myConnection3.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection3);
                ad.Fill(dsBas);
                ;
            }

            myConnection4.ConnectionString = dbConnectString;
            if (myConnection4.State == System.Data.ConnectionState.Closed)
            {
                myConnection4.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection4);
                ad.Fill(dsBas);
                ;
            }

            myConnection5.ConnectionString = dbConnectString;
            if (myConnection5.State == System.Data.ConnectionState.Closed)
            {
                myConnection5.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection5);
                ad.Fill(dsBas);
                ;
            }

            myConnection6.ConnectionString = dbConnectString;
            if (myConnection6.State == System.Data.ConnectionState.Closed)
            {
                myConnection6.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection6);
                ad.Fill(dsBas);
                ;
            }
            myConnection7.ConnectionString = dbConnectString;
            if (myConnection7.State == System.Data.ConnectionState.Closed)
            {
                myConnection7.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection7);
                ad.Fill(dsBas);
                ;
            }
        }
        catch (Exception ex)
        {
            //System.Windows.Forms.MessageBox.Show(ex.Message + " - " + query, MsgBoxStyle.OkOnly, "Modified Table/ModifyRecords");
        }

        if (myConnection1.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection2.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection3.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection4.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection5.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection6.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection7.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose(); myConnection1.Dispose(); myConnection2.Dispose(); myConnection3.Dispose();
        myConnection4.Dispose(); myConnection5.Dispose(); myConnection6.Dispose(); myConnection7.Dispose();

        myConnection = null; myConnection1 = null; myConnection2 = null; myConnection3 = null;
        myConnection4 = null; myConnection5 = null; myConnection6 = null; myConnection7 = null;
    }

    public static void modifyLocalRecords(ref string query, string path)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection1 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection2 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection3 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection4 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection5 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection6 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection7 = new System.Data.OleDb.OleDbConnection();
        // Dim ad As New OleDb.OleDbDataAdapter

        string dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " + path + @"\Local.mdb;";

        try
        {
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
                ad.Fill(dsBas);
                ;
            }

            myConnection1.ConnectionString = dbConnectString;
            if (myConnection1.State == System.Data.ConnectionState.Closed)
            {
                myConnection1.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection1);
                ad.Fill(dsBas);
                ;
            }

            myConnection2.ConnectionString = dbConnectString;
            if (myConnection2.State == System.Data.ConnectionState.Closed)
            {
                myConnection2.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection2);
                ad.Fill(dsBas);
                ;
            }

            myConnection3.ConnectionString = dbConnectString;
            if (myConnection3.State == System.Data.ConnectionState.Closed)
            {
                myConnection3.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection3);
                ad.Fill(dsBas);
                ;
            }

            myConnection4.ConnectionString = dbConnectString;
            if (myConnection4.State == System.Data.ConnectionState.Closed)
            {
                myConnection4.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection4);
                ad.Fill(dsBas);
                ;
            }

            myConnection5.ConnectionString = dbConnectString;
            if (myConnection5.State == System.Data.ConnectionState.Closed)
            {
                myConnection5.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection5);
                ad.Fill(dsBas);
                ;
            }

            myConnection6.ConnectionString = dbConnectString;
            if (myConnection6.State == System.Data.ConnectionState.Closed)
            {
                myConnection6.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection6);
                ad.Fill(dsBas);
                ;
            }

            myConnection7.ConnectionString = dbConnectString;
            if (myConnection7.State == System.Data.ConnectionState.Closed)
            {
                myConnection7.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection7);
                ad.Fill(dsBas);
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (myConnection1.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection2.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection3.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection4.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection5.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection6.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection7.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose(); myConnection1.Dispose(); myConnection2.Dispose(); myConnection3.Dispose();
        myConnection4.Dispose(); myConnection5.Dispose(); myConnection6.Dispose(); myConnection7.Dispose();

        myConnection = null; myConnection1 = null; myConnection2 = null; myConnection3 = null;
        myConnection4 = null; myConnection5 = null; myConnection6 = null; myConnection7 = null;
    }

    public static void modifyFEMRecords(ref string query, string path)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection1 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection2 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection3 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection4 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection5 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection6 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection7 = new System.Data.OleDb.OleDbConnection();
        // Dim ad As New OleDb.OleDbDataAdapter

        string dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " + path + @"\SWAPPFEMOut.mdb";

        try
        {
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
                ad.Fill(dsBas);
                ;
            }

            myConnection1.ConnectionString = dbConnectString;
            if (myConnection1.State == System.Data.ConnectionState.Closed)
            {
                myConnection1.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection1);
                ad.Fill(dsBas);
                ;
            }

            myConnection2.ConnectionString = dbConnectString;
            if (myConnection2.State == System.Data.ConnectionState.Closed)
            {
                myConnection2.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection2);
                ad.Fill(dsBas);
                ;
            }

            myConnection3.ConnectionString = dbConnectString;
            if (myConnection3.State == System.Data.ConnectionState.Closed)
            {
                myConnection3.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection3);
                ad.Fill(dsBas);
                ;
            }

            myConnection4.ConnectionString = dbConnectString;
            if (myConnection4.State == System.Data.ConnectionState.Closed)
            {
                myConnection4.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection4);
                ad.Fill(dsBas);
                ;
            }

            myConnection5.ConnectionString = dbConnectString;
            if (myConnection5.State == System.Data.ConnectionState.Closed)
            {
                myConnection5.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection5);
                ad.Fill(dsBas);
                ;
            }

            myConnection6.ConnectionString = dbConnectString;
            if (myConnection6.State == System.Data.ConnectionState.Closed)
            {
                myConnection6.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection6);
                ad.Fill(dsBas);
                ;
            }

            myConnection7.ConnectionString = dbConnectString;
            if (myConnection7.State == System.Data.ConnectionState.Closed)
            {
                myConnection7.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection7);
                ad.Fill(dsBas);
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (myConnection1.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection2.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection3.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection4.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection5.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection6.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection7.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose(); myConnection1.Dispose(); myConnection2.Dispose(); myConnection3.Dispose();
        myConnection4.Dispose(); myConnection5.Dispose(); myConnection6.Dispose(); myConnection7.Dispose();

        myConnection = null; myConnection1 = null; myConnection2 = null; myConnection3 = null;
        myConnection4 = null; myConnection5 = null; myConnection6 = null; myConnection7 = null;
    }

    public static string ParmDBName(ref string SQLString)
    {
        DataTable parmDB;

        parmDB = new DataTable();
        parmDB = getDBDataTable(ref SQLString);
        if (parmDB.Rows.Count > 0)
        {
            var v = (string)parmDB.Rows[0].ItemArray[0];
            if (v != "")
                return (string)parmDB.Rows[0].ItemArray[0];
        }

        parmDB.Dispose();
        parmDB = null;

        return "";
    }
    public static System.Data.DataTable getFEMDataTable(ref string query, string path)
    {
        System.Data.DataSet dsBas = new System.Data.DataSet();
        System.Data.OleDb.OleDbConnection myConnection = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection1 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection2 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection3 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection4 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection5 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection6 = new System.Data.OleDb.OleDbConnection();
        System.Data.OleDb.OleDbConnection myConnection7 = new System.Data.OleDb.OleDbConnection();
        // Dim ad As New OleDb.OleDbDataAdapter

        string dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " + path + @"\SWAPPFEMOut.mdb;";

        try
        {
            myConnection.ConnectionString = dbConnectString;
            if (myConnection.State == System.Data.ConnectionState.Closed)
            {
                myConnection.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection);
                ad.Fill(dsBas);
                ;
            }

            myConnection1.ConnectionString = dbConnectString;
            if (myConnection1.State == System.Data.ConnectionState.Closed)
            {
                myConnection1.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection1);
                ad.Fill(dsBas);
                ;
            }

            myConnection2.ConnectionString = dbConnectString;
            if (myConnection2.State == System.Data.ConnectionState.Closed)
            {
                myConnection2.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection2);
                ad.Fill(dsBas);
                ;
            }

            myConnection3.ConnectionString = dbConnectString;
            if (myConnection3.State == System.Data.ConnectionState.Closed)
            {
                myConnection3.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection3);
                ad.Fill(dsBas);
                ;
            }

            myConnection4.ConnectionString = dbConnectString;
            if (myConnection4.State == System.Data.ConnectionState.Closed)
            {
                myConnection4.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection4);
                ad.Fill(dsBas);
                ;
            }

            myConnection5.ConnectionString = dbConnectString;
            if (myConnection5.State == System.Data.ConnectionState.Closed)
            {
                myConnection5.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection5);
                ad.Fill(dsBas);
                ;
            }

            myConnection6.ConnectionString = dbConnectString;
            if (myConnection6.State == System.Data.ConnectionState.Closed)
            {
                myConnection6.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection6);
                ad.Fill(dsBas);
                ;
            }

            myConnection7.ConnectionString = dbConnectString;
            if (myConnection7.State == System.Data.ConnectionState.Closed)
            {
                myConnection7.Open();
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(query, myConnection7);
                ad.Fill(dsBas);
                ;
            }
        }
        catch (Exception ex)
        {
        }

        if (myConnection1.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection2.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection3.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection4.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection5.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection6.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection7.State == System.Data.ConnectionState.Open)
            myConnection.Close();
        if (myConnection.State == System.Data.ConnectionState.Open)
            myConnection.Close();

        myConnection.Dispose(); myConnection1.Dispose(); myConnection2.Dispose(); myConnection3.Dispose();
        myConnection4.Dispose(); myConnection5.Dispose(); myConnection6.Dispose(); myConnection7.Dispose();

        myConnection = null; myConnection1 = null; myConnection2 = null; myConnection3 = null;
        myConnection4 = null; myConnection5 = null; myConnection6 = null; myConnection7 = null;

        if (dsBas.Tables.Count > 0)
            return dsBas.Tables[0];
        else
            return null;
    }
}
