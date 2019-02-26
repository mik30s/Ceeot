using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ceeot_swapp
{
    class DatabaseManager
    {
        DatabaseManager()
        {
            // Try opening a connection to the access database for CEEOT.
            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
            // TODO: Modify the connection string and include any
            // additional required properties for your database.
            conn.ConnectionString = 
                @"Provider=Microsoft.Jet.OLEDB.4.0;" +
                @"Data source= C:\Documents and Settings\tiaer\" +
                @"My Documents\AccessFile.mdb";
            try
            {
                conn.Open();
                // Insert code to process data.
            }
            catch (Exception ex)
            {
               throw ex;
            }
            finally
            {
                conn.Close();
            }
        }
    }
}
