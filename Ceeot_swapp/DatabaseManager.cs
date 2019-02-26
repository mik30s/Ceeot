using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace Ceeot_swapp
{
    public class DatabaseManager
    {
        private System.Data.OleDb.OleDbConnection conn;
        // insert a project path
        private const string SET_PROJECT_PATH_QUERY = 
            @"INSERT INTO Paths (ProjectName, Scenario, Folder, Version1, APEX)
              VALUES ({0}, {1}, {2}, {3}, {4})";
        private OleDbCommand queryCommand;
        
        public DatabaseManager()
        {
            // Try opening a connection to the access database for CEEOT.
             this.conn = new OleDbConnection();
            // TODO: Modify the connection string and include any
            // additional required properties for your database.
            // WARNING: This only works for access databases older than 2007 per:
            // https://stackoverflow.com/a/17023942
            this.conn.ConnectionString = 
                @"Provider=Microsoft.Jet.OLEDB.4.0;" +
                @"Data source= resources\databases\Project_Parameters.mdb";
            try
            {
                this.conn.Open();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
            finally
            {
                this.conn.Close();
            }
        }

        ~DatabaseManager() { this.conn.Close(); }

        public Boolean setProjectPath(Project project)
        {
            String versionString = project.ApexVersion.ToString().Replace("_", "") 
                + " & " + project.SwattVersion.ToString().Replace("_", "");

            String.Format(
                SET_PROJECT_PATH_QUERY, project.Name, project.CurrentScenario, 
                project.Location, "4.0", versionString, "\resources\apex"
            );

            this.queryCommand = new OleDbCommand(SET_PROJECT_PATH_QUERY, this.conn);
            return queryCommand.ExecuteNonQuery() > 0;
        }
    }
}
