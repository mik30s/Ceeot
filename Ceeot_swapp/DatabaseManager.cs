using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Text.RegularExpressions;

namespace Ceeot_swapp
{
    public class DatabaseManager
    {
        private System.Data.OleDb.OleDbConnection conn;
        // query to insert a project path
        private const string SET_PROJECT_PATH_QUERY =
            @"INSERT INTO Projects ([project_name],location,versions,[swatt_files_location]) VALUES (?, ? ,? , ?);";//'{0}', '{1}', '{2}', '{3}' );";
        private const string INSERT_SCENARIO_FOR_PROJECT =
            @"INSERT INTO Scenarios (project_name,scenario_name) VALUES (?,?);";

        private OleDbCommand queryCommand;
        
        public DatabaseManager()
        {
            // Try opening a connection to the access database for CEEOT.
             this.conn = new OleDbConnection();
            // TODO: Modify the connection string and include any
            // additional required properties for your database.
            // WARNING: This only works for access databases older than 2007 as per:
            // https://stackoverflow.com/a/17023942
            this.conn.ConnectionString =
                //@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\resources\databases\Project_Parameters.mdb;Persist Security Info=True";
                @"Provider=Microsoft.Jet.OLEDB.4.0;" +
                @"Data Source=resources/databases/Project_Parameters.mdb;";
            try
            {
                this.conn.Open();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
        }

       // ~DatabaseManager() { this.conn.Close(); }

        public Boolean setProjectPath(Project project)
        {
            if (this.conn.State == System.Data.ConnectionState.Open)
            {
                // Build apex, swatt and fem version strings
                String versionString = project.ApexVersion.ToString().Replace("_", "")
                    + " & " + project.SwattVersion.ToString().Replace("_", "");

                // build query string from project variables
                /*
                string queryString = Regex.Replace(String.Format(
                    SET_PROJECT_PATH_QUERY, 
                    project.Name,   // Project Name
                    project.Location, // Project location on drive
                    versionString, // Project apex & swatt & fem versions used
                    project.SwattLocation // location of swatt files for Project
                ), @"\t|\n|\r", ""); ;
                */

                // build ms access insert new scenario command
                this.queryCommand =
                    new OleDbCommand(INSERT_SCENARIO_FOR_PROJECT, this.conn);
                this.queryCommand.Parameters.Add("@p1", OleDbType.VarWChar).Value = project.Name;
                this.queryCommand.Parameters.Add("@p2", OleDbType.VarWChar).Value = project.CurrentScenario;

                var success = queryCommand.ExecuteNonQuery() > 0;

                // build ms access query command
                this.queryCommand = 
                    new OleDbCommand(SET_PROJECT_PATH_QUERY, this.conn);

                this.queryCommand.Parameters.Add("@p1", OleDbType.VarChar).Value = project.Name;
                this.queryCommand.Parameters.Add("@p2", OleDbType.VarWChar).Value = project.Location;
                this.queryCommand.Parameters.Add("@p3", OleDbType.VarWChar).Value = versionString;
                this.queryCommand.Parameters.Add("@p4", OleDbType.VarWChar).Value = project.SwattLocation;

                // execute query and return
                success = queryCommand.ExecuteNonQuery() > 0;

                return success;
            }
            return false;
        }
    }
}
