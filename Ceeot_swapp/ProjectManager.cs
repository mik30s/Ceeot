using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace Ceeot_swapp
{
    class ProjectException : Exception
    {
        public ProjectException(String name) : base(name) { }
    }

    public class ProjectManager
    {
        // store for created projects
        private Hashtable projectMapping ;
        // the name of the current projects
        private String currentProjectName;
        // manages connection to an apex ms access database
        private DatabaseManager dbManager;

        public ProjectManager()
        {
            this.projectMapping = new Hashtable();
            try
            {
                this.dbManager = new DatabaseManager();
            } catch(Exception ex)
            {
                throw new ProjectException("Failed to establish connection to database. " + ex.Message);
            }
            projectMapping.Add("New Tab", null);
        }

        public void createProject(String name, String scenario, String location,
            String swattLocation, SwattProject.ProjectVersion apexVersion, SwattProject.ProjectVersion swattVersion)
        {
            // create project and add it to the store.
            this.Current = name;
            var project = new SwattProject();
            project.Name = this.Current;
            project.Location = location;
            project.CurrentScenario = scenario;
            project.SwattLocation = swattLocation;
            project.ApexVersion = apexVersion;
            project.SwattVersion = swattVersion;

            // TODO: Add database connection 
            projectMapping.Add(this.Current, project);

            try
            {
                // insert project into project path table
                if (!dbManager.setProjectPath(project)) {
                    throw new ProjectException("Couldn't update project path in database");
                }
            } catch(Exception ex) {
                throw new ProjectException("Couldn't update project path in database" + ex.Message);
            }
        }

        public void loadSubBasins()
        {
            try
            {
                var filenames = System.IO.Directory.GetFiles(CurrentProject.SwattLocation);
                foreach (var filename in filenames) {
                    int extensionStartIdx = filename.IndexOf(".sub");
                    int lastSlashIdx = filename.LastIndexOf("\\");
                    if (extensionStartIdx >= 0) {
                        var basinName = filename.Substring(lastSlashIdx+1, 9);
                        CurrentProject.SubBasins.Add(new SwattProject.SubBasin { name = basinName });
                    }
                }
            } catch (Exception ex) { throw new ProjectException("Failed to load sub basins " + ex.Message); }
        }

        public SwattProject CurrentProject
        {
            get { return (SwattProject)projectMapping[this.Current]; } 
        }

        public String Current
        {
            get { return this.currentProjectName;  }
            set { if (value == "") {
                    throw new ProjectException("Project name cannot be empty");
                }
                this.currentProjectName = value;
            }
        }
    }
}
