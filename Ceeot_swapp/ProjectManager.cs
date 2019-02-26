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
        public enum Version
        {
            APEX_0604, APEX_0806, SWATT_2005, SWATT_2009, SWATT_2012,
        }

        public class Project {
            public String name;
            public String location;
            public String swattLocation;
            public Version apexVersion;
            public Version swattVersion;
            public string initialProject;
            public string initialScenario;

            public struct SubBasin
            {
                public string name;
                private bool selected;

                public Boolean Selected {
                    get { return this.selected; }
                    set { this.selected = value; }
                }
                public String Name {
                    get { return this.name; }
                    set { this.name = value; }
                }
            }

            private List<ProjectManager.Project.SubBasin> subBasins;
            private DatabaseManager dbManager;

            public Project() {
                subBasins = new List<SubBasin>();
            }

            public List<string> SelectedSubBasins {
                get {
                    List<string> basins = new List<string>();
                    foreach (SubBasin s in subBasins) {
                        if (s.Selected) basins.Add(s.Name);
                    }
                    return basins;
                }
            }

            public List<string> AllSubBasins {
                get {
                    List<string> basins = new List<string>();
                    foreach (SubBasin s in subBasins)
                    {
                        basins.Add(s.Name);
                    }
                    return basins;
                }
            }
        }

        // store for created projects
        private Hashtable projectMapping ;
        // the name of the current projects
        private String currentProjectName;
        public ProjectManager()
        {
            this.projectMapping = new Hashtable();
            this.dbManager = new DatabaseManager();
            projectMapping.Add("New Tab", null);
        }

        public void createProject(String name, String location, String swattLocation, Version apexVersion, Version swattVersion) 
        {
            // create project and add it to the store.
            this.Current = name;
            var project = new Project();
            project.name = this.Current;
            project.location = location;
            project.swattLocation = swattLocation;
            project.apexVersion = apexVersion;
            project.swattVersion = swattVersion;

            // TODO: Add database connection 
            projectMapping.Add(this.Current, project);
        }

        public void loadSubBasins()
        {
            try
            {
                var filenames = System.IO.Directory.GetFiles(CurrentProject.swattLocation);
                foreach (var filename in filenames) {
                    int extensionStartIdx = filename.IndexOf(".sub");
                    int lastSlashIdx = filename.LastIndexOf("\\");
                    if (extensionStartIdx >= 0) {
                        var basinName = filename.Substring(lastSlashIdx+1, 9);
                        CurrentProject.subBasins.Add(new Project.SubBasin { name = basinName });
                    }
                }
            } catch (Exception ex) {}
        }

        public Project CurrentProject
        {
            get { return (Project)projectMapping[this.Current]; } 
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
