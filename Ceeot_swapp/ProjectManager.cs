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

        public struct Project {
            public String name;
            public String location;
            public Version apexVersion;
            public Version swattVersion;
            public string initialProject;
            public string initialScenario;

            public struct SubBasin
            {
                public string name;

                public String Name {
                    get { return this.name; }
                    set { this.name = value; }
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
            projectMapping.Add("New Tab", null);
        }

        public void createProject(String name, String location, Version apexVersion, Version swattVersion) 
        {
            // create project and add it to the store.
            this.Current = name;
            var project = new Project();
            project.name = this.Current;
            project.location = location;
            project.apexVersion = apexVersion;
            project.swattVersion = swattVersion;

            // TODO: Add database connection 
            projectMapping.Add(this.Current, project);
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

        public List<String> Projects
        {
            get {
                var list = new List<String>();
                foreach (string k in projectMapping.Keys)
                {
                    list.Add(k);
                }
                return list;
            }
        }

        public List<Project.SubBasin> getSubBasins()
        {
            return null;
        }
    }
}
