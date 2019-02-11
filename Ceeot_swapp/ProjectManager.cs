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
        }

        // store for created projects
        private Hashtable projectMapping;
        // the name of the current projects
        private String currentProjectName;
        public void createProject(String name, String location, Version apexVersion, Version swattVersion) 
        {
            // create project and add it to the store.
            this.Current = name;
            var project = new Project();
            project.name = this.Current;
            project.location = location;
            project.apexVersion = apexVersion;
            project.swattVersion = swattVersion;
            projectMapping.Add(this.Current, project);
        }

        public String Current
        {
            get { return this.currentProjectName;  }
            set { if (value == "") {
                    throw new ProjectException("Project name cannot be empty");
                }
            }
        }
    }
}
