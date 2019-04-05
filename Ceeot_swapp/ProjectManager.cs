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

        public void readFigFile()
        {
            String fileName = this.CurrentProject.SwattLocation + "\\fig.fig";
            String line;
            //- Read fig file for all sub basins
            System.IO.StreamReader file = new System.IO.StreamReader(fileName);
            while ((line = file.ReadLine()) != null)
            {
                if (line.Contains("subbasin"))
                {
                    string basinName = file.ReadLine().Trim();
                    Console.WriteLine("basin name " + basinName);
                    dbManager.fillBasins(this.CurrentProject, basinName);
                }
            }
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
                //this.readFigFile();
                this.writeProject(project.Location + "//" + project.Name );
            }
            catch (Exception ex) {
                throw new ProjectException("Couldn't update project path in database" + ex.Message);
            }
        }

        public void loadHRUs(SwattProject.SubBasin basin)
        {
            var b = basin.name.Substring(0,5);
            List<SwattProject.HRU> hrus = new List<SwattProject.HRU>();
            var filenames = System.IO.Directory.GetFiles(CurrentProject.SwattLocation,  b + "*.hru");
            foreach (var fileName in filenames)
            {
                System.IO.StreamReader file = new System.IO.StreamReader(fileName);
                SwattProject.HRU hru = new SwattProject.HRU();
                CropCodes.Code code;
                // Extract crop code from first line in file
                var line = file.ReadLine();
                line = line.Split()[5];
                line = line.Split(':')[1];
                Enum.TryParse(line, out code);
                // insert code
                hru.Code = code;
                // fill description
                hru.Description = CropCodes.getDescription(hru.Code);
                hru.SubBasin = basin.Name;
                // add hru to sub basin
                hrus.Add(hru);
            }
            basin.Hrus = hrus;
        }

        public void loadSubBasins()
        {
            try
            {
                var filenames = System.IO.Directory.GetFiles(CurrentProject.SwattLocation);
                List<SwattProject.SubBasin> basins = new List<SwattProject.SubBasin>();
                foreach (var filename in filenames) {
                    if (!filename.Contains("output")) {
                        int extensionStartIdx = filename.IndexOf(".sub");
                        int lastSlashIdx = filename.LastIndexOf("\\");
                        if (extensionStartIdx >= 0) {
                            var basinName = filename.Substring(lastSlashIdx + 1, 9);
                            var basin = new SwattProject.SubBasin { name = basinName };
                            this.loadHRUs(basin);
                            basins.Add(basin);
                        }
                    }
                }
                this.CurrentProject.SubBasins = basins;
            } catch (Exception ex) { throw new ProjectException("Failed to load sub basins. " + ex.Message); }
        }

        public void readProject(String path)
        {
            System.Xml.Serialization.XmlSerializer reader =
                new System.Xml.Serialization.XmlSerializer(typeof(SwattProject));

            System.IO.StreamReader file = new System.IO.StreamReader(path + "//project.xml");
            SwattProject proj = (SwattProject)reader.Deserialize(file);

            this.Current = proj.Name;
            // TODO: Add database connection 
            this.projectMapping.Add(this.Current, proj);
            this.loadSubBasins();

            file.Close();
        }

        public void writeProject(String path)
        {
            System.Xml.Serialization.XmlSerializer writer =
            new System.Xml.Serialization.XmlSerializer(typeof(SwattProject));

            System.IO.FileStream file = System.IO.File.Create(path + "//project.xml");

            writer.Serialize(file, this.CurrentProject);
            file.Close();
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
