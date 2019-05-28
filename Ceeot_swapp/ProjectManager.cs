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
        // apex project equivalent of the current swatt project
        private string apexBatFile;
        private string swatVersionBatFile;
        private string swatAuxVersionBatFile;
        private CEEOT_dll.General DataClass;

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

        public void createApexControlFiles()
        {
            CEEOT_dll.Control.Apexcont(0);
            // load pest file
            var pestFile = "";
            if (pestFile == "" && pestFile == " ")
                CEEOT_dll.Control.Pesticide();
            CEEOT_dll.Control.Fertilizer();
        }

        public void createApexOperationsFiles()
        {
            if (!this.validateLandUses()) {
                return;
            }

            // Create Control files
            this.cpyApex();
            var g = new System.IO.StreamWriter(System.IO.File.OpenWrite(CEEOT_dll.Initial.Output_files + @"\" + CEEOT_dll.Initial.Opcs));
            g.Close();
            g.Dispose();
            string num = "3";
            CEEOT_dll.Sitefiles.FEMFiles(ref num);
            CEEOT_dll.Sitefiles.SiteFiles(ref num);
            CEEOT_dll.Initial.CurrentOption = 24;
        }

        public void createSubAreaFiles()
        {
            if (this.validateLandUses()) {
                return;
            }

            CEEOT_dll.Initial.subareafile = 2;
            var d = new System.IO.StreamWriter(
                System.IO.File.OpenWrite(
                    CEEOT_dll.Initial.Output_files + "\\" + CEEOT_dll.Initial.suba));
            d.Close(); 
            d.Dispose();
            string num = "4";
            CEEOT_dll.Sitefiles.SiteFiles(ref num);
            CEEOT_dll.Initial.CurrentOption = 25;

            this.createApexBat();
            this.updateEnvironmentVariables();
        }

        public void createSoilFiles()
        {
            if (this.validateLandUses()) {
                return;
            }

            var d = new System.IO.StreamWriter(System.IO.File.OpenWrite(CEEOT_dll.Initial.Output_files + "\\" + CEEOT_dll.Initial.Soil));
            d.Close();
            d.Dispose();
            string num = "5";
            CEEOT_dll.Sitefiles.SiteFiles(ref num);

            CEEOT_dll.Initial.CurrentOption = 26;

            this.updateEnvironmentVariables();
        }

        public void createSiteFile()
        {
            if (this.validateLandUses())
            {
                return;
            }

            var d = new System.IO.StreamWriter(System.IO.File.OpenWrite(CEEOT_dll.Initial.Output_files + "\\" + CEEOT_dll.Initial.Site));
            d.Close();
            d.Dispose();

            string num = "6";
            CEEOT_dll.Sitefiles.SiteFiles(ref num);
            CEEOT_dll.Initial.CurrentOption = 27;

            this.updateEnvironmentVariables();
        }

        public void createWeatherFiles(int pcpages)
        {
            if (this.validateLandUses()) {
                return;
            }

            CEEOT_dll.Control.Apexcont(1);
            if (pcpages == 0)
            {
                if (CEEOT_dll.Initial.Version == "4.0.0"
                   || CEEOT_dll.Initial.Version == "4.1.0"
                   || CEEOT_dll.Initial.Version == "4.2.0"
                   || CEEOT_dll.Initial.Version == "4.3.0"
                   || CEEOT_dll.Initial.Version == "1.1.0"
                   || CEEOT_dll.Initial.Version == "1.2.0"
                   || CEEOT_dll.Initial.Version == "1.3.0")
                {
                    this.DataClass.Weather1();
                }
                else
                {
                    string num = "7";
                    CEEOT_dll.Sitefiles.SiteFiles(ref num);
                }
            }
        }

        public void createWmpFiles()
        {
            if (this.validateLandUses()) {
                return;
            }
            var d = new System.IO.StreamWriter(System.IO.File.OpenWrite(CEEOT_dll.Initial.Output_files + "\\" + CEEOT_dll.Initial.wpm1));
            d.Close();
            d.Dispose();

            string num = "8";
            CEEOT_dll.Sitefiles.SiteFiles(ref num);

            this.cpyApexSwat();

            //this.cpySwat();

            CEEOT_dll.Initial.CurrentOption = 30;
            //*****************************
            this.updateEnvironmentVariables();
        }

        public void cpyApexSwat()
        {
            var SWATF = new System.Data.DataTable();

            SWATF = CEEOT_dll.AccessDB.getDBDataTable("SELECT * FROM SwatApexF WHERE Version='" + CEEOT_dll.Initial.Version + "'");

            for (int i = 0; i < SWATF.Rows.Count - 1; i++) {
                var name1 = CEEOT_dll.Initial.OrgDir + "\\" + SWATF.Rows[i]["File"];
                var fileo = CEEOT_dll.Initial.OrgDir + "\\" + SWATF.Rows[i]["File"];

                string filet, filet1;

                if (((String)SWATF.Rows[i]["file"]).ToCharArray()[0] == '*' ) { 
                    filet = CEEOT_dll.Initial.Output_files + "\\";
                    filet1 = CEEOT_dll.Initial.Swat_Output + "\\";
                } else {
                    filet = CEEOT_dll.Initial.Output_files + "\\" + SWATF.Rows[i]["File"];
                    filet1 = CEEOT_dll.Initial.Swat_Output + "\\" + SWATF.Rows[i]["File"];
                }

                System.IO.File.Copy(fileo, filet, true);
                System.IO.File.Copy(fileo, filet1, true);
            }
            SWATF.Dispose();
        }

        public void createApexBat()
        {
            var swFile = new System.IO.StreamWriter(System.IO.File.Create(CEEOT_dll.Initial.Output_files + "\\APEXBat.txt"));
            swFile.Write("del *.swt");

            var ADORecordset = CEEOT_dll.AccessDB.getDBDataTable("SELECT SubBasin FROM Sub_Included");

            for (int i = 0; i < ADORecordset.Rows.Count; i++)
            {
                if (ADORecordset.Rows[i]["SubBasin"] == null 
                    && (string)ADORecordset.Rows[i]["SubBasin"] == "")
                {
                    var tempo = (string)ADORecordset.Rows[i]["Subbasin"];
                    var temp = tempo.Substring(2, 8);
                    var apexrunx = new System.IO.StreamWriter(System.IO.File.Create(CEEOT_dll.Initial.Output_files + "\\" + temp + ".dat"));
                    var apexrunall = new System.IO.StreamReader(System.IO.File.OpenRead(CEEOT_dll.Initial.Output_files + "\\APEXRUN.dat"));

                    while (!apexrunall.EndOfStream)
                    {
                        var rec = apexrunall.ReadLine();
                        if (rec.Substring(0,8) == temp)
                        {
                            apexrunx.WriteLine(rec);
                            swFile.WriteLine("");
                            swFile.WriteLine("del apexrun.dat");
                            swFile.WriteLine("copy " + temp + ".dat apexrun.dat");
                            switch (CEEOT_dll.Initial.Version) {
                                case "4.1.0":
                                case "4.0.0":
                                case "4.2.0": 
                                case "4.3.0":
                                    swFile.WriteLine("apex0604.exe");
                                break;
                                case "1.1.0":
                                case "1.2.0": 
                                case "1.3.0":
                                    swFile.WriteLine("apex0806.exe");
                                break;
                                default:
                                    swFile.WriteLine("apex2110.exe");
                                    break;
                            }
                        }
                    }

                    apexrunx.Close();
                    apexrunx.Dispose();
                    apexrunall.Close();
                    apexrunall.Dispose();
                }
            }

            swFile.Close();
            swFile.Dispose();
            System.IO.File.Copy(
                CEEOT_dll.Initial.Output_files +
                "\\APEXRUN.dat", CEEOT_dll.Initial.Output_files +
                "\\APEXRunAll.dat", 
                true
            );
        }

        // copy apex files
        public void cpyApex()
        {
            var rs = CEEOT_dll.AccessDB.getDBDataTable("SELECT * FROM Input_Files WHERE Version=" + "'" + CEEOT_dll.Initial.Version + "'");

            for (int i = 0; i < rs.Rows.Count ; i++)
            {
                var filet = "";
                var fileo = CEEOT_dll.Initial.OrgDir + "\\" + rs.Rows[i]["File"];
                if (((String)rs.Rows[i]["file"]).ToCharArray()[0] == '*') {
                    filet = CEEOT_dll.Initial.Output_files + "\\";
                } else {
                    filet = CEEOT_dll.Initial.Output_files + "\\" + rs.Rows[i]["File"];
                }
                if ((String)rs.Rows[i]["File"] == "APEX2110_2000.EXE") {
                    fileo = CEEOT_dll.Initial.OrgDir + "\\" + rs.Rows[i]["File"];
                    filet = CEEOT_dll.Initial.Output_files + "\\APEX2110.EXE";
                }
                //System.IO.File.Copy(fileo, filet, true);
                using (var inputFile = new System.IO.FileStream(
                        fileo,
                        System.IO.FileMode.Open,
                        System.IO.FileAccess.Read,
                        System.IO.FileShare.ReadWrite
                    ))
                {
                    using (var outputFile = new System.IO.FileStream(filet, System.IO.FileMode.Create))
                    {
                        var buffer = new byte[0x10000];
                        int bytes; 
                        while ((bytes = inputFile.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            outputFile.Write(buffer, 0, bytes);
                        }
                    }
                }
            }

            rs = CEEOT_dll.AccessDB.getDBDataTable("SELECT * FROM apexfile_dat WHERE Version='" + CEEOT_dll.Initial.Version + "' ORDER BY apexfile_dat.Order");

            System.IO.StreamWriter z = null;
            if (CEEOT_dll.Initial.Version == "3.0.0" || CEEOT_dll.Initial.Version == "3.1.0") {
                z = new System.IO.StreamWriter(System.IO.File.Create(CEEOT_dll.Initial.Output_files + "\\EPICfile.DAT"));
            } else {
                z = new System.IO.StreamWriter(System.IO.File.Create(CEEOT_dll.Initial.Output_files + "\\Apexfile.DAT"));
            }
            for (int i = 0; i < rs.Rows.Count; i++) {
                z.Write(" ");
                var tmp = String.Format("{0, -5}", ((String)rs.Rows[i]["FileCode"]).ToUpper());
                z.Write(tmp);
                z.Write("    ");
                z.WriteLine(rs.Rows[i]["FileName"]);
            }
            z.Close();
        }

        public void updateEnvironmentVariables()
        {
            string query = "SELECT CurrentOption FROM Paths";
            var proj1 = CEEOT_dll.AccessDB.getDBDataSet(ref query);

            if ((byte)(proj1.Tables[0].Rows[0]["CurrentOption"]) <= CEEOT_dll.Initial.CurrentOption)
            {
                query = "UPDATE paths SET CurrentOption=" + CEEOT_dll.Initial.CurrentOption;
                CEEOT_dll.AccessDB.getDBDataSet(ref query);
            }
            proj1.Dispose();
        }

        public void readFigFile()
        {
            String fileName = this.CurrentProject.SwattLocation + @"\fig.fig";
            String line;
            //- Read fig file for all sub basins
            System.IO.StreamReader file = new System.IO.StreamReader(fileName);
            while ((line = file.ReadLine()) != null)
            {
                if (line.Contains("subbasin"))
                {
                    string basinName = file. ReadLine().Trim();
                    Console.WriteLine("basin name " + basinName);
                    dbManager.fillBasins(this.CurrentProject, basinName);
                }
            }
            file.Close();
        }

        public void createProject(String name, String scenario, String location,
            String swattLocation, SwattProject.ProjectVersion apexVersion, SwattProject.ProjectVersion swattVersion)
        {
            // create project and add it to the store.
            this.Current = name;
            var project = new SwattProject();
            project.Name = this.Current;
            project.Location = location + @"\" + project.Name ;
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
                this.writeProject(project.Location);
                CEEOTDLLManager.initializeGlobals(this);
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
                file.Close();
            }
            basin.Hrus = hrus;
        }

        public bool validateLandUses()
        {
            string swBat = "";
            var exclude = new System.Data.DataTable();
            exclude = CEEOT_dll.AccessDB.getDBDataTable("SELECT * FROM exclude WHERE folder='"
                + this.CurrentProject.Location + "' AND project = '" + this.CurrentProject.Name + "'");
            // select version
            this.selectVersion();
            return (exclude.Rows.Count <= 0);
        }

        public void selectVersion()
        {
            switch (CEEOT_dll.Initial.Version) {
                case "1.0.0":    //versions 1.x.x were changed from APEX0806 to APEX0806 (Last one). 1/14/2015
                    apexBatFile = "Apex0806.bat";
                    swatVersionBatFile = "Sw0604_2000.bat";
                    swatAuxVersionBatFile = "Swat2000.bat";
                    break;
                case "1.1.0":
                    apexBatFile = "Apex0806.bat";
                    swatVersionBatFile = "Sw0604_2003.bat";
                    swatAuxVersionBatFile = "Swat2003.bat";
                    break;
                case "1.2.0":
                    apexBatFile = "Apex0806.bat";
                    swatVersionBatFile = "Sw0604_2003.bat";
                    swatAuxVersionBatFile = "Swat2009.bat";
                    break;
                //<New version of SWAT_2012 is added
                case "1.3.0":
                    apexBatFile = "Apex0806.bat";
                    swatVersionBatFile = "Sw0604_2003.bat";
                    swatAuxVersionBatFile = "Swat2012.bat";
                    //4/16/2013>
                    break;
                case "2.0.0":
                    apexBatFile = "Apex2110.bat";
                    swatVersionBatFile = "Sw2110_2000.bat";
                    swatAuxVersionBatFile = "Swat2000.bat";
                    break;
                case "2.1.0":
                    apexBatFile = "Apex2110.bat";
                    swatVersionBatFile = "Sw2110_2003.bat";
                    swatAuxVersionBatFile = "Swat2003.bat";
                    break;
                case "3.0.0":
                    apexBatFile = "Epic3060.bat";
                        break;
                case "3.1.0":
                    apexBatFile = "Epic3060.bat";
                    break;
                case "4.0.0":
                    apexBatFile = "Apex0604.bat";
                    swatVersionBatFile = "Sw0604_2000.bat";
                    swatAuxVersionBatFile = "Swat2000.bat";
                    break;
                case "4.1.0":
                    apexBatFile = "Apex0604.bat";
                    swatVersionBatFile = "Sw0604_2003.bat";
                    swatAuxVersionBatFile = "Swat2003.bat";
                    break;
                case "4.2.0":
                    apexBatFile = "Apex0604.bat";
                    swatVersionBatFile = "Sw0604_2003.bat";
                    swatAuxVersionBatFile = "Swat2009.bat";
                    break;
                //<New version of SWAT_2012 is added
                case "2.3.0":
                    apexBatFile = "Apex2110.bat";
                    swatVersionBatFile = "Sw2110_2003.bat";
                    swatAuxVersionBatFile = "Swat2012.bat";
                    break;
                case "4.3.0":
                    apexBatFile = "Apex0604.bat";
                    swatVersionBatFile = "Sw0604_2003.bat";
                    swatAuxVersionBatFile = "Swat2012.bat";
                    break;
                }
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
                        int lastSlashIdx = filename.LastIndexOf(@"\");
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

            System.IO.StreamReader file = new System.IO.StreamReader(path + @"\project.xml");
            SwattProject proj = (SwattProject)reader.Deserialize(file);

            this.Current = proj.Name;
            // TODO: Add database connection 
            this.projectMapping.Add(this.Current, proj);
            this.loadSubBasins();
            CEEOTDLLManager.initializeGlobals(this);

            file.Close();
        }

        public void writeProject(String path)
        {
            System.Xml.Serialization.XmlSerializer writer =
            new System.Xml.Serialization.XmlSerializer(typeof(SwattProject));

            System.IO.FileStream file = System.IO.File.Create(path + @"\project.xml");

            writer.Serialize(file, this.CurrentProject);
            file.Close();
        }

        public SwattProject CurrentProject
        {
            get {
                return (this.Current != null) 
                        ? (SwattProject)projectMapping[this.Current] 
                        : null;
            }
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
