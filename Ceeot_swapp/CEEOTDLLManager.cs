using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;

namespace Ceeot_swapp
{
    class CEEOTDLLManager
    {
        public static void initializeGlobals(ProjectManager projectManager)
        {
            CEEOT_dll.Initial.Scenario = projectManager.CurrentProject.CurrentScenario;
            CEEOT_dll.Initial.CurrentOption = 10;
            CEEOT_dll.Initial.Output_files = projectManager.CurrentProject.Location + @"\APEX";
            CEEOT_dll.Initial.New_Swat = projectManager.CurrentProject.Location + @"\New_SWAT";
            CEEOT_dll.Initial.FEM = projectManager.CurrentProject.Location + @"\FEM";
            CEEOT_dll.Initial.Swat_Output = projectManager.CurrentProject.Location + @"\SWAT_Output";
            CEEOT_dll.Initial.Input_Files = projectManager.CurrentProject.SwattLocation; 
            CEEOT_dll.Initial.Pest = "Pest.dat";
            CEEOT_dll.Initial.Dir1 = (System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) 
                                    + @"/resources/databases").Substring(6);
            //CEEOT_dll.Initial.Dir2 = projectManager.CurrentProject.Location;
            CEEOT_dll.Initial.OrgDir = CEEOT_dll.Initial.Output_files;

            string selectedVersion = "";
            if (projectManager.CurrentProject.ApexVersion == SwattProject.ProjectVersion.APEX_0806) selectedVersion += "1";
            if (projectManager.CurrentProject.ApexVersion == SwattProject.ProjectVersion.APEX_0604) selectedVersion += "3";
            if (projectManager.CurrentProject.SwattVersion == SwattProject.ProjectVersion.SWATT_2005) selectedVersion += "1";
            if (projectManager.CurrentProject.SwattVersion == SwattProject.ProjectVersion.SWATT_2009) selectedVersion += "2";
            if (projectManager.CurrentProject.SwattVersion == SwattProject.ProjectVersion.SWATT_2012) selectedVersion += "3";

            switch (selectedVersion) {
                case "10":
                    CEEOT_dll.Initial.Year_Col = 3;  //year size
                   CEEOT_dll.Initial.Version = "1.0.0";
                   CEEOT_dll.Initial.Version2 = "APEX0806 + SWAT2000";
                   CEEOT_dll.Initial.Pesticide = "Pest1310.dat";
                   CEEOT_dll.Initial.Espace = "  ";
                   CEEOT_dll.Initial.Espace1 = "   ";
                   CEEOT_dll.Initial.col1 = 7;
                   CEEOT_dll.Initial.col3 = 7;
                   CEEOT_dll.Initial.col4 = 2;
                   CEEOT_dll.Initial.col2 = 12;  //operation code in operatin file 
                    CEEOT_dll.Initial.col5 = 26;   //fertilizer code in operatino file
                    CEEOT_dll.Initial.col6 = 8; //*.dat File Name position
                    CEEOT_dll.Initial.col7 = 21; //Crop code in operation file
                    CEEOT_dll.Initial.col8 = 38; //Curve number in operation file
                   CEEOT_dll.Initial.limit = 14;
                   CEEOT_dll.Initial.SubaLines = 11; //# of lines for each field within subarea files.
                   CEEOT_dll.Initial.cntrl2 = "\\file.cio";
                    CEEOT_dll.Initial.cntrl = "\\file1.cio";
                    CEEOT_dll.Initial.cntrl3 = "\\file.cio";
                    CEEOT_dll.Initial.cntrl4 = "\\file2.cio";
                    CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                    CEEOT_dll.Initial.suba = "suba1310.dat";
                    CEEOT_dll.Initial.Soil = "soil1310.dat";
                   CEEOT_dll.Initial.Site = "site1310.dat";
                    CEEOT_dll.Initial.Opcs = "opcs1310.dat";
                   CEEOT_dll.Initial.wpm1 = "wpm11310.dat";
                   CEEOT_dll.Initial.parm = "parm1310.dat";
                   CEEOT_dll.Initial.Till = "Till1310.dat";
                   CEEOT_dll.Initial.Fertilizer = "Fert1310.dat";
                   CEEOT_dll.Initial.sublines = "11";
                   CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "11":
                    CEEOT_dll.Initial.Year_Col = 3;  //year size
                   CEEOT_dll.Initial.Version = "1.1.0";
                   CEEOT_dll.Initial.Version2 = "APEX0806 + SWAT2005";
                   CEEOT_dll.Initial.Pesticide = "Pest0806.dat";
                   CEEOT_dll.Initial.Espace = "  ";
                   CEEOT_dll.Initial.Espace1 = "   ";
                   CEEOT_dll.Initial.col1 = 11;
                   CEEOT_dll.Initial.col2 = 12;  //operation code in operatin file 
                   CEEOT_dll.Initial.col3 = 7;
                   CEEOT_dll.Initial.col4 = 2;
                   CEEOT_dll.Initial.col5 = 26;   //fertilizer code in operatino file
                   CEEOT_dll.Initial.col6 = 7;
                   CEEOT_dll.Initial.col7 = 21; //Crop code in operation file
                   CEEOT_dll.Initial.col8 = 38; //Curve number in operation file
                   CEEOT_dll.Initial.limit = 1;
                   CEEOT_dll.Initial.SubaLines = 12;
                    CEEOT_dll.Initial.cntrl2 = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl = "\\" + CEEOT_dll.Initial.figsFile;
                   CEEOT_dll.Initial.cntrl3 = "\\" + CEEOT_dll.Initial.figsFile;
                   CEEOT_dll.Initial.cntrl4 = "\\file2.cio";
                   CEEOT_dll.Initial.cntrl = "\\file1.cio";
                   CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                   CEEOT_dll.Initial.suba = "suba0806.dat";
                   CEEOT_dll.Initial.Soil = "soil0806.dat";
                   CEEOT_dll.Initial.Site = "site0806.dat";
                   CEEOT_dll.Initial.Opcs = "opcs0806.dat";
                   CEEOT_dll.Initial.wpm1 = "wpm10806.dat";
                   CEEOT_dll.Initial.parm = "parm0806.dat";
                   CEEOT_dll.Initial.Till = "Till0806.dat";
                   CEEOT_dll.Initial.Fertilizer = "Fert0806.dat";
                   CEEOT_dll.Initial.herd = "Herd0806.dat";
                   CEEOT_dll.Initial.sublines = ""+12;
                   CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "12":
                    CEEOT_dll.Initial.Year_Col = 3;  //year size
                   CEEOT_dll.Initial.Version = "1.2.0";
                   CEEOT_dll.Initial.Version2 = "APEX0806 + SWAT2009";
                   CEEOT_dll.Initial.Pesticide = "Pest0806.dat";
                   CEEOT_dll.Initial.Espace = " ";
                   CEEOT_dll.Initial.Espace1 = " ";
                   CEEOT_dll.Initial.col1 = 11;
                   CEEOT_dll.Initial.col2 = 12; //Operatio code in operation file
                   CEEOT_dll.Initial.col3 = 7;
                   CEEOT_dll.Initial.col4 = 2;
                   CEEOT_dll.Initial.col5 = 26;   //fertilizer code in operatino file
                   CEEOT_dll.Initial.col6 = 7;
                   CEEOT_dll.Initial.col7 = 21; //Crop code in operation file
                   CEEOT_dll.Initial.col8 = 38; //Curve number in operation file;
                   CEEOT_dll.Initial.limit = 1;
                   CEEOT_dll.Initial.SubaLines = 12;
                   CEEOT_dll.Initial.cntrl2 = "\" +CEEOT_dll.Initial.figsFile";
                   CEEOT_dll.Initial.cntrl = "\" +CEEOT_dll.Initial.figsFile";
                   CEEOT_dll.Initial.cntrl3 = "\" +CEEOT_dll.Initial.figsFile" ;
                   CEEOT_dll.Initial.cntrl4 = "\file2.cio";
                   CEEOT_dll.Initial.cntrl = "\file1.cio";
                   CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                   CEEOT_dll.Initial.suba = "suba0806.dat";
                   CEEOT_dll.Initial.Soil = "soil0806.dat";
                   CEEOT_dll.Initial.Site = "site0806.dat";
                   CEEOT_dll.Initial.Opcs = "opcs0806.dat";
                   CEEOT_dll.Initial.wpm1 = "wpm10806.dat";
                   CEEOT_dll.Initial.parm = "parm0806.dat";
                   CEEOT_dll.Initial.Till = "Till0806.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert0806.dat";
                   CEEOT_dll.Initial.herd = "Herd0806.dat";
                   CEEOT_dll.Initial.sublines = ""+12;
                   CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "13":
                    CEEOT_dll.Initial.Year_Col = 3;  //year size
                    CEEOT_dll.Initial.Version = "1.3.0";
                    CEEOT_dll.Initial.Version2 = "APEX0806 + SWAT2012";
                    CEEOT_dll.Initial.Pesticide = "Pest0806.dat";
                    CEEOT_dll.Initial.Espace = " ";
                    CEEOT_dll.Initial.Espace1 = " ";
                    CEEOT_dll.Initial.col1 = 11;
                    CEEOT_dll.Initial.col2 = 12; //Operatio code in operation file
                   CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col5 = 26;  //fertilizer code in operatino file
                    CEEOT_dll.Initial.col6 = 7;
                    CEEOT_dll.Initial.col7 = 21; //Crop code in operation file
                    CEEOT_dll.Initial.col8 = 38; //Curve number in operation file
                    CEEOT_dll.Initial.limit = 1;
                    CEEOT_dll.Initial.SubaLines = 12;
                    CEEOT_dll.Initial.cntrl2 = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl = "\\" + CEEOT_dll.Initial.figsFile;
                   CEEOT_dll.Initial.cntrl3 = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl4 = "\file2.cio";
                    CEEOT_dll.Initial.cntrl = "\file1.cio";
                    CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                    CEEOT_dll.Initial.suba = "suba0806.dat";
                    CEEOT_dll.Initial.Soil = "soil0806.dat";
                    CEEOT_dll.Initial.Site = "site0806.dat";
                    CEEOT_dll.Initial.Opcs = "opcs0806.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm10806.dat";
                    CEEOT_dll.Initial.parm = "parm0806.dat";
                    CEEOT_dll.Initial.Till = "Till0806.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert0806.dat";
                    CEEOT_dll.Initial.herd = "Herd0806.dat";
                    CEEOT_dll.Initial.sublines =""+ 12;
                    CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "00":
                    CEEOT_dll.Initial.Year_Col = 2;  //year size
                    CEEOT_dll.Initial.Version = "2.0.0";
                    CEEOT_dll.Initial.Version2 = "APEX2110 + SWAT2000";
                    CEEOT_dll.Initial.Pesticide = "Pest2110.dat";
                    CEEOT_dll.Initial.Espace = " ";
                    CEEOT_dll.Initial.Espace1 = " ";
                    CEEOT_dll.Initial.col1 = 7;
                    CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col2 = 9; //Operation code in operation file
                   CEEOT_dll.Initial.col5 = 23;   //fertilizer code in operatino file
                    CEEOT_dll.Initial.col6 = 7;
                    CEEOT_dll.Initial.col7 = 12; //Crop code in operation file
                    CEEOT_dll.Initial.col8 = 35; //Curve number in operation file
                    CEEOT_dll.Initial.limit = 14; ;
                    CEEOT_dll.Initial.SubaLines = 12;
                    CEEOT_dll.Initial.cntrl2 = "\\file.cio";
                    CEEOT_dll.Initial.cntrl = "\\file1.cio";
                    CEEOT_dll.Initial.cntrl3 = "\\file.cio";
                    CEEOT_dll.Initial.cntrl4 = "\\file2.cio";
                    CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                    CEEOT_dll.Initial.suba = "suba2110.dat";
                    CEEOT_dll.Initial.Soil = "soil2110.dat";
                    CEEOT_dll.Initial.Site = "site2110.dat";
                    CEEOT_dll.Initial.Opcs = "opcs2110.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm12110.dat";
                    CEEOT_dll.Initial.parm = "parm2110.dat";
                    CEEOT_dll.Initial.Till = "Till2110.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert2110.dat";
                    CEEOT_dll.Initial.sublines = ""+12;
                    CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "01":
                   CEEOT_dll.Initial.Year_Col = 2; //year size
                   CEEOT_dll.Initial.Version = "2.1.0";
                    CEEOT_dll.Initial.Version2 = "APEX2110 + SWAT2005";
                    CEEOT_dll.Initial.Pesticide = "Pest2110.dat";
                    CEEOT_dll.Initial.Espace = " ";
                    CEEOT_dll.Initial.Espace1 = " ";
                    CEEOT_dll.Initial.col1 = 11;
                    CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col2 = 9; //Operation code in operation file
                    CEEOT_dll.Initial.col5 = 23;  //fertilizer code in operatino file
                    CEEOT_dll.Initial.col6 = 7;
                    CEEOT_dll.Initial.col7 = 12; //Crop code in operation file
                    CEEOT_dll.Initial.col8 = 35;//Curve number in operation file
                    CEEOT_dll.Initial.limit = 1;
                    CEEOT_dll.Initial.SubaLines = 12;
                   CEEOT_dll.Initial.cntrl2 = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl3 = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl4 = "\file2.cio";
                    CEEOT_dll.Initial.cntrl = "\file1.cio";
                    CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                    CEEOT_dll.Initial.suba = "suba2110.dat";
                    CEEOT_dll.Initial.Soil = "soil2110.dat";
                    CEEOT_dll.Initial.Site = "site2110.dat";
                    CEEOT_dll.Initial.Opcs = "opcs2110.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm12110.dat";
                    CEEOT_dll.Initial.parm = "parm2110.dat";
                    CEEOT_dll.Initial.Till = "Till2110.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert2110.dat";
                    CEEOT_dll.Initial.sublines = ""+12;
                    CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "03":
                   CEEOT_dll.Initial.Year_Col = 2; //year size
                   CEEOT_dll.Initial.Version = "2.3.0";
                    CEEOT_dll.Initial.Version2 = "APEX2110 + SWAT2012";
                    CEEOT_dll.Initial.Pesticide = "Pest2110.dat";
                    CEEOT_dll.Initial.Espace = " ";
                    CEEOT_dll.Initial.Espace1 = " ";
                    CEEOT_dll.Initial.col1 = 11;
                    CEEOT_dll.Initial.col2 = 9; //Operation code in operation file
                    CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col5 = 23; //fertilizer code in operatino file
                    CEEOT_dll.Initial.col6 = 7;
                    CEEOT_dll.Initial.col7 = 12; //Crop code in operation file
                    CEEOT_dll.Initial.col8 = 35; //Curve number in operation file
                    CEEOT_dll.Initial.limit = 1;
                    CEEOT_dll.Initial.SubaLines = 12;
                   CEEOT_dll.Initial.cntrl2 = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl3 = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl4 = "\\file2.cio";
                    CEEOT_dll.Initial.cntrl = "\\file1.cio";
                    CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                    CEEOT_dll.Initial.suba = "suba2110.dat";
                    CEEOT_dll.Initial.Soil = "soil2110.dat";
                    CEEOT_dll.Initial.Site = "site2110.dat";
                    CEEOT_dll.Initial.Opcs = "opcs2110.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm12110.dat";
                    CEEOT_dll.Initial.parm = "parm2110.dat";
                    CEEOT_dll.Initial.Till = "Till2110.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert2110.dat";
                    CEEOT_dll.Initial.sublines = ""+12;
                    CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "20":
                    CEEOT_dll.Initial.Year_Col = 2;  //year size
                    CEEOT_dll.Initial.Version = "3.0.0";
                    CEEOT_dll.Initial.Version2 = "EPIC3060 + SWAT2000"; ;
                    CEEOT_dll.Initial.Espace = "  ";
                    CEEOT_dll.Initial.Espace1 = "   ";
                    CEEOT_dll.Initial.col1 = 7;
                    CEEOT_dll.Initial.col2 = 9; //Operation code in operation file
                    CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col5 = 23;  //fertilizer code in operatino file
                    CEEOT_dll.Initial.col6 = 8;
                    CEEOT_dll.Initial.col7 = 12; //Crop code in operation file
                    CEEOT_dll.Initial.col8 = 35; //Curve number in operation file
                    CEEOT_dll.Initial.limit = 14;
                    CEEOT_dll.Initial.cntrl2 = "\\file.cio";
                    CEEOT_dll.Initial.cntrl = "\\file1.cio";
                    CEEOT_dll.Initial.cntrl3 = "\\file.cio";
                    CEEOT_dll.Initial.cntrl4 = "\\file2.cio";
                    CEEOT_dll.Initial.RunFile = "\\EpicRun.dat";
                    CEEOT_dll.Initial.suba = "suba3060.dat";
                    CEEOT_dll.Initial.Soil = "soil3060.dat";
                    CEEOT_dll.Initial.Site = "site3060.dat";
                    CEEOT_dll.Initial.Opcs = "opcs3060.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm13060.dat";
                    CEEOT_dll.Initial.parm = "PARM3060.dat";
                    CEEOT_dll.Initial.Till = "Till2110.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert3060.dat";
                    CEEOT_dll.Initial.cont = "Epiccont.dat";
                    break;
                case "21":
                    CEEOT_dll.Initial.Year_Col = 2;  //year size
                    CEEOT_dll.Initial.Version = "3.1.0";
                    CEEOT_dll.Initial.Version2 = "EPIC3060 + SWAT2005";
                    CEEOT_dll.Initial.Espace = "  ";
                    CEEOT_dll.Initial.Espace1 = "   ";
                    CEEOT_dll.Initial.col1 = 11;
                    CEEOT_dll.Initial.col2 = 9; //Operation code in operation file
                    CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col5 = 23;  //fertilizer code in operatino file
                    CEEOT_dll.Initial.col7 = 12; //Crop code in operation file
                    CEEOT_dll.Initial.col8 = 35; //Curve number in operation file
                    CEEOT_dll.Initial.limit = 1;
                    CEEOT_dll.Initial.col6 = 8;
                    CEEOT_dll.Initial.cntrl2 = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl = "\\file1.cio";
                    CEEOT_dll.Initial.cntrl3 = "\\file.cio";
                   CEEOT_dll.Initial.cntrl4 = "\\file2.cio";
                    CEEOT_dll.Initial.RunFile = "\\EpicRun.dat";
                    CEEOT_dll.Initial.suba = "suba3060.dat";
                    CEEOT_dll.Initial.Soil = "soil3060.dat";
                    CEEOT_dll.Initial.Site = "site3060.dat";
                    CEEOT_dll.Initial.Opcs = "opcs3060.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm13060.dat";
                    CEEOT_dll.Initial.parm = "PARM3060.dat";
                    CEEOT_dll.Initial.Till = "Till2110.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert3060.dat";
                    CEEOT_dll.Initial.cont = "Epiccont.dat";
                    break;
                case "30":
                    CEEOT_dll.Initial.Year_Col = 2; //year size
                    CEEOT_dll.Initial.Version = "4.0.0";
                    CEEOT_dll.Initial.Version2 = "APEX0604 + SWAT2000";
                    CEEOT_dll.Initial.Pesticide = "Pest2110.dat";
                    CEEOT_dll.Initial.Espace = " ";
                    CEEOT_dll.Initial.Espace1 = " ";
                    CEEOT_dll.Initial.col1 = 7;
                    CEEOT_dll.Initial.col2 = 9;//Operation code in operation file
                    CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col5 = 23;   //fertilizer code in operatino file
                    CEEOT_dll.Initial.col6 = 7;
                    CEEOT_dll.Initial.col7 = 18; //Crop code in operation file
                    CEEOT_dll.Initial.col8 = 35; //Curve number in operation file
                    CEEOT_dll.Initial.limit = 14;
                    CEEOT_dll.Initial.SubaLines = 12;
                    CEEOT_dll.Initial.cntrl2 = "\file.cio";
                    CEEOT_dll.Initial.cntrl = "\file1.cio";
                    CEEOT_dll.Initial.cntrl3 = "\file.cio";
                    CEEOT_dll.Initial.cntrl4 = "\file2.cio";
                    CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                    CEEOT_dll.Initial.suba = "suba2110.dat";
                    CEEOT_dll.Initial.Soil = "soil2110.dat";
                    CEEOT_dll.Initial.Site = "site2110.dat";
                    CEEOT_dll.Initial.Opcs = "opcs2110.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm12110.dat";
                    CEEOT_dll.Initial.parm = "parm0604.dat";
                    CEEOT_dll.Initial.Till = "Till2110.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert2110.dat";
                    CEEOT_dll.Initial.herd = "Herd0604.dat";
                    CEEOT_dll.Initial.sublines = ""+12;
                    CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "31":
                    CEEOT_dll.Initial.Year_Col = 2; //year size
                    CEEOT_dll.Initial.Version = "4.1.0";
                    CEEOT_dll.Initial.Version2 = "APEX0604 + SWAT2005";
                    CEEOT_dll.Initial.Pesticide = "Pest2110.dat";
                    CEEOT_dll.Initial.Espace = " ";
                    CEEOT_dll.Initial.Espace1 = " ";
                    CEEOT_dll.Initial.col1 = 11;
                    CEEOT_dll.Initial.col2 = 9; //Operation code in operation file
                    CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col5 = 23;  //fertilizer code in operatino file
                   CEEOT_dll.Initial.col6 = 7;
                    CEEOT_dll.Initial.col7 = 18;//Crop code in operation file
                    CEEOT_dll.Initial.col8 = 35; //Curve number in operation file
                    CEEOT_dll.Initial.limit = 1;
                    CEEOT_dll.Initial.SubaLines = 12;
                   CEEOT_dll.Initial.cntrl2 = "\\" +CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl = "\\" + CEEOT_dll.Initial.figsFile;
                   CEEOT_dll.Initial.cntrl3 = "\\" + CEEOT_dll.Initial.figsFile;
                   CEEOT_dll.Initial.cntrl4 = "\\file2.cio";
                    CEEOT_dll.Initial.cntrl = "\\file1.cio";
                    CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                    CEEOT_dll.Initial.suba = "suba2110.dat";
                    CEEOT_dll.Initial.Soil = "soil2110.dat";
                    CEEOT_dll.Initial.Site = "site2110.dat";
                    CEEOT_dll.Initial.Opcs = "opcs2110.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm12110.dat";
                    CEEOT_dll.Initial.parm = "parm0604.dat";
                    CEEOT_dll.Initial.Till = "Till2110.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert2110.dat";
                    CEEOT_dll.Initial.herd = "Herd0604.dat";
                    CEEOT_dll.Initial.sublines = ""+12;
                    CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "32":
                    CEEOT_dll.Initial.Year_Col = 2;  //year size
                   CEEOT_dll.Initial.Version = "4.2.0";
                    CEEOT_dll.Initial.Version2 = "APEX0604 + SWAT2009";
                    CEEOT_dll.Initial.Pesticide = "Pest2110.dat";
                    CEEOT_dll.Initial.Espace = " ";
                    CEEOT_dll.Initial.Espace1 = " ";
                    CEEOT_dll.Initial.col1 = 11;
                    CEEOT_dll.Initial.col2 = 9; //Operation code in operation file
                    CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col5 = 23;   //fertilizer code in operatino file
                    CEEOT_dll.Initial.col6 = 7;
                    CEEOT_dll.Initial.col7 = 18;//Crop code in operation file
                    CEEOT_dll.Initial.col8 = 35; //Curve number in operation file
                    CEEOT_dll.Initial.limit = 1;
                    CEEOT_dll.Initial.SubaLines = 12;
                    CEEOT_dll.Initial.cntrl2 = "\\" + CEEOT_dll.Initial.figsFile;
                   CEEOT_dll.Initial.cntrl = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl3 = "\\" + CEEOT_dll.Initial.figsFile;
                   CEEOT_dll.Initial.cntrl4 = "\\file2.cio";
                    CEEOT_dll.Initial.cntrl = "\\file1.cio";
                    CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                    CEEOT_dll.Initial.suba = "suba2110.dat";
                    CEEOT_dll.Initial.Soil = "soil2110.dat";
                    CEEOT_dll.Initial.Site = "site2110.dat";
                    CEEOT_dll.Initial.Opcs = "opcs2110.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm12110.dat";
                    CEEOT_dll.Initial.parm = "parm0604.dat";
                    CEEOT_dll.Initial.Till = "Till2110.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert2110.dat";
                    CEEOT_dll.Initial.herd = "Herd0604.dat";
                    CEEOT_dll.Initial.sublines = ""+12;
                    CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
                case "33":
                    CEEOT_dll.Initial.Year_Col = 2; //year size
                    CEEOT_dll.Initial.Version = "4.3.0";
                    CEEOT_dll.Initial.Version2 = "APEX0604 + SWAT2012";
                    CEEOT_dll.Initial.Pesticide = "Pest2110.dat";
                    CEEOT_dll.Initial.Espace = " ";
                    CEEOT_dll.Initial.Espace1 = " ";
                    CEEOT_dll.Initial.col1 = 11;
                    CEEOT_dll.Initial.col2 = 9; //Operation code in operation file
                    CEEOT_dll.Initial.col3 = 7;
                    CEEOT_dll.Initial.col4 = 2;
                    CEEOT_dll.Initial.col5 = 23;   //fertilizer code in operatino file
                    CEEOT_dll.Initial.col6 = 7;
                    CEEOT_dll.Initial.col7 = 18; //Crop code in operation file
                    CEEOT_dll.Initial.col8 = 35; //Curve number in operation file
                    CEEOT_dll.Initial.limit = 1;
                    CEEOT_dll.Initial.SubaLines = 12;
                    CEEOT_dll.Initial.cntrl2 = "\\" + CEEOT_dll.Initial.figsFile;
                    CEEOT_dll.Initial.cntrl = "\\" + CEEOT_dll.Initial.figsFile;
                   CEEOT_dll.Initial.cntrl3 = "\\" + CEEOT_dll.Initial.figsFile;
                   CEEOT_dll.Initial.cntrl4 = "\\file2.cio";
                    CEEOT_dll.Initial.cntrl = "\\file1.cio";
                    CEEOT_dll.Initial.RunFile = "\\ApexRun.dat";
                    CEEOT_dll.Initial.suba = "suba2110.dat";
                    CEEOT_dll.Initial.Soil = "soil2110.dat";
                    CEEOT_dll.Initial.Site = "site2110.dat";
                    CEEOT_dll.Initial.Opcs = "opcs2110.dat";
                    CEEOT_dll.Initial.wpm1 = "wpm12110.dat";
                    CEEOT_dll.Initial.parm = "parm0604.dat";
                    CEEOT_dll.Initial.Till = "Till2110.dat";
                    CEEOT_dll.Initial.Fertilizer = "Fert2110.dat";
                    CEEOT_dll.Initial.herd = "Herd0604.dat";
                    CEEOT_dll.Initial.sublines =""+ 12;
                    CEEOT_dll.Initial.cont = "Apexcont.dat";
                    break;
            }

            CEEOT_dll.Initial.cntrl1 = CEEOT_dll.Initial.cntrl;
        }
    }
}
