using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ceeot_swapp
{
    public class SwattProject
    {
        public enum ProjectVersion
        {
            APEX_0604, APEX_0806, SWATT_2005, SWATT_2009, SWATT_2012,
        }

        private String name;
        private String location;
        private String swattLocation;
        private String currentScenario;

        private ProjectVersion apexVersion;
        private ProjectVersion swattVersion;
        
        public String Name { get { return this.name; } set { this.name = value; } }
        public String Location { get { return this.location; } set { this.location = value; } }
        public String SwattLocation { get { return this.swattLocation; } set { this.swattLocation = value; }  }
        public String CurrentScenario { get { return this.currentScenario; } set { this.currentScenario = value; } }

        public ProjectVersion ApexVersion { get { return this.apexVersion; } set { this.apexVersion = value; }  }
        public ProjectVersion SwattVersion { get { return this.swattVersion; } set { this.swattVersion= value; }  }

        private List<SubBasin> subBasins;

        public ApexProject toApexProject()
        {
            return null;
        }

        public struct HRU {
            CropCodes.Code code;
            String description;
            List<SubBasin> subBasin; 

            public CropCodes.Code Code { get {return code; } set { code = value; } }
            public String Description { get { return description; } set { description = value; } }
            public List<SubBasin> SubBasin { get { return subBasin;  } set { subBasin = value; } }
        }
        
        public class SubBasin
        {
            public string name;
            private bool selected;
            private List<HRU> hrus;

            public Boolean Selected { get { return this.selected; } set { this.selected = value; } }
            public List<HRU> Hrus { get { return this.hrus; } set { this.hrus = value;  } }
            public String Name { get { return this.name; } set { this.name = value; } }

            public SubBasin(){
                hrus = new List<HRU>();
            }
        }
        
        public SwattProject()
        {
            subBasins = new List<SubBasin>();
        }

        public List<SubBasin> SubBasins
        {
            get { return this.subBasins; }
            set { this.subBasins = value; }
        }

        public List<HRU> SelectedSubBasinHrus
        {
            get
            {
                List<HRU> hrus = new List<HRU>();
                foreach (SubBasin s in this.SubBasins)
                {
                    // If the sub basin was selected add its 
                    if (s.Selected)
                    {
                        s.Hrus.ForEach(h => hrus.Add(h));
                    }
                }
                return hrus;
            }
        }

        
    }
}
