using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ceeot_swapp
{
    public class Project
    {
        public enum ProjectVersion
        {
            APEX_0604, APEX_0806, SWATT_2005, SWATT_2009, SWATT_2012,
        }

        public String name;
        public String location;
        public String swattLocation;
        public ProjectVersion apexVersion;
        public ProjectVersion swattVersion;
        public string initialProject;
        public string initialScenario;

       
        public struct SubBasin
        {
            public string name;
            private bool selected;

            public Boolean Selected
            {
                get { return this.selected; }
                set { this.selected = value; }
            }
            public String Name
            {
                get { return this.name; }
                set { this.name = value; }
            }
        }

        private List<SubBasin> subBasins;

        public Project()
        {
            subBasins = new List<SubBasin>();
        }

        public List<SubBasin> SubBasins
        {
            get { return this.subBasins; }
        }

        public List<string> SelectedSubBasins
        {
            get
            {
                List<string> basins = new List<string>();
                foreach (SubBasin s in subBasins)
                {
                    if (s.Selected) basins.Add(s.Name);
                }
                return basins;
            }
        }

        public List<string> AllSubBasins
        {
            get
            {
                List<string> basins = new List<string>();
                foreach (SubBasin s in subBasins)
                {
                    basins.Add(s.Name);
                }
                return basins;
            }
        }
    }

}
