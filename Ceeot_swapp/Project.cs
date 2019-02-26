﻿using System;
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

        private String name;
        private String location;
        private String swattLocation;
        private ProjectVersion apexVersion;
        private ProjectVersion swattVersion;
        
        public String Name { get { return this.name; } set { this.name = value; } }
        public String Location { get { return this.location; } set { this.location = value; } }
        public String SwattLocation { get { return this.swattLocation; } set { this.swattLocation = value; }  }
        public ProjectVersion ApexVersion { get { return this.apexVersion; } set { this.apexVersion = value; }  }
        public ProjectVersion SwattVersion { get { return this.swattVersion; } set { this.swattVersion= value; }  }
        
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
