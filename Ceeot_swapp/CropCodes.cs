using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ceeot_swapp
{
    public sealed class CropCodes 
    {
        public enum Code
        {
            PAST,
            AGRR,
            FRSD
        }

        public static String getDescription(Code code)
        {
            switch(code)
            {
                case Code.PAST: return "Pasture";
                case Code.AGRR: return "Agricultural Land-Row Crops";
                case Code.FRSD: return "Forrest Deciduous";
                default: return "I dunno!!!";
            }
        }
    }
    
}
