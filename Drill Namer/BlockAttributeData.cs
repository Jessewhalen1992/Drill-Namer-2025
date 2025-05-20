using Autodesk.AutoCAD.DatabaseServices;

namespace Drill_Namer.Models
{
    public class BlockAttributeData
    {
        public BlockReference BlockReference { get; set; }
        public string DrillName { get; set; }
        public double YCoordinate { get; set; }
    }
}
