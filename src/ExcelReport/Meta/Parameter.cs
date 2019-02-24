using System.Collections.Generic;

namespace ExcelReport.Meta
{
    public class Parameter : Named
    {
        public List<Location> Locations { get; } = new List<Location>();

        public void Append(Location location)
        {
            Locations.Add(location);
        }
    }
}