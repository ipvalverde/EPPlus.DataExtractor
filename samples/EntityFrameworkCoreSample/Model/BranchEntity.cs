using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EntityFrameworkCoreSample.Model
{
    public class BranchEntity
    {
        public static string TableName => "Branches";

        public string Id { get; set; }

        public string Name { get; set; }

        public string Location { get; set; }

        public string Phone { get; set; }

        public ICollection<MonthlyRevenueEntity> Revenues { get; set; }

        public BranchEntity()
        {
            this.Revenues = new HashSet<MonthlyRevenueEntity>();
        }

        public override string ToString()
        {
            var vehiclesStringBuilder = Revenues
                .Select((v, index) => new
                {
                    Index = index,
                    VehicleStr = v.ToString(),
                })
                .Aggregate(
                    new StringBuilder(),
                    (ac, cu) => ac.AppendLine($"Revenues[{cu.Index}]={cu}"));

            return $@"{nameof(BranchEntity)}
Id={Id}
Name={Name}
Location={Location}
Phone={Phone}
{vehiclesStringBuilder.ToString()}

----------------------------------------";
        }
    }
}
