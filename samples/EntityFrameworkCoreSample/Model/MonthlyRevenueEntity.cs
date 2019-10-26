using System;

namespace EntityFrameworkCoreSample.Model
{
    public class MonthlyRevenueEntity
    {
        public static string TableName => "Revenues";

        public Guid Id { get; set; }

        public decimal Value { get; set; }

        public DateTime MonthYear { get; set; }

        public string BranchId { get; set; }
        public BranchEntity Branch { get; set; }

        public override string ToString() =>
            $"Id={Id}, BranchId={BranchId}, MonthYear={MonthYear.ToString("Y")}, Value={Value}";
    }
}
