using System;

namespace EntityFrameworkCoreSample.Model
{
    public class MonthlyRevenueEntity
    {
        public Guid Id { get; set; }

        public decimal Value { get; set; }

        public DateTime MonthYear { get; set; }

        public DateTime DateCreated { get; set; }

        public override string ToString() =>
            $"Id={Id}, MonthYear={MonthYear.ToString("Y")}, Value={Value}, DateCreated={DateCreated.ToShortDateString()}";
    }
}
