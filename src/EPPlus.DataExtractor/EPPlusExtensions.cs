using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DataExtractor
{
    public static class EPPlusExtensions
    {
        /// <summary>
        /// Creates an <see cref="IDataExtractor{TRow}"/> to extract data from
        /// the worksheet to <typeparamref name="TRow"/> objects.
        /// </summary>
        /// <exception cref="ArgumentNullException">Thrown if the <paramref name="worksheet"/> is null.</exception>
        /// <typeparam name="TRow">The type that will be populated
        /// with the data from the worksheet.</typeparam>
        /// <param name="worksheet">The worksheet parameter.</param>
        /// <returns>An instance of <see cref="IDataExtractor"/>.</returns>
        public static IDataExtractor<TRow> Extract<TRow>(this ExcelWorksheet worksheet)
            where TRow : class, new()
        {
            if (worksheet == null)
                throw new ArgumentNullException("worksheet");

            return new DataExtractor<TRow>(worksheet);
        }
    }
}
