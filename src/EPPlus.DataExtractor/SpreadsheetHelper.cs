using System;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

[assembly: InternalsVisibleTo("EPPlus.DataExtractor.Tests")]

namespace EPPlus.DataExtractor
{
    internal static class SpreadsheetHelper
    {
        private const int ACharValue = 'A';
        private const int IntervalValue = 'Z' - ACharValue + 1;

        /// <summary>
        /// Converts a string column header with letters only
        /// to a numeric value representing the column.
        /// </summary>
        /// <param name="columnHeader"></param>
        /// <returns></returns>
        public static int ConvertColumnHeaderToNumber(string columnHeader)
        {
            if (string.IsNullOrWhiteSpace(columnHeader))
                throw new ArgumentNullException(nameof(columnHeader));


            columnHeader = columnHeader.ToUpperInvariant();

            if (!Regex.IsMatch(columnHeader, "^[A-Z]+$"))
                throw new ArgumentException("The given column header is in an invalid format. Only A to Z letters are supported.");

            int columnIndex = columnHeader.Last() - ACharValue + 1;
            for (int index = 0; index < columnHeader.Length - 1; index++)
            {
                int letterValue = columnHeader[index] - ACharValue + 1;

                int power = columnHeader.Length - (index + 1);

                columnIndex += (letterValue * (int) Math.Pow(IntervalValue, power));
            }

            return columnIndex;
        }
    }
}
