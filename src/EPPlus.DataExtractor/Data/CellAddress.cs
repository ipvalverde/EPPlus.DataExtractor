using OfficeOpenXml;

namespace EPPlus.DataExtractor.Data
{
    /// <summary>
    /// Represents a specific cell address.
    /// </summary>
    public class CellAddress
    {
        internal CellAddress(ExcelRangeBase excelAddress)
        {
            Row = excelAddress.Start.Row;
            Column = excelAddress.Start.Column;
            Address = excelAddress.Address;
        }

        /// <summary>
        /// User-friendly cell address using letters for column. Like "B2".
        /// </summary>
        public string Address { get; private set; }

        /// <summary>
        /// The index of the columns, starting at one.
        /// </summary>
        public int Column { get; private set; }

        /// <summary>
        /// The index of the row, starting at one.
        /// </summary>
        public int Row { get; private set; }
    }
}
