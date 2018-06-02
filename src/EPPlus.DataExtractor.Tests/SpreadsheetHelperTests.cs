using Xunit;

namespace EPPlus.DataExtractor.Tests
{
    public class SpreadsheetHelperTests
    {
        [Theory]
        [InlineData("A", 1)]
        [InlineData("Z", 26)]
        [InlineData("AA", (26 * 1) + 1)]
        [InlineData("AB", (26 * 1) + 2)]
        [InlineData("BA", (26 * 2) + 1)]
        [InlineData("ZA", (26 * 26) + 1)]
        [InlineData("ZZ", (26 * 26) + 26)]
        [InlineData("AAA", (26 * 26) + 26 + 1)]
        [InlineData("AAB", (26 * 26) + 26 + 2)]
        [InlineData("AAZ", (26 * 26) + 26 + 26)]
        [InlineData("ABA", (26 * 26) + (26 * 2) + 1)]
        [InlineData("CBZ", (26 * 26 * 3) + (26 * 2) + 26)]
        public void ShouldConvertSimpleLetterString(string columnString, int expectedResult)
        {
            var result = SpreadsheetHelper.ConvertColumnHeaderToNumber(columnString);
            Assert.Equal(expectedResult, result);
        }
    }
}
