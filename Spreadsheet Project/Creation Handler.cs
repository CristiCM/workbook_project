namespace Spreadsheet_Project
{
    public class CreationHandler
    {
        public CreationHandler(Sheet sheetReference)
        {
            sheetReference.Alphabet = new();

            _ = sheetReference.Alphabet.Append("");

            for (int i = 0; i < 16384; i++)
            {
                sheetReference.Alphabet.Add(GetReferenceColumnNumber(i));
            }
        }

        /// <summary>
        /// Creates and returns a proper IValue object based on the cell contents type.
        /// </summary>
        internal static IValue FindValueType(string cellContents)
        {
            return cellContents switch
            {
                string when int.TryParse(cellContents, out int result) => new IntType(result),
                string when double.TryParse(cellContents, out double result) && cellContents.Last() != '.' => new DoubleType(result),
                _ => new StringType(cellContents)
            };
        }

        /// <summary>
        /// Converting and returning a base-10 number to a base-26 number, where each digit corresponds to a letter in the alphabet
        /// (A = 1, B = 2, C = 3, ..., Z = 26).
        /// </summary>
        public static string GetReferenceColumnNumber(int columnNumber)
        {
            string columnName = string.Empty;
            int remaining;

            while (columnNumber > 0)
            {
                remaining = (columnNumber - 1) % 26;
                columnName = Convert.ToChar(65 + remaining).ToString() + columnName;
                columnNumber = (int)((columnNumber - remaining) / 26);
            }

            return columnName;
        }

    }
}
