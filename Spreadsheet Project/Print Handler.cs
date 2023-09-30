using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spreadsheet_Project
{
    public class PrintHandler
    {
        readonly Sheet sheetReference;

        (int start, int end) colDisplay;
        (int start, int end) rowDisplay;

        public PrintHandler(Sheet sheetReference)
        {
            this.sheetReference = sheetReference;
            UpdateConsoleSize();
        }

        /// <summary>
        /// Updates the number of rows and columns to be printed based on a possible window resieze made by the user.
        /// </summary>
        internal void UpdateConsoleSize()
        {
            try
            {
                colDisplay = (1, Console.WindowWidth / 10);
                rowDisplay = (1, Console.WindowHeight - 3);

            }
            catch { }
            
        }

        /// <summary>
        /// Prints the entire sheet.
        /// </summary>
        public void PrintSheet()
        {
            Console.CursorVisible = false;

            AdjustColAndRowDisplay();

            PrintPositionAndCellContents(colDisplay);

            for (int row = rowDisplay.start; row <= rowDisplay.end; row++)
            {
                for (int col = colDisplay.start; col <= colDisplay.end; col++)
                {
                    if (col == colDisplay.start)
                    {
                        PrintCell(row, 0);
                    }
                    PrintCell(row, col);
                }

                Console.WriteLine();
            }
        }

        /// <summary>
        /// Prints the sheet header where the current position and cell contents (Formula takes priority over contents).
        /// </summary>
        private void PrintPositionAndCellContents((int start, int end) colDisplay)
        {
            Console.SetCursorPosition(0, 0);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, 0);

            var currentPosition = sheetReference.currentPosition;
            var position = $"{sheetReference.Alphabet[currentPosition.Y]}{currentPosition.X}";
            var contents = string.Empty;

            if (sheetReference.cellData.TryGetValue(currentPosition, out var cell))
            {
                contents = cell.Formula is null ? cell.Content.TypeValue.ToString() : cell.Formula;
            }

            Console.WriteLine($"{sheetReference.currentSheetName}|CTRL+L:Leg|XY: {position} [{contents}]");

            PrintCell(0, 0);

            for (int i = colDisplay.start; i <= colDisplay.end; i++)
            {
                PrintCell(0, i);
            }

            Console.WriteLine();
        }

        /// <summary>
        /// Prints the proper position value. Furthermore, it takes care of the formatting based on the cursor position.
        /// </summary>
        private void PrintCell(int row, int col)
        {
            var cellPrintValue = GetPrintableCellValue(row, col);

            cellPrintValue = cellPrintValue.Length > 10 ? cellPrintValue[..10] : cellPrintValue;

            HighlightCellIfHeadingOrCursor(row, col);

            if (col == 0)
            {
                Console.Write(string.Format("{0,4}", cellPrintValue));
            }
            else if (col == colDisplay.end)
            {
                Console.Write(string.Format("{0,-6}", string.Empty));
            }
            else
            {
                Console.Write(
                    string.Format(
                        "{0,-10}",
                        string.Format(
                            "{0," + ((10 + cellPrintValue.Length) / 2).ToString() + "}",
                            cellPrintValue)));
            }

            Console.ResetColor();
        }



        /// <summary>
        /// Returns the proper printable value, either a row/col notation or a cell value if it exists.
        /// </summary>
        private string GetPrintableCellValue(int row, int col)
        {
            if (row == 0)
            {
                return CreationHandler.GetReferenceColumnNumber(col);
            }
            else if (col == 0)
            {
                return row.ToString();
            }
            else if (sheetReference.cellData.TryGetValue((row, col), out var cell))
            {
                if (cell.Formula is not null)
                {
                    UpdateContentBasedOnFormula(row, col);
                }

                return cell.Content.TypeValue.ToString();
            }

            return string.Empty;
        }


        /// <summary>
        /// Updates the contents of the cell based on the formula. This is to cover the case where formula dependent cells changed.
        /// </summary>
        internal void UpdateContentBasedOnFormula(int row, int col)
        {
            var formula = sheetReference.cellData[(row, col)].Formula;

            sheetReference.cellData[(row, col)] = (
                CreationHandler.FindValueType(formula),
                null
            );

            var updatedFormulaResult = sheetReference.formulaHandler.TryFormulas((row, col));

            if (updatedFormulaResult is not null)
            {
                sheetReference.cellData[(row, col)] = (
                    CreationHandler.FindValueType(updatedFormulaResult),
                    sheetReference.cellData[(row, col)].Content.TypeValue.ToString()
                );
            }
        }

        /// <summary>
        /// Changes the BackgroundColor and ForegroudColor based on if the cursor position is a heading or currentPosition.
        /// </summary>
        private void HighlightCellIfHeadingOrCursor(int row, int col)
        {
            if ((row == 0 && col != sheetReference.currentPosition.Y)
                || (col == 0 && row != sheetReference.currentPosition.X)
                || (row, col) == sheetReference.currentPosition)
            {
                Console.BackgroundColor = ConsoleColor.Gray;
                Console.ForegroundColor = ConsoleColor.Black;
            }
        }

        /// <summary>
        /// Updates and adjusts colDisplay and row Display based on the cursor Current Position. This allows for an infinite spreadsheet illusion.
        /// </summary>
        private void AdjustColAndRowDisplay()
        {
            if (sheetReference.currentPosition.Y > colDisplay.end - 1)
            {
                colDisplay.start++;
                colDisplay.end++;
            }
            else if (sheetReference.currentPosition.Y < colDisplay.start)
            {
                colDisplay.start--;
                colDisplay.end--;
            }

            if (sheetReference.currentPosition.X > rowDisplay.end)
            {
                rowDisplay.start++;
                rowDisplay.end++;
            }
            else if (sheetReference.currentPosition.X < rowDisplay.start)
            {
                rowDisplay.start--;
                rowDisplay.end--;
            }
        }

    }
}
