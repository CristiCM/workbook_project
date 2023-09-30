namespace Spreadsheet_Project.Formulas
{
    public class Formulas
    {
        readonly Sheet sheetReference;

        List<(int, string)>? subtotalCells;

        public Formulas(Sheet sheetReference)
        {
            subtotalCells = null;
            this.sheetReference = sheetReference;
        }

        /// <summary>
        /// Returns a list of delegate functions that can be used to perform calculations on cells.
        /// </summary>
        public IEnumerable<
            Func<(int X, int Y), (bool formulaMatch, string formulaResult)>
        > GetAllAvailableFormulas()
        {
            return new List<Func<(int X, int Y), (bool formulaMatch, string formulaResult)>>()
            {
                AddCellReference,
                Now,
                Today,
                StringOverride,
                Sum,
                Avg,
                Count,
                Subtotal,
                Mod,
                Power,
                Ceiling,
                Floor,
                Len,
                Concat,
                Replace,
                Substitute,
                Vlookup
            };
        }

        /// <summary>
        /// Tries to lookup a value from a table in a specific column given based on a search value given.
        /// </summary>
        private (bool formulaMatch, string formulaResult) Vlookup((int X, int Y) currentPosition)
        {
            var (formulaPass, errorMessage) =
                VlookupInitialChecks(currentPosition, out string lookUpValue, out List<(int sheetRefereceIndex, string cellOrValue)> lookUpTableRange, out int columnIndexNumber);

            if (!formulaPass)
            {
                return (errorMessage);
            }

            try
            {
                var (sheetRefereceIndex, cellOrValue) = FindFirstLookupValueOccurrence(lookUpValue, lookUpTableRange);

                var resultTableArrayReference = GetReferenceCellColumnLetterOrRowNumber(lookUpTableRange[columnIndexNumber - 1].cellOrValue) +
                    GetReferenceCellColumnLetterOrRowNumber(cellOrValue, rowNumber: true);

                GetCellReferencePositionIfValid(resultTableArrayReference, out var cellPosition);

                return (true, sheetReference.sheets[sheetRefereceIndex].sheet.cellData[cellPosition].Content.TypeValue.ToString());
            }
            catch (Exception)
            {
                return (true, "#N/A");
            }
        }

        /// <summary>
        /// Validates the formula for nameMatch, recursion, correct text format, correct index type, index in range. 
        /// </summary>
        private (bool formulaPass, (bool formulaMatch, string formulaResult) errorMessage) VlookupInitialChecks(
            (int X, int Y) currentPosition,
            out string outLookUpValue, 
            out List<(int sheetRefereceIndex, string cellOrValue)> outLookUpTable, 
            out int outColumnIndexNumber)
        {
            outLookUpValue = string.Empty;

            outLookUpTable = new();

            outColumnIndexNumber = -1;

            if (!GetElementsIfFormulaNameMatch(currentPosition, "=VLOOKUP", out List<(int sheetRefereceIndex, string cellOrValue)> referenceIndexesAndElements))
                return (false, (false, string.Empty));

            if (!CheckIfFormulaIsRecursive(referenceIndexesAndElements, currentPosition))
                return (false, (true, "RecursErr"));

            if (!CheckIfFormulaTextElementHasQuotations(referenceIndexesAndElements[0], out var lookupValue))
                return (false, (true, "#NAME"));

            if (!int.TryParse(referenceIndexesAndElements[^1].cellOrValue, out var columnIndexNumber))
                return (false, (false, string.Empty));

            if (sheetReference.Alphabet.IndexOf(GetReferenceCellColumnLetterOrRowNumber(referenceIndexesAndElements[^2].cellOrValue)) < columnIndexNumber)
                return (false, (true, "#REF!"));

            outLookUpValue = lookupValue;

            outLookUpTable = referenceIndexesAndElements.GetRange(1, referenceIndexesAndElements.Count - 2);

            outColumnIndexNumber = columnIndexNumber;

            return (true, (true, string.Empty));
        }

        /// <summary>
        /// Vlookup helper function. Returns the sheetReferenceIndex and cell/value from a given TableRange.
        /// </summary>
        private (int sheetRefereceIndex, string cellOrValue) FindFirstLookupValueOccurrence(string lookupValue, List<(int sheetRefereceIndex, string cellOrValue)> lookupTableRange)
        {
            return lookupTableRange.First(cell =>
                GetReferenceCellColumnLetterOrRowNumber(cell.cellOrValue) == GetReferenceCellColumnLetterOrRowNumber(lookupTableRange[0].cellOrValue) &&
                GetCellReferencePositionIfValid(cell.cellOrValue, out var cellPosition) &&
                sheetReference.sheets[cell.sheetRefereceIndex].sheet.cellData[cellPosition].Content.TypeValue.ToString() == lookupValue
            );
        }

        /// <summary>
        /// Tries to substitute a specified string in a given text with another string.
        /// </summary>
        private (bool Match, string formulaResult) Substitute((int X, int Y) currentPosition)
        {
            var (formulaPass, errorMessage) = SubstituteInitialChecks(currentPosition, out string text, out string oldText, out string newText, out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements);

            if (!formulaPass)
            {
                return (errorMessage);
            }

            if (formulaElements.Count == 3)
            {
                return (true, text.Replace(oldText, newText));
            }
            else if (int.TryParse(formulaElements[3].cellOrValue, out int occurence))
            {
                var indexes = Enumerable
                    .Range(0, text.Length - oldText.Length + 1)
                    .Where(i => text.Substring(i, oldText.Length) == oldText);

                int selectedIndex = indexes.ElementAtOrDefault(occurence - 1);

                return (
                    true,
                    occurence < 0 || occurence > indexes.Count()
                        ? text
                        : text[..selectedIndex] + newText + text[(selectedIndex + newText.Length)..]
                );
            }

            return (false, string.Empty);
        }

        /// <summary>
        /// Validates the formula for nameMatch, recursion, correct text format. 
        /// </summary>
        private (bool formulaPass, (bool formulaMatch, string formulaResult) errorMessage) SubstituteInitialChecks(
            (int X, int Y) currentPosition,
            out string outText,
            out string outOldText,
            out string outNewText,
            out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements)
        {
            (outText, outOldText, outNewText) = (string.Empty, string.Empty, string.Empty);

            formulaElements = new();

            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=SUBSTITUTE",
                out List<(int sheetRefereceIndex, string cellOrValue)> referenceIndexesAndCells
            ) || referenceIndexesAndCells.Count < 3 || referenceIndexesAndCells.Count > 4)
            {
                return (false, (false, string.Empty));
            }

            if (!CheckIfFormulaIsRecursive(referenceIndexesAndCells, currentPosition))
            {
                return (false, (true, "RecursErr"));
            }

            if (!CheckIfFormulaTextElementHasQuotations(referenceIndexesAndCells[0], out string text)
                || !CheckIfFormulaTextElementHasQuotations(referenceIndexesAndCells[1], out string oldText)
                || !CheckIfFormulaTextElementHasQuotations(referenceIndexesAndCells[2], out string newText))
            {
                return (false, (true, "#NAME"));
            }

            (outText, outOldText, outNewText) = (text, oldText, newText);

            formulaElements = referenceIndexesAndCells;

            return (true, (true, string.Empty));
        }

        ///// <summary>
        ///// Tries to replace a specified number of characters in a given string with another string starting at a specified position.
        ///// </summary>
        private (bool Match, string formulaResult) Replace((int X, int Y) currentPosition)
        {
            var (formulaPass, errorMessage) = ReplaceInitialChecks(currentPosition, out int startIndex, out int numChars, out string oldText, out string newText);

            if (!formulaPass)
            {
                return errorMessage;
            }

            var result =
                oldText.Length > startIndex
                    ? oldText[..(startIndex - 1)] + newText + oldText[(startIndex - 1 + numChars)..]
                    : oldText + newText;

            return (true, result);
        }

        /// <summary>
        /// Validates the formula for nameMatch, recursion, correct text format, correct index type. 
        /// </summary>
        private (bool formulaPass, (bool formulaMatch, string formulaResult) errorMessage) ReplaceInitialChecks(
            (int X, int Y) currentPosition,
            out int outStartIndex,
            out int outNumChars,
            out string outOldText,
            out string outNewText)
        {
            (outStartIndex, outNumChars, outOldText, outNewText) = (0, 0, string.Empty, string.Empty);

            if (!GetElementsIfFormulaNameMatch(
                    currentPosition,
                    "=REPLACE",
                    out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements
                ) || formulaElements.Count != 4)
            {
                return (false, (false, string.Empty));
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
            {
                return (false, (true, "RecursErr"));
            }

            if (!int.TryParse(formulaElements[1].cellOrValue, out int startIndex)
                || startIndex <= 0
                || !int.TryParse(formulaElements[2].cellOrValue, out int numChars)
                || numChars < 0)
            {
                return (false, (true, "#NAME"));
            }

            if (!CheckIfFormulaTextElementHasQuotations(formulaElements[0], out string oldText)
                || !CheckIfFormulaTextElementHasQuotations(formulaElements[3], out string newText))
            {
                return (false, (true, "#NAME"));
            }

            (outStartIndex, outNumChars, outOldText, outNewText) = (startIndex, numChars, oldText, newText);

            return (true, (true, string.Empty));
        }

        ///// <summary>
        ///// Tries to concatenate values from cells.
        ///// </summary>
        private (bool Match, string formulaResult) Concat((int X, int Y) currentPosition)
        {
            var (formulaPass, errorMessage) = 
                ConcatInitialChecks(currentPosition, out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements, out IEnumerable<string> nonReferenceElements);

            if (!formulaPass)
            {
                return errorMessage;
            }

            var concatResult = formulaElements.Aggregate(
                "",
                (acc, cur) =>
                    acc += nonReferenceElements.Contains(GetReferenceCellValueOrDirectValue(cur))
                        ? GetReferenceCellValueOrDirectValue(cur)[1..^1]
                        : GetReferenceCellValueOrDirectValue(cur)
            );

            return (true, concatResult);
        }

        /// <summary>
        /// Validates the formula for nameMatch, recursion, correct text format. 
        /// </summary>
        private (bool formulaPass, (bool formulaMatch, string formulaResult) errorMessage) ConcatInitialChecks(
            (int X, int Y) currentPosition,
            out List<(int sheetRefereceIndex, string cellOrValue)> outFormulaElements,
            out IEnumerable<string> outNonReferenceElements)
        {
            outFormulaElements = new();
            outNonReferenceElements = Array.Empty<string>();

            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=CONCATENATE",
                out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements
            ))
            {
                return (false, (false, string.Empty));
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
            {
                return (false, (true, "RecursErr"));
            }

            var nonReferenceElements = formulaElements
                .Select(refIndexAndCell => refIndexAndCell.cellOrValue)
                .Where(cellValue => !GetCellReferencePositionIfValid(cellValue, out _));

            if (!nonReferenceElements.All(value => value[0] == '"' && value[^1] == '"'))
            {
                return (false, (true, "#NAME"));
            };

            outFormulaElements = formulaElements;
            outNonReferenceElements = nonReferenceElements;

            return (true, (true, string.Empty));
        }

        ///// <summary>
        ///// Tries to return the length of the value found in the specified cell.
        ///// </summary>
        private (bool Match, string formulaResult) Len((int X, int Y) currentPosition)
        {
            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=LEN",
                out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements
                )
                || formulaElements.Count != 1
            )
            {
                return (false, string.Empty);
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
                return (true, "RecursErr");
            

            if (!CheckIfFormulaTextElementHasQuotations(formulaElements[0], out string lenString))
                return (true, "#NAME");
            

            return (true, lenString.Length.ToString());
        }

        ///// <summary>
        ///// Tries to round down the value found in the specified cell.
        ///// </summary>
        private (bool Match, string formulaResult) Floor((int X, int Y) currentPosition)
        {
            var result = (false, string.Empty);

            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=FLOOR",
                out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements
                )
                || formulaElements.Count != 2
            )
            {
                return result;
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
                return (true, "RecursErr");

            try
            {
                ConvertValueToProperType(formulaElements[0], out dynamic number);
                ConvertValueToProperType(formulaElements[1], out dynamic significance);
                number = number is int ? (double)number : number;
                significance = significance is int ? (double)significance : significance;
                result = (true, (Math.Floor(number / significance) * significance).ToString());
            }
            catch (Exception)
            {
                result = (true, "#VALUE!");
            }

            return result;
        }

        ///// <summary>
        ///// Tries to round up the value found in the specified cell.
        ///// </summary>
        private (bool Match, string formulaResult) Ceiling((int X, int Y) currentPosition)
        {
            var result = (false, string.Empty);

            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=CEILING",
                out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements
                )
                || formulaElements.Count != 2
            )
            {
                return result;
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
                return (true, "RecursErr");

            try
            {
                ConvertValueToProperType(formulaElements[0], out dynamic number);
                ConvertValueToProperType(formulaElements[1], out dynamic significance);
                number = number is int ? (double)number : number;
                significance = significance is int ? (double)significance : significance;
                result = (true, (Math.Ceiling(number / significance) * significance).ToString());
            }
            catch (Exception)
            {
                result = (true, "#VALUE!");
            }

            return result;
        }

        ///// <summary>
        ///// Tries to calculate the power of a number by raising it to a specified power.
        ///// </summary>
        private (bool Match, string formulaResult) Power((int X, int Y) currentPosition)
        {
            var result = (false, string.Empty);

            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=POWER",
                out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements
                )
                || formulaElements.Count != 2
            )
            {
                return result;
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
                return (true, "RecursErr");

            try
            {
                ConvertValueToProperType(formulaElements[0], out dynamic number);
                ConvertValueToProperType(formulaElements[1], out dynamic power);
                result = (true, Math.Pow(number, power).ToString());
            }
            catch (Exception)
            {
                result = (true, "#VALUE!");
            }

            return result;
        }

        ///// <summary>
        ///// Tries to calculate the modulo operation on two values.
        ///// </summary>
        private (bool Match, string formulaResult) Mod((int X, int Y) currentPosition)
        {
            var result = (false, string.Empty);

            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=MOD",
                out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements
                )
                || formulaElements.Count != 2
            )
            {
                return result;
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
                return (true, "RecursErr");

            try
            {
                ConvertValueToProperType(formulaElements[0], out dynamic number);
                ConvertValueToProperType(formulaElements[1], out dynamic divisor);
                result = (true, (number % divisor).ToString());
            }
            catch (Exception)
            {
                result = (true, "#VALUE!");
            }

            return result;
        }

        ///// <summary>
        ///// Tries to evaluate a SUBTOTAL formula.
        ///// </summary>
        private (bool Match, string formulaResult) Subtotal((int X, int Y) currentPosition)
        {
            var functionResult = (false, string.Empty);

            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=SUBTOTAL",
                out List<(int sheetRefereceIndex, string cellOrValue)> formulaElements)
            )
            {
                return functionResult;
            }

            this.subtotalCells = formulaElements.Skip(1).ToList();

            switch (formulaElements[0].cellOrValue)
            {
                case "1":
                    functionResult = Avg(currentPosition);
                    break;
                case "2":
                    functionResult = Count(currentPosition);
                    break;
                case "9":
                    functionResult = Sum(currentPosition);
                    break;
            }

            subtotalCells = null;

            return functionResult;
        }

        ///// <summary>
        ///// Tries to count the number of non-blank cells in the range of cells specified in the formula.
        ///// </summary>
        private (bool Match, string formulaResult) Count((int X, int Y) currentPosition)
        {
            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=COUNT",
                out List<(int, string)> formulaElements)
            )
            {
                return (false, string.Empty);
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
                return (true, "RecursErr");
            

            return !GetReferenceCellsOrValuesSumAndCount(formulaElements, out _, out int count)
                ? (true, "#NAME?")
                : (true, count.ToString());
        }

        ///// <summary>
        ///// Tries to calculate the average of the range of cells specified in the formula.
        ///// </summary>
        private (bool Match, string formulaResult) Avg((int X, int Y) currentPosition)
        {
            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=AVERAGE",
                out List<(int, string)> formulaElements)
            )
            {
                return (false, string.Empty);
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
                return (true, "RecursErr");
            

            return ((bool Match, string formulaResult))(!GetReferenceCellsOrValuesSumAndCount(formulaElements, out dynamic sum, out int nonStringCells)
                    ? (true, "#NAME?")
                    : (true, (sum / (double)nonStringCells).ToString()));
        }

        ///// <summary>
        ///// Tries to calculate the sum of the range of cells specified in the formula.
        ///// </summary>
        private (bool Match, string formulaResult) Sum((int X, int Y) currentPosition)
        {
            if (!GetElementsIfFormulaNameMatch(
                currentPosition,
                "=SUM",
                out List<(int, string)> formulaElements)
            )
            {
                return (false, string.Empty);
            }

            if (!CheckIfFormulaIsRecursive(formulaElements, currentPosition))
                return (true, "RecursErr");

            return ((bool Match, string formulaResult))(!GetReferenceCellsOrValuesSumAndCount(formulaElements, out dynamic sum, out _)
                    ? (true, "#NAME?")
                    : (true, sum.ToString()));
        }

        ///// <summary>
        ///// Tries to allow anything to take place in a cell as a string using "".
        ///// </summary>
        private (bool Match, string formulaResult) StringOverride((int X, int Y) currentPosition)
        {
            string formula = sheetReference.cellData[currentPosition].Content.TypeValue.ToString();

            if (formula.Length >= 3 && formula[0] == '=')
            {
                formula = formula[1..];
            }
            else
            {
                return (false, string.Empty);
            }

            if (formula[0] == '"' && formula[^1] == '"')
            {
                return (true, formula[1..^1]);
            }

            return (false, string.Empty);
        }

        ///// <summary>
        ///// Displays the current time in HH:mm:ss: format.
        ///// </summary>
        private (bool Match, string formulaResult) Now((int X, int Y) currentPosition)
        {
            if (sheetReference.cellData[currentPosition].Content.TypeValue.ToString().ToUpper()
                == "=NOW()")
            {
                DateTime currentTime = DateTime.Now;
                return (true, currentTime.ToString("HH:mm:ss"));
            }

            return (false, string.Empty);
        }

        ///// <summary>
        ///// Displays the current date in dd-MM-yyyy format.
        ///// </summary>
        private (bool Match, string formulaResult) Today((int X, int Y) currentPosition)
        {
            if (sheetReference.cellData[currentPosition].Content.TypeValue.ToString().ToUpper()
                == "=TODAY()")
            {
                DateTime currentTime = DateTime.Today;
                return (true, currentTime.ToString("dd-MM-yyyy"));
            }

            return (false, string.Empty);
        }

        ///// <summary>
        ///// Tries to add a reference to the current cell. If a chain of references are present the current cell will
        ///// inherit the last link values.
        ///// </summary>
        private (bool Match, string formulaResult) AddCellReference(
            (int X, int Y) currentPosition)
        {
            (int X, int Y) refPosition;
            int referenceSheetIndex;

            var (formulaPass, errorMessage) =
                AddCellReferenceInitialChecks(currentPosition);

            if (!formulaPass)
            {
                return errorMessage;
            }

            (referenceSheetIndex, refPosition) = GetLastReferencedCell(currentPosition);

            if (refPosition == (-1, -1))
                return (true, "RecursErr");

            if (refPosition == (-2, -2))
                return (true, "SheetDeleted");
            

            return ((bool Match, string formulaResult))(
                sheetReference.sheets[referenceSheetIndex].sheet.cellData.ContainsKey(refPosition)
                    ? (true, sheetReference.sheets[referenceSheetIndex].sheet.cellData[refPosition].Content.TypeValue.ToString())
                    : (true, string.Empty)
            );
        }

        /// <summary>
        /// Validates the formula for nameMatch, correct text format, size. 
        /// </summary>
        private (bool formulaPass, (bool formulaMatch, string formulaResult) errorMessage) AddCellReferenceInitialChecks(
            (int X, int Y) currentPosition)
        {
            string currentPositionValue =
                sheetReference.cellData[currentPosition].Content.TypeValue.ToString();

            if (string.IsNullOrEmpty(currentPositionValue) || currentPositionValue[0] != '=')
                return (false, (false, string.Empty));

            (int referenceSheetIndex, _, string referenceStrippedValue) =
                CheckForOtherSheetReference(currentPositionValue[1..]);

            referenceStrippedValue = "=" + referenceStrippedValue;

            if (referenceStrippedValue.Length < 3
                || !GetCellReferencePositionIfValid(referenceStrippedValue, out _)
                || sheetReference.sheets.Count <= referenceSheetIndex
            )
            {
                return (false, (false, string.Empty));
            }

            return (true, (true, string.Empty));
        }

        ///// <summary>
        ///// Tries to return the last reference (X,Y) coordinate in case there is a chain of references. In the case of a circular
        ///// reference chain it returns (-1,-1).
        ///// </summary>
        private (int, (int X, int Y)) GetLastReferencedCell((int X, int Y) currentPosition)
        {
            (int X, int Y) startingPosition = currentPosition;
            int cycleCount = 0;
            int startingSheetIndex = sheetReference.currentSheetIndex;
            int referenceSheetIndex = sheetReference.currentSheetIndex;

            while (currentPosition != (-1, -1) && sheetReference.sheets[referenceSheetIndex].sheet.cellData.ContainsKey(currentPosition))
            {
                cycleCount++;

                if (sheetReference.sheets.Count <= referenceSheetIndex)
                    return (referenceSheetIndex, (-2, -2));

                string formula = sheetReference.sheets[referenceSheetIndex].sheet.cellData[currentPosition].Formula is null
                    ? sheetReference.sheets[referenceSheetIndex].sheet.cellData[currentPosition].Content.TypeValue.ToString()
                    : sheetReference.sheets[referenceSheetIndex].sheet.cellData[currentPosition].Formula;

                var (newSheetReferenceIndex, _, newFormula) = CheckForOtherSheetReference(formula[1..]);
                GetCellReferencePositionIfValid(newFormula, out var newRefPosition);

                if (newRefPosition == (-1, -1))
                    break;

                if (startingPosition == newRefPosition && startingSheetIndex == newSheetReferenceIndex || cycleCount > 100)
                {
                    currentPosition = (-1, -1);
                }
                else
                {
                    currentPosition = newRefPosition;
                    referenceSheetIndex = newSheetReferenceIndex != -1 ? newSheetReferenceIndex : referenceSheetIndex;
                }
            }

            return (referenceSheetIndex, currentPosition);
        }

        ///// <summary>
        ///// Checks if the cells given in the formula overlap with the cell in which the formula was written.
        ///// </summary>
        private bool CheckIfFormulaIsRecursive(
            IEnumerable<(int index, string cell)> formulaElements,
            (int X, int Y) currentPosition
        )
        {
            foreach (var (index, cell) in formulaElements)
            {
                if (
                    GetCellReferencePositionIfValid(cell, out (int X, int Y) refPosition)
                    && currentPosition == refPosition
                    && sheetReference.currentSheetIndex == index
                )
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Tries to calculate the count and sum of all the non-string cell values.
        /// </summary>
        internal bool GetReferenceCellsOrValuesSumAndCount(
            List<(int index, string cell)> formulaElements,
            out dynamic sum,
            out int count
        )
        {
            sum = 0;
            count = 0;

            bool validFormulaElements = formulaElements.All(
                indexAndCellOrValue =>
                    GetCellReferencePositionIfValid(indexAndCellOrValue.cell, out _)
                    || CreationHandler.FindValueType(indexAndCellOrValue.cell)
                        .TypeValue is not string
            );

            foreach (var indexAndCellOrValue in formulaElements)
            {
                if (ConvertValueToProperType(indexAndCellOrValue, out dynamic cellOrNumberValue))
                {
                    sum += cellOrNumberValue;
                    count++;
                }
            }

            return validFormulaElements;
        }

        /// <summary>
        /// Tries to return the cell values given to the formula as a string array.
        /// </summary>
        private bool GetElementsIfFormulaNameMatch(
            (int X, int Y) currentPosition,
            string formulaName,
            out List<(int, string)> referenceIndexesAndCells
        )
        {
            referenceIndexesAndCells = new List<(int, string)>();

            if (subtotalCells is null && !MatchFormulaName(currentPosition, formulaName))
            {
                return false;
            }

            referenceIndexesAndCells = subtotalCells ?? GetReferenceCellsFromFormula(
                sheetReference.cellData[currentPosition].Content.TypeValue.ToString(),
                formulaName.Length
            );

            return true;
        }

        /// <summary>
        /// Extracts the mentioned cells from a formula for either format (a1,b1,c1 / a1:c1).
        /// </summary>
        internal List<(int, string)> GetReferenceCellsFromFormula(string formula, int skipLength)
        {
            List<(int, string)> referenceIndexesAndCells = new();

            string trimmedFormula = formula[skipLength..].Trim();
            string[] cells = trimmedFormula.Trim('(', ')').Split(',', StringSplitOptions.TrimEntries);

            foreach (var cell in cells)
            {
                if (cell.Contains(':'))
                {
                    referenceIndexesAndCells.AddRange(GetReferenceCellsFromFormulaIfRange(cell));
                }
                else
                {
                    referenceIndexesAndCells.AddRange(GetReferenceCellsFromFormulaIfNotRange(cell));
                }
            }

            return referenceIndexesAndCells;
        }

        /// <summary>
        /// Extracts the reference cells from a formula if the format is not a range (a1, b1, c1) as a List<(int sheetRefIndex, string cell)>.
        /// </summary>
        internal List<(int, string)> GetReferenceCellsFromFormulaIfNotRange(string formula)
        {
            List<(int referenceSheetIndex, string referenceStrippedCell)> referenceSheetCellAndIndexes = new List<(int, string)>();

            (int referenceIndex, _, string strippedCell) = CheckForOtherSheetReference(formula);

            referenceSheetCellAndIndexes.Add((referenceIndex, strippedCell));

            return referenceSheetCellAndIndexes;
        }

        /// <summary>
        /// Extracts the reference cells from a formula if the format is a range (a1:c1) as a List<(int sheetRefIndex, string cell)>.
        /// </summary>
        public List<(int, string)> GetReferenceCellsFromFormulaIfRange(string formula)
        {
            List<(int, string)> referenceSheetCellAndIndexes = new List<(int, string)>();

            (int referenceIndex, _, formula) = CheckForOtherSheetReference(formula);

            var rangeStartEndPoints = formula.Split(':', StringSplitOptions.RemoveEmptyEntries);

            if (GetCellReferencePositionIfValid(rangeStartEndPoints[0], out (int X, int Y) start)
                && GetCellReferencePositionIfValid(rangeStartEndPoints[1], out (int X, int Y) end))
            {
                for (int row = start.X; row <= end.X; row++)
                {
                    for (int col = start.Y; col <= end.Y; col++)
                    {
                        referenceSheetCellAndIndexes.Add(
                            (referenceIndex, sheetReference.Alphabet[col] + row.ToString())
                        );
                    }
                }
            }

            return referenceSheetCellAndIndexes;
        }

        /// <summary>
        /// Converts a string if it's in the correct reference format "=b1" to its (X,Y) coordinate format.
        /// </summary>
        internal bool GetCellReferencePositionIfValid(string cell, out (int X, int Y) cellPosition)
        {
            cellPosition = (-1, -1);

            if (cell.Length > 0 && cell[0] == '=')
            {
                cell = cell[1..].Trim();
            }

            var column = GetReferenceCellColumnLetterOrRowNumber(cell);
            var row = GetReferenceCellColumnLetterOrRowNumber(cell, rowNumber: true);

            bool isValid = false;
            int Y = -1;

            if (int.TryParse(row, out int X) && X > 0)
            {
                if (!string.IsNullOrEmpty(column))
                {
                    Y = Array.IndexOf(sheetReference.Alphabet.ToArray(), column.ToUpper());
                    isValid = Y != -1;
                }
            }

            if (isValid)
            {
                cellPosition = (X, Y);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Tries to return a reference value if a reference exists, otherwise it returns the direct value.
        /// </summary>
        private string GetReferenceCellValueOrDirectValue(
            (int referenceIndex, string cellOrValue) referenceIndexAndCellOrValue
        )
        {
            if (GetCellReferencePositionIfValid(
                referenceIndexAndCellOrValue.cellOrValue,
                out (int X, int Y) cellPosition
            ))
            {
                return sheetReference.sheets[referenceIndexAndCellOrValue.referenceIndex].sheet.cellData.ContainsKey(cellPosition)
                    ? sheetReference.sheets[referenceIndexAndCellOrValue.referenceIndex].sheet.cellData[cellPosition].Content.TypeValue.ToString()
                    : string.Empty;
            }

            return referenceIndexAndCellOrValue.cellOrValue;
        }

        /// <summary>
        /// Checks for a sheet reference.
        /// Returns the a tuple containing the sheet index where the cell can be found and the cell/range without the reference notation.
        /// </summary>
        internal (int, string, string) CheckForOtherSheetReference(string formula)
        {
            int referenceSheetIndex = sheetReference.currentSheetIndex;
            string referenceSheetName = sheetReference.currentSheetName;

            if (formula.Contains('!'))
            {
                formula = formula[0] == '(' && formula[^1] == ')' ? formula[1..^1] : formula;

                referenceSheetName = formula
                    .TakeWhile(c => c != '!')
                    .Aggregate("", (acc, cur) => acc += cur);

                if (sheetReference.sheets.Any(sheet => sheet.sheetName == referenceSheetName))
                {
                    formula = formula
                        .Skip(referenceSheetName.Length + 1)
                        .Aggregate("", (acc, cur) => acc += cur);

                    referenceSheetIndex = sheetReference.sheets.FindIndex(
                        sheet => sheet.sheetName == referenceSheetName
                    );
                }
                else
                {
                    referenceSheetName = string.Empty;
                }
            }

            return (referenceSheetIndex, referenceSheetName, formula);
        }

        /// <summary>
        /// Converts a string or cell value into it's proper type ("123.5" -> 123.5).
        /// </summary>
        private bool ConvertValueToProperType(
            (int index, string cell) referenceIndexAndCellOrNumber,
            out dynamic cellOrNumberValue
        )
        {
            if (GetCellReferencePositionIfValid(
                referenceIndexAndCellOrNumber.cell,
                out (int X, int Y) cellPosition
            )
                && sheetReference.sheets[referenceIndexAndCellOrNumber.index].sheet.cellData.ContainsKey(cellPosition)
                && sheetReference.sheets[referenceIndexAndCellOrNumber.index].sheet.cellData[cellPosition].Content.TypeValue is not string)
            {
                cellOrNumberValue = sheetReference.sheets[referenceIndexAndCellOrNumber.index]
                    .sheet
                    .cellData[cellPosition]
                    .Content
                    .TypeValue;
                return true;
            }
            else if (CreationHandler.FindValueType(referenceIndexAndCellOrNumber.cell).TypeValue is not string)
            {
                cellOrNumberValue = CreationHandler.FindValueType(referenceIndexAndCellOrNumber.cell).TypeValue;
                return true;
            }

            cellOrNumberValue = null;
            return false;
        }

        /// <summary>
        /// Tries to match if a formula contains a given name. Used to select the proper formula method based on the match.
        /// </summary>
        private bool MatchFormulaName((int X, int Y) position, string formulaToLookFor)
        {
            string formula = sheetReference.cellData[position].Content.TypeValue.ToString();
            string trimmedFormula = formula.Split('(')[0].ToUpper().Replace(" ", "");

            return trimmedFormula == formulaToLookFor;
        }

        /// <summary>
        /// Checks if a value is text by checking if there is a valid cell for that text. If not it checks if it has quotations.
        /// </summary>
        private bool CheckIfFormulaTextElementHasQuotations(
            (int index, string cell) refIndexAndCell,
            out string? newValue
        )
        {
            if (!GetCellReferencePositionIfValid(refIndexAndCell.cell, out _))
            {
                if (refIndexAndCell.cell.Length >= 2
                    && refIndexAndCell.cell[0] == '"'
                    && refIndexAndCell.cell[^1] == '"')
                {
                    newValue = refIndexAndCell.cell[1..^1];
                    return true;
                }
                else
                {
                    newValue = null;
                    return false;
                }
            }

            newValue = GetReferenceCellValueOrDirectValue(refIndexAndCell);
            return true;
        }

        /// <summary>
        /// Returns the column letters or the row number from a cell reference notation.
        /// </summary>
        internal string GetReferenceCellColumnLetterOrRowNumber(string cell, bool rowNumber = false)
        {
            if (rowNumber)
            {
                return cell.SkipWhile(c => char.IsLetter(c))
                    .Aggregate("", (acc, cur) => acc += cur);
            }

            return cell.TakeWhile(c => char.IsLetter(c)).Aggregate("", (acc, cur) => acc += cur);
        }
    }
}
