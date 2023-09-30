namespace Spreadsheet_Project
{
    public class PivotHandler
    {
        int rowFieldIndex = 1;
        int valueFieldsIndex = 1;
        int locationIndex = 1;

        Sheet sheetReference;

        Formulas.Formulas formulas;

        Dictionary<(int refIndex, string cell), List<(int refIndex, string cell)>> pivotTableCellReferenceData;

        Dictionary<(int refIndex, string value), List<(int refIndex, string cell)>> pivotTablPreFormulasData;

        Dictionary<string, List<string>> pivotTableFinal;

        private static List<List<string>> pivotSelectionMenu = new()
        {
            new() { "Range...", "" },
            new() { "Exit...", "" }
        };

        List<List<bool>> selectedFormulas;

        // Keeps track if a field was selected.
        List<bool> selected = Enumerable.Repeat(false, pivotSelectionMenu.Count).ToList();

        // EXPORT-----------------------------------------------------------------

        internal (int sheetRefIndex, string cellRange) dataRefIndexAndRange;

        internal (
            int sheetRefIndex,
            string startingCell,
            (int X, int Y) positionKey
        ) locationRefIndexStartingCellAndPosKey;

        internal List<(string headerName, int formula)> headersAndFormulas;

        internal List<(int sheetIndex, (int, int) positionKey)> pivotCellKeysToNotBeExported;

        //-----------------------------------------------------------------------

        public PivotHandler(Sheet sheetReference)
        {
            this.sheetReference = sheetReference;
            formulas = new Formulas.Formulas(sheetReference);
            pivotTableCellReferenceData = new Dictionary<(int, string), List<(int, string)>>();
            pivotTablPreFormulasData = new Dictionary<(int, string), List<(int, string)>>();
            pivotTableFinal = new Dictionary<string, List<string>>();
            pivotCellKeysToNotBeExported = new();
            dataRefIndexAndRange = (-1, string.Empty);
            locationRefIndexStartingCellAndPosKey = (-1, string.Empty, (-1, -1));
            headersAndFormulas = new List<(string headerName, int formula)>();
        }

        public void RunPivot()
        {
            Console.TreatControlCAsInput = false;
            Console.CursorVisible = false;
            sheetReference.pivotMenuActive = true;
            selected[^1] = false;
            GatherPivotInformation();
            if (!new int[] { valueFieldsIndex, locationIndex }.Contains(1))
            {
                GeneratePivotTable();
                AddGeneratedPivotTableToSheet();
                AssignPivotTableDataToSheet();
            }
            ClearOldPivotInformation();
            sheetReference.pivotMenuActive = false;
        }

        public void ImportPivotTable(
            string dataRange,
            string location,
            string selectedRow,
            IEnumerable<(int formula, string valueName)> formulaAndSelectedColumn
        )
        {
            InitializeAndPopulateVariablesFromImportInformation(dataRange, location, selectedRow, formulaAndSelectedColumn);

            if (!new int[] { valueFieldsIndex, locationIndex }.Contains(1))
            {
                GeneratePivotTable();
                AddGeneratedPivotTableToSheet();
                AssignPivotTableDataToSheet();
            }

            ClearOldPivotInformation();
        }

        /// <summary>
        /// Helper function that populates and initializez variables that would normally be done
        /// in the Pivot Creation process using the Pivot Menu.
        /// </summary>
        private void InitializeAndPopulateVariablesFromImportInformation(
            string dataRange,
            string location,
            string selectedRow,
            IEnumerable<(int formula, string valueName)> formulaAndSelectedColumn
        )
        {
            pivotSelectionMenu[0][1] = dataRange;

            PopulatePivotTableCellReferenceTableBasedOnInputtedRange();
            InitializePivotSelectionMenuWithNewRowValuesLocationFields();

            pivotSelectionMenu[locationIndex][1] = location;

            selected = Enumerable.Repeat(false, pivotSelectionMenu.Count).ToList();
            selected[pivotSelectionMenu.FindIndex(x => x[1] == selectedRow)] = true;

            foreach (var (formula, valueName) in formulaAndSelectedColumn)
            {
                int formulaIndex = pivotSelectionMenu
                    .GetRange(valueFieldsIndex + 1, locationIndex - valueFieldsIndex)
                    .FindIndex(x => x[1] == valueName);

                int valueIndex = formulaIndex + valueFieldsIndex + 1;

                selected[valueIndex] = true;
                selectedFormulas[formulaIndex][formula] = true;
            }
        }

        //PIVOT TABLE INFORMATION COLLECTION-----DOWN-------------------------------------------------------------------------------

        /// <summary>
        /// Prints the Pivot Menu and lets you input and select the required information in order to create the table.
        /// </summary>
        private void GatherPivotInformation()
        {
            (int row, int col) selectedMenuRow = (0, 0); // Keeps track of the current selected row in the pivot menu.

            while (!selected[^1])
            {
                ConsoleKeyInfo keyInfo = default;

                PrintPivotMenu(selectedMenuRow, selected);

                if (selected[0]) // Range Field is always on the first position (0).
                {
                    (selected, keyInfo) = EditRangeField(selected, keyInfo);
                }

                if (selected[locationIndex] && pivotSelectionMenu.Count > 0)
                {
                    (selected, keyInfo) = EditLocationField(selected, keyInfo);
                }

                // GENERAL MENU KEY READER - If key is Up/Down skip ReadKey.
                keyInfo =
                    !selected[0]
                    && !selected[locationIndex]
                    && !new ConsoleKey[] { ConsoleKey.DownArrow, ConsoleKey.UpArrow }.Contains(
                        keyInfo.Key
                    )
                        ? Console.ReadKey(true)
                        : keyInfo;

                selectedMenuRow = ExecutePivotMenuAction(keyInfo, selectedMenuRow);
            }
        }

        /// <summary>
        /// Executes the menu available options such as selecting and navigating between rows,
        /// values, formulas, etc.
        /// </summary>
        private (int, int) ExecutePivotMenuAction(
            ConsoleKeyInfo keyInfo,
            (int row, int col) selectedMenuRow
        )
        {
            switch (keyInfo.Key)
            {
                case ConsoleKey.UpArrow:
                    selectedMenuRow = GetNewSelectedMenuRow(
                        selectedMenuRow.row,
                        -1,
                        pivotSelectionMenu.Count - 1
                    );
                    break;

                case ConsoleKey.DownArrow:
                    selectedMenuRow = GetNewSelectedMenuRow(
                        selectedMenuRow.row,
                        1,
                        pivotSelectionMenu.Count - 1
                    );
                    break;

                case ConsoleKey.RightArrow:
                    if (
                        (
                            selectedMenuRow.row > valueFieldsIndex
                            && selectedMenuRow.row < locationIndex
                        )
                        && selectedMenuRow.col + 2 < pivotSelectionMenu[selectedMenuRow.row].Count
                    )
                    {
                        selectedMenuRow = (selectedMenuRow.row, selectedMenuRow.col + 2);
                    }
                    break;

                case ConsoleKey.LeftArrow:
                    if (
                        (
                            selectedMenuRow.row > valueFieldsIndex
                            && selectedMenuRow.row < locationIndex
                        )
                        && selectedMenuRow.col - 2 >= 0
                    )
                    {
                        selectedMenuRow = (selectedMenuRow.row, selectedMenuRow.col - 2);
                    }
                    break;

                case ConsoleKey.Spacebar:
                    selected[selectedMenuRow.row] = !selected[selectedMenuRow.row];
                    MarkSelectedRowsAndValues(selectedMenuRow, selected);
                    break;

                default:
                    break;
            }

            return selectedMenuRow;
        }

        /// <summary>
        /// Helper function used to skip over the menu notation fields (Row Fields, Values Fields).
        /// Also skips the value field that matches the selected row field.
        /// </summary>
        private (int, int) GetNewSelectedMenuRow(int currentRow, int direction, int maxRow)
        {
            int newRow = currentRow + direction;
            int matchingRowValueFieldIndex = 1;

            int selectedRowIndex = selected.FindIndex(
                rowFieldIndex + 1,
                valueFieldsIndex - rowFieldIndex + 1,
                selected => selected == true
            );

            if (selectedRowIndex >= 0)
            {
                string matchingRowValueField = pivotSelectionMenu[selectedRowIndex][1];

                matchingRowValueFieldIndex = pivotSelectionMenu.FindIndex(
                    valueFieldsIndex + 1,
                    locationIndex - valueFieldsIndex + 1,
                    value => value[1] == matchingRowValueField
                );
            }

            while (
                new int[] { rowFieldIndex, valueFieldsIndex, matchingRowValueFieldIndex }.Contains(
                    newRow
                )
            )
            {
                newRow += direction;
            }

            return (Math.Max(0, Math.Min(newRow, maxRow)), 0);
        }

        private void PrintPivotMenu((int row, int col) selectedRow, List<bool> selected)
        {
            Console.Clear();

            Console.WriteLine("Navigate with arrow keys and press space to select fields or edit them:\n");

            // Print the menu items
            for (int row = 0; row < pivotSelectionMenu.Count; row++)
            {
                for (int col = 0; col < pivotSelectionMenu[row].Count; col++)
                {
                    if ((row, col) == selectedRow)
                    {
                        Console.BackgroundColor = ConsoleColor.White;
                        Console.ForegroundColor = ConsoleColor.Black;
                    }

                    if (new int[] { rowFieldIndex, valueFieldsIndex, locationIndex, locationIndex + 1 }.Contains(row))
                    {
                        Console.WriteLine();
                    }

                    Console.Write($" {pivotSelectionMenu[row][col]} ");

                    Console.ResetColor();
                }

                Console.WriteLine();
            }
        }

        /// <summary>
        /// Allows for Range Field modification if selected and once it's deselected it populates pivotTableCellReferenceData
        /// as well as adding row, values, location fields to the pivot menu based on the range selected.
        /// </summary>
        private (List<bool>, ConsoleKeyInfo) EditRangeField(List<bool> selected, ConsoleKeyInfo keyInfo)
        {
            (pivotSelectionMenu[0][1], selected[0], _) = GatherRangeOrLocationInformation(pivotSelectionMenu[0][1]);

            // If the select bool is false it means the down key was pressed and this
            // allows the user to not be forced to press it again in order to navigate the pivot menu.
            keyInfo = !selected[0]
                ? new ConsoleKeyInfo((char)0, ConsoleKey.DownArrow, false, false, false)
                : keyInfo;

            if (!selected[0])
            {
                try
                {
                    PopulatePivotTableCellReferenceTableBasedOnInputtedRange();
                }
                catch (Exception)
                {
                    TableRangeErrorPrompt();
                    RunPivot();
                }

                InitializePivotSelectionMenuWithNewRowValuesLocationFields();

                selected = Enumerable.Repeat(false, pivotSelectionMenu.Count).ToList();
            }

            return (selected, keyInfo);
        }

        /// <summary>
        /// Populates pivotTableCellReferenceData with cell reference data based on the range selected.
        /// </summary>
        private void PopulatePivotTableCellReferenceTableBasedOnInputtedRange()
        {
            List<(int referenceIndex, string cellOrValue)> referenceIndexAndCells =
                formulas.GetReferenceCellsFromFormulaIfRange(pivotSelectionMenu[0][1]);

            string lastColumnLetter = formulas
                .GetReferenceCellColumnLetterOrRowNumber(referenceIndexAndCells[^1].cellOrValue)
                .ToUpper();

            var pivotTableHeaders = referenceIndexAndCells.GetRange(0, sheetReference.Alphabet.IndexOf(lastColumnLetter));

            foreach (var header in pivotTableHeaders)
            {
                var headerValues = referenceIndexAndCells.Where(indexAndCell =>
                    formulas.GetReferenceCellColumnLetterOrRowNumber(indexAndCell.cellOrValue)
                    .Equals(formulas.GetReferenceCellColumnLetterOrRowNumber(header.cellOrValue)));

                pivotTableCellReferenceData.Add(header, headerValues.Skip(1).ToList());
            }

            dataRefIndexAndRange = (
                referenceIndexAndCells[0].referenceIndex,
                pivotSelectionMenu[0][1].Substring(pivotSelectionMenu[0][1].IndexOf('!') + 1)
            );
        }

        /// <summary>
        /// Populates the pivotSelectionMenu with Row/Values/Location fields based on the range selected.
        /// </summary>
        private void InitializePivotSelectionMenuWithNewRowValuesLocationFields()
        {
            pivotSelectionMenu = new List<List<string>> { pivotSelectionMenu[0] };

            foreach (var fieldCategory in new string[] { "Row Field", "Value Fields" })
            {
                pivotSelectionMenu.Add(new List<string>() { fieldCategory, "" });

                foreach (var kvp in pivotTableCellReferenceData)
                {
                    formulas.GetCellReferencePositionIfValid(
                        kvp.Key.cell,
                        out (int X, int Y) cellPosition
                    );

                    if (
                        cellPosition != (-1, -1)
                        && sheetReference.sheets[kvp.Key.refIndex].sheet.cellData.ContainsKey(
                            cellPosition
                        )
                    )
                    {
                        pivotSelectionMenu.Add(
                            new List<string>()
                            {
                                "[ ]",
                                sheetReference.sheets[kvp.Key.refIndex].sheet.cellData[
                                    cellPosition
                                ].Content.TypeValue.ToString()
                            }
                        );
                    }
                }
            }

            pivotSelectionMenu.Add(new List<string>() { "Location...", "" });

            pivotSelectionMenu.Add(new List<string>() { "Exit...", "" });

            InitializeMenuIndexesAndSelectedValueFieldsFormulas();
        }

        /// <summary>
        /// Saves the Menu Indexes for notations.
        /// </summary>
        private void InitializeMenuIndexesAndSelectedValueFieldsFormulas()
        {
            rowFieldIndex = pivotSelectionMenu.FindIndex(
                innerList => innerList.SequenceEqual(new List<string> { "Row Field", "" })
            );
            valueFieldsIndex = pivotSelectionMenu.FindIndex(
                innerList => innerList.SequenceEqual(new List<string> { "Value Fields", "" })
            );
            locationIndex = pivotSelectionMenu.FindIndex(
                innerList => innerList.SequenceEqual(new List<string> { "Location...", "" })
            );

            selectedFormulas = Enumerable
                .Range(0, locationIndex - valueFieldsIndex - 1)
                .Select(_ => Enumerable.Repeat(false, 3).ToList())
                .ToList();
        }

        /// <summary>
        /// Allows for Location Field modification if selected.
        /// </summary>
        private (List<bool>, ConsoleKeyInfo) EditLocationField(
            List<bool> selected,
            ConsoleKeyInfo keyInfo
        )
        {
            string? exitMovement;

            (pivotSelectionMenu[locationIndex][1], selected[locationIndex], exitMovement) =
                GatherRangeOrLocationInformation(
                    pivotSelectionMenu[locationIndex][1],
                    location: true
                );

            // If the select bool is false it means the up key was pressed and this
            // allows the user to not be forced to press it again in order to navigate the menu.
            if (!selected[locationIndex])
            {
                keyInfo =
                    exitMovement == "up"
                        ? new ConsoleKeyInfo((char)0, ConsoleKey.UpArrow, false, false, false)
                        : new ConsoleKeyInfo((char)0, ConsoleKey.DownArrow, false, false, false);
            }

            return (selected, keyInfo);
        }

        /// <summary>
        /// Helper function that helps with the Range or Location input modification. Appends chars, backspace, delete.
        /// </summary>
        private (string, bool, string) GatherRangeOrLocationInformation(
            string rangeOrLocationField,
            bool location = false
        )
        {
            bool locationStillSelected = true;
            string selectionExitMovement = string.Empty;

            ConsoleKeyInfo keyInfo = Console.ReadKey(true);

            switch (keyInfo.Key)
            {
                case ConsoleKey.Delete:
                    rangeOrLocationField = string.Empty;
                    break;
                case ConsoleKey.Backspace:
                    rangeOrLocationField =
                        rangeOrLocationField.Length > 0
                            ? rangeOrLocationField[..^1]
                            : rangeOrLocationField;
                    break;
                case ConsoleKey.UpArrow:
                    locationStillSelected = !location;
                    selectionExitMovement = "up";
                    break;
                case ConsoleKey.DownArrow:
                    locationStillSelected = false;
                    selectionExitMovement = "down";
                    break;
                default:
                    rangeOrLocationField += keyInfo.KeyChar.ToString();
                    break;
            }

            return (rangeOrLocationField, locationStillSelected, selectionExitMovement);
        }

        /// <summary>
        /// Allows Rows and Value Formula selection by modifying the visual aspect of the brackets "[ ] -> [x]" and
        /// marking the position with "true" in the selected and selectedFormulas bool lists.
        /// </summary>
        private void MarkSelectedRowsAndValues(
            (int row, int col) selectedRowCol,
            List<bool> selected
        )
        {
            switch (selectedRowCol.row)
            {
                case int row when row > rowFieldIndex && row < valueFieldsIndex:
                    MarkSelectedRow(selectedRowCol, selected);
                    break;

                case int r when r > valueFieldsIndex && r < locationIndex:
                    if (selectedRowCol.col == 0)
                    {
                        MarkSelectedValue(selectedRowCol);
                    }
                    else
                    {
                        MarkSelectedValueFormula(selectedRowCol);
                    }
                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Helper function for MarkSelectedRowsAndValues.
        /// </summary>
        private void MarkSelectedRow((int row, int col) selectedRowCol, List<bool> selected)
        {
            if (selected
                .GetRange(rowFieldIndex + 1, valueFieldsIndex - rowFieldIndex - 1)
                .Count(x => x is true) <= 1)
            {
                pivotSelectionMenu[selectedRowCol.row][0] = selected[selectedRowCol.row]
                    ? "[x]"
                    : "[ ]";
            }
            else
            {
                for (int i = rowFieldIndex + 1; i < valueFieldsIndex; i++)
                {
                    if (pivotSelectionMenu[i][0] == "[x]")
                    {
                        pivotSelectionMenu[i][0] = "[ ]";
                        selected[i] = false;
                    }
                }

                pivotSelectionMenu[selectedRowCol.row][0] = "[x]";
            }
        }

        /// <summary>
        /// Helper function for MarkSelectedRowsAndValues.
        /// </summary>
        private void MarkSelectedValue((int row, int col) selectedRowCol)
        {
            pivotSelectionMenu[selectedRowCol.row][0] =
                pivotSelectionMenu[selectedRowCol.row][0] != "[x]" ? "[x]" : "[ ]";

            pivotSelectionMenu[selectedRowCol.row].RemoveRange(
                2,
                pivotSelectionMenu[selectedRowCol.row].Count - 2
            );

            if (pivotSelectionMenu[selectedRowCol.row][0] == "[x]")
            {
                pivotSelectionMenu[selectedRowCol.row].AddRange(
                    new List<string>() { "[x]", "SUM", "[ ]", "AVG", "[ ]", "COUNT" }
                );
                selectedFormulas[selectedFormulas.Count - (locationIndex - selectedRowCol.row)][0] =
                    true;
            }
        }


        /// <summary>
        /// Helper function for MarkSelectedRowsAndValues.
        /// </summary>
        private void MarkSelectedValueFormula((int row, int col) selectedRowCol)
        {
            selected[selectedRowCol.row] = !selected[selectedRowCol.row];

            var correspondingSelectedFormulasCol =
                selectedRowCol.col == 0 ? 0 : selectedRowCol.col / 2 - 1;

            for (int i = 2; i < pivotSelectionMenu[selectedRowCol.row].Count; i++)
            {
                pivotSelectionMenu[selectedRowCol.row][i] =
                    pivotSelectionMenu[selectedRowCol.row][i] == "[x]"
                        ? "[ ]"
                        : pivotSelectionMenu[selectedRowCol.row][i];
            }

            selectedFormulas[selectedFormulas.Count - (locationIndex - selectedRowCol.row)] =
                Enumerable.Repeat(false, 3).ToList();

            pivotSelectionMenu[selectedRowCol.row][selectedRowCol.col] =
                pivotSelectionMenu[selectedRowCol.row][selectedRowCol.col] != "[x]" ? "[x]" : "[ ]";

            selectedFormulas[selectedFormulas.Count - (locationIndex - selectedRowCol.row)][
                correspondingSelectedFormulasCol
            ] = !selectedFormulas[selectedFormulas.Count - (locationIndex - selectedRowCol.row)][
                correspondingSelectedFormulasCol
            ];
        }

        //PIVOT TABLE INFORMATION COLLECTION-----UP---------------------------------------------------------------------------------

        //PIVOT TABLE CREATION-----DOWN---------------------------------------------------------------------------------------------
        private void GeneratePivotTable()
        {
            (int referenceIndex, string referenceStrippedCell) referenceIndexAndCell =
                formulas.GetReferenceCellsFromFormulaIfNotRange(
                    pivotSelectionMenu[locationIndex][1]
                )[0];

            formulas.GetCellReferencePositionIfValid(
                referenceIndexAndCell.referenceStrippedCell,
                out (int X, int Y) cellPosition
            );

            locationRefIndexStartingCellAndPosKey = (
                referenceIndexAndCell.referenceIndex,
                referenceIndexAndCell.referenceStrippedCell,
                cellPosition
            );

            DeletePreviousPivotOnPosition(
                locationRefIndexStartingCellAndPosKey.sheetRefIndex,
                locationRefIndexStartingCellAndPosKey.positionKey
            );

            CreatePreFormulaAndFinalPivotTable();
        }

        /// <summary>
        /// Checks if the sheet (sheetReferenceIndex) holds any pivotTables on the positionKey position and if so it deletes it.
        /// </summary>
        internal void DeletePreviousPivotOnPosition(int sheetReferenceIndex, (int X, int Y) positionKey)
        {
            if (sheetReference.sheets[sheetReferenceIndex].sheet.PivotTableData.Count > 0)
            {
                List<(
                    (int sheetRefIndex, string cellRange) dataRefIndexAndRange,
                    (
                        int sheetRefIndex,
                        string startingCell,
                        (int X, int Y) positionKey
                    ) locationRefIndexAndStartingCell,
                    List<(string headerName, int sheetRefIndex)> headersAndFormulas,
                    List<(int sheetIndex, (int, int) positionKey)> pivotCellKeysToNotBeExported
                )> deletedPivots = new();

                foreach (var pivotData in sheetReference.sheets[sheetReferenceIndex].sheet.PivotTableData)
                {
                    if (pivotData.pivotCellKeysToNotBeExported.Contains((sheetReferenceIndex, positionKey)))
                    {
                        foreach (var (_, key) in pivotData.pivotCellKeysToNotBeExported)
                        {
                            sheetReference.sheets[sheetReferenceIndex].sheet.cellData.Remove(key);
                        }

                        deletedPivots.Add(pivotData);
                    }
                }

                deletedPivots.ForEach(pivot =>
                    sheetReference.sheets[sheetReferenceIndex].sheet.PivotTableData.Remove(pivot)
                );
            }
        }

        /// <summary>
        /// Populates pivotTableFinal with the selected Row and Values values after specific formula application. It also adds a final row that
        /// consists of the column totals.
        /// </summary>
        private void CreatePreFormulaAndFinalPivotTable()
        {
            CreatePreFormulaPivotTable();

            int selectedRowIndex = selected.FindIndex(
                rowFieldIndex + 1,
                valueFieldsIndex - rowFieldIndex + 1,
                selected => selected == true
            );

            List<int> selectedValuesIndexes = Enumerable
                .Range(valueFieldsIndex + 1, locationIndex - valueFieldsIndex - 1)
                .Where(i => selected[i])
                .Select(i => i)
                .ToList();

            PopulateRowsFinalPivotTable(selectedRowIndex);

            PopulateValuesFinalPivotTable(selectedRowIndex, selectedValuesIndexes);

            PopulateTotalsFinalPivotTable(selectedValuesIndexes);
        }

        /// <summary>
        /// Populates pivotTablPreFormulasData only with the Row and Values selected. The Row values are as in the
        /// sheet so later a .Distinct() method can be applied.
        /// </summary>
        private void CreatePreFormulaPivotTable()
        {
            foreach (var kvp in pivotTableCellReferenceData)
            {
                formulas.GetCellReferencePositionIfValid(kvp.Key.cell, out (int X, int Y) cellPosition);

                PopulatePreFormulaTableDataRow(kvp, cellPosition);
            }

            foreach (var kvp in pivotTableCellReferenceData)
            {
                formulas.GetCellReferencePositionIfValid(kvp.Key.cell, out (int X, int Y) cellPosition);

                PopulatePreFormulaTableDataValues(kvp, cellPosition);
            }
        }

        /// <summary>
        /// Helper function that populates pivotTablPreFormulasData with the selected Row column.
        /// </summary>
        private void PopulatePreFormulaTableDataRow(
            KeyValuePair<(int refIndex, string cell), List<(int refIndex, string cell)>> kvp,
            (int X, int Y) cellPosition
        )
        {
            int selectedRowIndex = selected.FindIndex(
                rowFieldIndex + 1,
                valueFieldsIndex - rowFieldIndex + 1,
                selected => selected == true
            );

            if (sheetReference.sheets[kvp.Key.refIndex].sheet.cellData[cellPosition].Content.TypeValue.ToString() == pivotSelectionMenu[selectedRowIndex][1])
            {
                headersAndFormulas.Add((pivotSelectionMenu[selectedRowIndex][1], -1));

                List<(int refIndex, string value)> rowValues = new();
                foreach (var item in pivotTableCellReferenceData[kvp.Key])
                {
                    formulas.GetCellReferencePositionIfValid(item.cell, out (int Y, int Z) itemCellPosition);

                    rowValues.Add(
                        (
                            item.refIndex,
                            sheetReference.sheets[item.refIndex].sheet.cellData[itemCellPosition].Content.TypeValue.ToString()
                        )
                    );
                }

                pivotTablPreFormulasData.Add(
                    (
                        kvp.Key.refIndex,
                        sheetReference.sheets[kvp.Key.refIndex].sheet.cellData[cellPosition].Content.TypeValue.ToString()
                    ),
                    rowValues
                );
            }
        }

        /// <summary>
        /// Helper function that populates pivotTablPreFormulasData with the selected Values columns.
        /// </summary>
        private void PopulatePreFormulaTableDataValues(
            KeyValuePair<(int refIndex, string cell), List<(int refIndex, string cell)>> kvp,
            (int X, int Y) cellPosition
        )
        {
            List<string> selectedValueFields = Enumerable
                .Range(valueFieldsIndex + 1, locationIndex - valueFieldsIndex)
                .Where(i => selected[i])
                .Select(i => pivotSelectionMenu[i][1])
                .ToList();

            if (selectedValueFields.Any(x => x == sheetReference.sheets[kvp.Key.refIndex].sheet.cellData[cellPosition].Content.TypeValue.ToString()))
            {
                pivotTablPreFormulasData.Add(
                    (
                        kvp.Key.refIndex,
                        sheetReference.sheets[kvp.Key.refIndex].sheet.cellData[cellPosition].Content.TypeValue.ToString()
                    ),
                    pivotTableCellReferenceData[kvp.Key]
                );

                if (headersAndFormulas.Count == 1)
                {
                    selectedValueFields.ForEach(headerName => headersAndFormulas.Add((headerName, -1)));
                }
            }
        }

        /// <summary>
        /// Helper function that populates pivotTableFinal with the Row values. The values are distinct and sorted.
        /// </summary>
        private void PopulateRowsFinalPivotTable(int selectedRowIndex)
        {
            pivotTableFinal.Add(
                pivotSelectionMenu[selectedRowIndex][1],
                pivotTablPreFormulasData
                    .ElementAt(0)
                    .Value.Select(refIndexAndValue => refIndexAndValue.cell)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList()
            );

            for (int i = 1; i < pivotTablPreFormulasData.Count; i++)
            {
                pivotTableFinal.Add(
                    pivotTablPreFormulasData.ElementAt(i).Key.value,
                    new List<string>()
                );
            }
        }

        /// <summary>
        /// Helper function that populates pivotTableFinal with the Value values. The values are transformed based on the selected formula.
        /// </summary>
        private void PopulateValuesFinalPivotTable(int selectedRowIndex, List<int> selectedValuesIndexes)
        {
            foreach (var row in pivotTableFinal[pivotSelectionMenu[selectedRowIndex][1]])
            {
                for (int i = 1; i < pivotTablPreFormulasData.Count; i++)
                {
                    var elementsReadyForFormulaApplication = pivotTablPreFormulasData
                        .ElementAt(i)
                        .Value
                        .Where(x =>
                            pivotTablPreFormulasData.ElementAt(0).Value[
                                pivotTablPreFormulasData.ElementAt(i).Value.IndexOf(x)
                            ].cell == row
                        )
                        .ToList();

                    formulas.GetReferenceCellsOrValuesSumAndCount(
                        elementsReadyForFormulaApplication,
                        out dynamic sum,
                        out int count
                    );

                    switch (selectedFormulas[
                        selectedFormulas.Count - (locationIndex - selectedValuesIndexes[i - 1])
                    ].FindIndex(bl => bl == true))
                    {
                        case 0:
                            pivotTableFinal.ElementAt(i).Value.Add(sum.ToString());
                            headersAndFormulas[i] = (headersAndFormulas[i].headerName, 0);
                            break;

                        case 1:
                            pivotTableFinal.ElementAt(i).Value.Add((sum / count).ToString());
                            headersAndFormulas[i] = (headersAndFormulas[i].headerName, 1);
                            break;

                        case 2:
                            pivotTableFinal.ElementAt(i).Value.Add(count.ToString());
                            headersAndFormulas[i] = (headersAndFormulas[i].headerName, 2);
                            break;

                        default:
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Helper function that populates pivotTableFinal with a bottom row that holds the grand total of each column.
        /// </summary>
        private void PopulateTotalsFinalPivotTable(List<int> selectedValuesIndexes)
        {
            int valueColCount = 0;

            int selectedRowIndex = selected.FindIndex(
                rowFieldIndex + 1,
                valueFieldsIndex - rowFieldIndex + 1,
                selected => selected == true
            );

            foreach (var kvp in pivotTableFinal)
            {
                if (kvp.Key == pivotSelectionMenu[selectedRowIndex][1])
                {
                    pivotTableFinal[kvp.Key].Add("Total");
                }
                else
                {
                    switch (selectedFormulas[
                        selectedFormulas.Count - (locationIndex - selectedValuesIndexes[valueColCount])
                    ].FindIndex(formulaSelected => formulaSelected == true))
                    {
                        case 1:
                            pivotTableFinal[kvp.Key].Add(
                                (pivotTableFinal[kvp.Key].Sum(x => ParseValue(x))
                                    / pivotTableFinal[kvp.Key].Count).ToString()
                            );
                            break;

                        default:
                            pivotTableFinal[kvp.Key].Add(
                                pivotTableFinal[kvp.Key].Sum(x => ParseValue(x)).ToString()
                            );
                            break;
                    }

                    valueColCount++;
                }
            }
        }

        private void AddGeneratedPivotTableToSheet()
        {
            if (!CheckSheetPrintSpaceAvailability())
            {
                return;
            }

            (int referenceIndex, string referenceStrippedCell) referenceIndexAndCell =
                formulas.GetReferenceCellsFromFormulaIfNotRange(
                    pivotSelectionMenu[locationIndex][1]
                )[0];

            formulas.GetCellReferencePositionIfValid(
                referenceIndexAndCell.referenceStrippedCell,
                out (int X, int Y) headerPosition
            );

            for (int i = 0; i < pivotTableFinal.Count(); i++)
            {
                sheetReference.sheets[referenceIndexAndCell.referenceIndex].sheet.cellData.Add(
                    (headerPosition.X, headerPosition.Y + i),
                    (
                        CreationHandler.FindValueType(pivotTableFinal.ElementAt(i).Key),
                        null
                    )
                );
            }

            (int row, int col) valuesPosition = (headerPosition.X + 1, headerPosition.Y);

            for (int i = 0; i < pivotTableFinal.Count(); i++)
            {
                for (int j = 0; j < pivotTableFinal.ElementAt(i).Value.Count(); j++)
                {
                    sheetReference.sheets[referenceIndexAndCell.referenceIndex].sheet.cellData.Add(
                        (valuesPosition.row + j, valuesPosition.col + i),
                        (
                            CreationHandler.FindValueType(pivotTableFinal.ElementAt(i).Value.ElementAt(j)),
                            null
                        )
                    );
                }
            }
        }

        /// <summary>
        /// Iterates through all the to-be pivot table cells and makes sure that the cellData dictionary
        /// does not contain any of those key positions.
        /// </summary>
        private bool CheckSheetPrintSpaceAvailability()
        {
            for (int i = 0; i < pivotTableFinal.Count(); i++)
            {
                for (int j = 0; j < pivotTableFinal.ElementAt(i).Value.Count() + 1; j++)
                {
                    if (sheetReference.sheets[locationRefIndexStartingCellAndPosKey.sheetRefIndex].sheet.cellData.ContainsKey(
                        (
                            locationRefIndexStartingCellAndPosKey.positionKey.X + j,
                            locationRefIndexStartingCellAndPosKey.positionKey.Y + i
                        )
                    ))
                    {
                        if(!sheetReference.isTesting)
                        {
                            LocationNotEmptyErrorPrompt();
                            pivotCellKeysToNotBeExported.Clear();
                        }
                        return false;
                    }

                    pivotCellKeysToNotBeExported.Add(
                        (
                            locationRefIndexStartingCellAndPosKey.sheetRefIndex,
                            (
                                locationRefIndexStartingCellAndPosKey.positionKey.X + j,
                                locationRefIndexStartingCellAndPosKey.positionKey.Y + i
                            )
                        )
                    );
                }
            }

            return true;
        }

        /// <summary>
        /// Add the Pivot Table data to the sheet that holds it.
        /// Data is only added if Pivot Creation is successful.
        /// </summary>
        private void AssignPivotTableDataToSheet()
        {
            (int sheetRefIndex, string cellRange) dataReferenceIndexAndRangeExit =
                (dataRefIndexAndRange.sheetRefIndex, dataRefIndexAndRange.cellRange);

            (
                int sheetRefIndex,
                string startingCell,
                (int X, int Y) positionKey
            ) locationReferenceIndexStartingCellAndPosKeyExit = (
                locationRefIndexStartingCellAndPosKey.sheetRefIndex,
                locationRefIndexStartingCellAndPosKey.startingCell,
                locationRefIndexStartingCellAndPosKey.positionKey
            );

            List<(string headerName, int formula)> headersAndFormulasExit = new();
            headersAndFormulasExit.AddRange(headersAndFormulas);

            List<(int sheetIndex, (int, int) positionKey)> keysToNotBeExportedExit = new();
            keysToNotBeExportedExit.AddRange(pivotCellKeysToNotBeExported);

            sheetReference.sheets[locationRefIndexStartingCellAndPosKey.sheetRefIndex].sheet.PivotTableData.Add(
                (
                    dataReferenceIndexAndRangeExit,
                    locationReferenceIndexStartingCellAndPosKeyExit,
                    headersAndFormulasExit,
                    keysToNotBeExportedExit
                )
            );
        }

        /// <summary>
        /// Clears all the initialized variables to avoid conflicts for the next Pivot Table Creation.
        /// </summary>
        private void ClearOldPivotInformation()
        {
            pivotSelectionMenu = new()
            {
                new() { "Range...", "" },
                new() { "Exit...", "" }
            };

            selected = Enumerable.Repeat(false, pivotSelectionMenu.Count).ToList();

            List<List<bool>> selectedFormulas = new();

            pivotTableCellReferenceData.Clear();
            pivotTablPreFormulasData.Clear();
            pivotTableFinal.Clear();

            headersAndFormulas.Clear();
            dataRefIndexAndRange = (-1, string.Empty);
            locationRefIndexStartingCellAndPosKey = (-1, string.Empty, (-1, -1));
            pivotCellKeysToNotBeExported.Clear();

            rowFieldIndex = 1;
            valueFieldsIndex = 1;
            locationIndex = 1;
        }

        /// <summary>
        /// Helper method used in PopulateTotalFinalPivotTable used to parse grand total values.
        /// </summary>
        private decimal ParseValue(string value)
        {
            if (int.TryParse(value, out int intValue))
            {
                return (decimal)intValue;
            }
            else if (double.TryParse(value, out double doubleValue))
            {
                return (decimal)doubleValue;
            }
            else if (decimal.TryParse(value, out decimal decimalValue))
            {
                return decimalValue;
            }
            else
            {
                return 0M; // Return 0 if the value cannot be parsed as a numeric type
            }
        }

        /// <summary>
        /// Error prompt for when the inputted range is not valid.
        /// </summary>
        private void TableRangeErrorPrompt()
        {
            Console.Clear();
            Console.WriteLine(
                "-Please check if the table range and table location is correctly entered."
                    + "\n Press enter to return to the MENU."
            );
            Console.ReadLine();
        }

        /// <summary>
        /// Error prompt for when the space needed for the pivot is occupied.
        /// </summary>
        private void LocationNotEmptyErrorPrompt()
        {
            Console.Clear();
            Console.WriteLine(
                "-Please make sure that there's enough space for the desired pivot to be printed."
                    + "\n Press enter to return to the Sheet."
            );
            Console.ReadLine();
        }

    }
}
