using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using Spreadsheet_Project.Formulas;
using Spreadsheet_Project.Network;
using System.Net.Sockets;

namespace Spreadsheet_Project
{
    public class ApplicationInitializer
    {
        internal List<(string sheetName, Sheet sheet)> sheets = new();
        internal  string? currentSheetName;
        internal  int currentSheetIndex;
        private  int lastWidth = 120;
        private  int lastHeight = 30;
        private bool serverInProgress = false;
        private bool clientInProgress = false;

        Client ?clientObject;
        Server ?serverObject;

        public static void Main()
        {
            ApplicationInitializer? initializer = new();

            // 100 milliseconds timer that runs the CheckConsoleSize callback delegate.
            var timer = new Timer(initializer.CheckConsoleSize, null, 0, 100);
            initializer.RunNewSheet();
        }

        /// <summary>
        /// Checks if the app window was changed by the user and updates the number of rows/columns to be printed.
        /// </summary>
        private void CheckConsoleSize(object state)
        {
            if (Console.BufferWidth != lastWidth || Console.BufferHeight != lastHeight
                && !sheets[currentSheetIndex].sheet.pivotMenuActive)
            {
                lastWidth = Console.BufferWidth;
                lastHeight = Console.BufferHeight;
                sheets[currentSheetIndex].sheet.printHandler.UpdateConsoleSize();
                sheets[currentSheetIndex].sheet.printHandler.PrintSheet();
            }
        }

        private void RunNewSheet()
        {
            sheets.Add((GenerateNewSheetName(), new Sheet()));

            UpdateSheetReferences(currentSheetName);

            PassImportExportDelegates();

            PassMultipleSheetsDelegates();

            sheets[currentSheetIndex].sheet.printHandler.PrintSheet();

            RegisterAndExecuteKeyActions();
        }

        private void RegisterAndExecuteKeyActions()
        {
            ConsoleKeyInfo pressedKey;
            do
            {
                Console.TreatControlCAsInput = true;

                pressedKey = Console.ReadKey(true);

                switch (pressedKey.Key)
                {
                    case ConsoleKey.F9 when !pressedKey.Modifiers.HasFlag(ConsoleModifiers.Shift | ConsoleModifiers.Control | ConsoleModifiers.Alt):
                        ServerInitializer();
                        break;
                    case ConsoleKey.F10 when !pressedKey.Modifiers.HasFlag(ConsoleModifiers.Shift | ConsoleModifiers.Control | ConsoleModifiers.Alt):
                        ClientInitializer();
                        break;
                    default:
                        if (clientInProgress)
                        {
                            clientObject?.HandleKeyPressFromClientLocalConsoleAsync(pressedKey);
                        }
                        else if (serverInProgress)
                        {
                            serverObject?.HandleKeyPressFromServerLocalConsoleAsync(pressedKey);
                        }
                        else
                        {
                            sheets[currentSheetIndex].sheet.Execute(pressedKey);
                        }
                        break;
                }
                
                sheets[currentSheetIndex].sheet.printHandler.PrintSheet();

            } while (pressedKey.Key != ConsoleKey.Escape);
        }

        private async void ServerInitializer()
        {
            if (serverInProgress)
            {
                //shutdown logic

                serverInProgress = false;
            }
            else if(!clientInProgress)
            {
                serverInProgress= true;

                serverObject = new(sheets, currentSheetIndex, this);

                await serverObject.Initialize();
            }
        }

        private async void ClientInitializer()
        {
            if (clientInProgress)
            {
                //disconnect logic

                clientInProgress = false;
            }
            else if (!serverInProgress)
            {
                clientInProgress = true;

                clientObject = new(sheets, currentSheetIndex, this);

                //(string ipAdress, int port) = CollectServerIpAndPort();

                //await clientObject.Initialize("192.168.0.136", 8080); //Home
                await clientObject.Initialize("192.168.88.161", 8080); //Office
            }
        }

        private (string, int) CollectServerIpAndPort()
        {
            Console.Clear();
            Console.TreatControlCAsInput = false;

            Console.WriteLine("Please enter the ip:\n");
            string ipAdress = Console.ReadLine();

            Console.WriteLine("\nPlease enter the port:");
            int.TryParse(Console.ReadLine(), out int port);

            Console.Clear();
            Console.TreatControlCAsInput = true;

            return (ipAdress, port);
        }

        // MULTIPLE SHEETS
        //-------------------------------------------------------------------------------------
        private void AddNewSheet(string? sheetName = null)
        {
            string newSheetName = sheetName is null ? GenerateNewSheetName() : sheetName;

            sheets.Add((newSheetName, new Sheet()));

            currentSheetName = newSheetName;

            currentSheetIndex = GetSheetIndexFromName(newSheetName);

            UpdateSheetReferences(newSheetName);

            PassImportExportDelegates();

            PassMultipleSheetsDelegates();
        }

        /// <summary>
        /// Allows sheet navigation between sheets if multiple are present.
        /// </summary>
        private void SheetNavigation(int prevOrNext)
        {
            int resultingSheetNameIndex = currentSheetIndex + prevOrNext;

            if (resultingSheetNameIndex >= 0 && resultingSheetNameIndex < sheets.Count)
            {
                currentSheetName = sheets[resultingSheetNameIndex].sheetName;
                currentSheetIndex = resultingSheetNameIndex;
            }
        }

        /// <summary>
        /// Deletes the currnet sheet if there are more than one. The current sheet will always be the
        /// one to the left if exists, otherwise the one to the right.
        /// </summary>
        private void DeleteSheet()
        {
            if (sheets.Count > 1)
            {
                int sheetNameIndex = currentSheetIndex;

                sheets.RemoveAt(sheetNameIndex);

                foreach (var (sheetName, _) in sheets)
                {
                    UpdateSheetReferences(sheetName);
                }

                switch (sheets.Count)
                {
                    case 1:
                        currentSheetIndex = 0;
                        currentSheetName = sheets[0].sheetName;
                        break;

                    default:
                        if (sheetNameIndex - 1 >= 0)
                            (currentSheetIndex, currentSheetName) = (sheetNameIndex - 1, sheets[sheetNameIndex - 1].sheetName);
                        break;
                }
            }
        }

        private void RenameSheet(string newSheetName)
        {
            var oldSheet = sheets[currentSheetIndex];

            sheets.RemoveAt(currentSheetIndex);

            sheets.Insert(currentSheetIndex, (newSheetName, oldSheet.sheet));

            sheets[currentSheetIndex].sheet.currentSheetName = newSheetName;

            currentSheetName = newSheetName;

            UpdateOtherSheetsFormulaReferenceToNewSheetName(oldSheet.sheetName, newSheetName);

            foreach (var (_, sheet) in sheets)
            {
                sheet.sheets = sheets;
            }
        }

        /// <summary>
        /// Updates the cell contents holding formulas from the old sheet name to the new.
        /// </summary>
        private void UpdateOtherSheetsFormulaReferenceToNewSheetName(string oldSheetName, string newSheetName)
        {
            foreach (var (name, sheet) in sheets)
            {
                if (name != newSheetName)
                {
                    var contentToReplace = GetCellsThatNeedModifications(sheet, oldSheetName, newSheetName);

                    foreach (var (key, newFormula) in contentToReplace)
                    {
                        if (sheet.cellData.ContainsKey(key))
                        {
                            sheet.cellData[key] = (CreationHandler.FindValueType(newFormula), newFormula);
                        }
                        else
                        {
                            sheet.cellData.Add(key, (CreationHandler.FindValueType(newFormula), newFormula));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Checks and returns a list of all the cells that contain the old sheet name in a
        /// formula context as well as what the cell should look based on the new sheet name.
        /// </summary>
        private List<((int X, int Y), string newFormula)> GetCellsThatNeedModifications(
            Sheet sheet,
            string oldSheetName,
            string newSheetName
        )
        {
            List<((int X, int Y), string newFormula)> contentToReplace = new();

            foreach (var (key, cell) in sheet.cellData)
            {
                if (cell.Formula is not null && cell.Formula.Contains(oldSheetName + "!") ||
                    cell.Content.TypeValue.ToString().Contains(oldSheetName + "!"))
                {
                    string newFormula = cell.Formula.Replace(oldSheetName, newSheetName);
                    contentToReplace.Add((key, newFormula));
                }
            }

            return contentToReplace;
        }

        /// <summary>
        /// Generates a sheet name of type 'SheetN' where N is the first unused value starting with 1.
        /// </summary>
        private string GenerateNewSheetName()
        {
            int count = 1;
            string sheetName;

            do
            {
                sheetName = "Sheet" + count.ToString();
                count++;
            } while (sheets.Any(sheet => sheet.sheetName == sheetName));

            currentSheetName = sheetName;

            return sheetName;
        }

        // IMPORT/EXPORT
        //-------------------------------------------------------------------------------------

        /// <summary>
        /// Overrides the current sheet with a new instance of Sheet, keeping only the name.
        /// </summary>
        private void NewSheet()
        {
            Console.TreatControlCAsInput = false;

            var oldSheet = sheets[currentSheetIndex];

            sheets[currentSheetIndex] = (oldSheet.sheetName, new Sheet());

            sheets[currentSheetIndex].sheet.currentSheetName = currentSheetName;

            PassImportExportDelegates();

            PassMultipleSheetsDelegates();
        }

        // IMPORT
        //----------------------------------------------------------

        /// <summary>
        /// Creates the .xlsx workbook based on the data found in sheets.
        /// </summary>
        internal ExcelPackage SaveWorkbook(bool networkAction = false)
        {
            string path = string.Empty;
            if (!networkAction && !GetPath(out path))
                return null;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var package = new ExcelPackage();

            foreach (var (sheetName, sheet) in sheets)
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                var pivotCellsToNotBeExported = sheet.PivotTableData.SelectMany(
                    pivot => pivot.pivotCellKeysToNotBeExported
                );

                foreach (var cell in sheet.cellData)
                {
                    if (!pivotCellsToNotBeExported.Contains(
                        (sheets.IndexOf((sheetName, sheet)), cell.Key)
                    ))
                    {
                        if (cell.Value.Formula is not null)
                        {
                            FormatCellForDayOrTimeFormula(worksheet, cell);

                            worksheet.Cells[cell.Key.X, cell.Key.Y].Formula = cell.Value.Formula;
                        }
                        else
                        {
                            worksheet.Cells[cell.Key.X, cell.Key.Y].Value = cell.Value.Content.TypeValue;
                        }
                    }
                }
            }

            sheets.ForEach(nameAndSheet => AddPivotTables(nameAndSheet.sheet, package));

            if (!networkAction)
            {
                package.SaveAs(new FileInfo(path));
                return null;
            }

            return package;
        }

        /// <summary>
        /// Adds pivot table to export .xlsx worksheet if the terminal sheet has one/more based on the
        /// sheet.PivotTableData variable.
        /// </summary>
        private void AddPivotTables(Sheet sheet, ExcelPackage package)
        {
            foreach (var pivotTableData in sheet.PivotTableData)
            {
                var dataWorksheet = package.Workbook.Worksheets[pivotTableData.dataRefIndexAndRange.sheetRefIndex];
                var locationWorksheet = package.Workbook.Worksheets[pivotTableData.locationRefIndexAndStartingCell.sheetRefIndex];

                // define data range for pivot table
                var dataRange = dataWorksheet.Cells[pivotTableData.dataRefIndexAndRange.cellRange];

                // create new pivot table and set location
                var pivotTable = locationWorksheet.PivotTables.Add(
                    locationWorksheet.Cells[pivotTableData.locationRefIndexAndStartingCell.startingCell],
                    dataRange,
                    null
                );

                ExcelPivotTableDataField field;

                foreach (var headerNameAndFormula in pivotTableData.headersAndFormulas)
                {
                    if (pivotTableData.headersAndFormulas.IndexOf(headerNameAndFormula) == 0)
                    {
                        var rowfield = pivotTable.RowFields.Add(pivotTable.Fields[headerNameAndFormula.headerName]);
                        rowfield.Sort = eSortType.Ascending;
                        pivotTable.DataOnRows = false;
                    }
                    else
                    {
                        field = pivotTable.DataFields.Add(pivotTable.Fields[headerNameAndFormula.headerName]);
                        SetFieldNameFunctionAndFormat(field, headerNameAndFormula);
                    }
                }
            }
        }

        /// <summary>
        /// Sets the Name, Format and Function/Formula for a Value Field in the .xlsx worksheet,
        /// after being added to the Pivot Table.
        /// </summary>
        private void SetFieldNameFunctionAndFormat(
            ExcelPivotTableDataField field,
            (string fieldName, int formula) headerNameAndFormula
        )
        {
            switch (headerNameAndFormula.formula)
            {
                case 0:
                    field.Name = $"Sum of {headerNameAndFormula.fieldName}";
                    field.Function = DataFieldFunctions.Sum;
                    field.Format = "0.00";
                    break;

                case 1:
                    field.Name = $"Avg of {headerNameAndFormula.fieldName}";
                    field.Function = DataFieldFunctions.Average;
                    field.Format = "0.00";
                    break;

                case 2:
                    field.Name = $"Count of {headerNameAndFormula.fieldName}";
                    field.Function = DataFieldFunctions.Count;
                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Sets the proper cell format in the .xlsx worksheet if the terminal sheet formula is NOW or TODAY.
        /// </summary>
        private void FormatCellForDayOrTimeFormula(
            ExcelWorksheet worksheet,
            KeyValuePair<(int X, int Y), (IValue Content, string Formula)> cell
        )
        {
            switch (cell.Value.Formula)
            {
                case "=TODAY()":
                    worksheet.Cells[cell.Key.X, cell.Key.Y].Style.Numberformat.Format = "d/m/yyyy";
                    break;
                case "=NOW()":
                    worksheet.Cells[cell.Key.X, cell.Key.Y].Style.Numberformat.Format = "h:mm:ss AM/PM";
                    break;

                default:
                    break;
            }
        }

        //----------------------------------------------------------

        internal void OpenWorkbook(ExcelPackage networkPackage = null)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelPackage package;

            if (networkPackage == null)
            {
                if (!GetPath(out string path, open: true))
                    return;

                sheets.Clear();

                package = new ExcelPackage(new FileInfo(path));
            }
            else
            {
                package = networkPackage;
            }

            using (package)
            {
                foreach (var sheet in package.Workbook.Worksheets)
                {
                    AddNewSheet(sheet.Name);

                    if (sheet.Dimension != null)
                    {
                        AddValuesAndFormulas(sheet);
                    }
                }

                foreach (var sheet in package.Workbook.Worksheets)
                {
                    AddPivotTablesToSheet(sheet);
                }
            }
        }

        private void AddValuesAndFormulas(ExcelWorksheet sheet)
        {
            for (int row = sheet.Dimension.Start.Row; row <= sheet.Dimension.End.Row; row++)
            {
                for (
                    int col = sheet.Dimension.Start.Column;
                    col <= sheet.Dimension.End.Column;
                    col++
                )
                {
                    if (sheet.Cells[row, col].Formula != "")
                    {
                        sheets[currentSheetIndex].sheet.cellData.Add(
                            (row, col),
                            (
                                CreationHandler.FindValueType(""),
                                "=" + sheet.Cells[row, col].Formula
                            )
                        );
                    }
                    else if (sheet.Cells[row, col].Value != null)
                    {
                        sheets[currentSheetIndex].sheet.cellData.Add(
                            (row, col),
                            (
                                CreationHandler.FindValueType(
                                    sheet.Cells[row, col].Value.ToString()
                                ),
                                null
                            )
                        );
                    }
                }
            }
        }

        private void AddPivotTablesToSheet(ExcelWorksheet worksheet)
        {
            foreach (var pivotTable in worksheet.PivotTables)
            {
                string dataRange = pivotTable.CacheDefinition.SourceRange.FullAddress;

                string location = pivotTable.Address.ToString();

                string selectedRow = pivotTable.RowFields.Select(x => x.Name).First();

                IEnumerable<(int formula, string columnName)> formulaAndSelectedColumn =
                    pivotTable.DataFields
                        .Select(x => x.Name)
                        .Select(name =>
                        {
                            string[] words = name.Split();
                            int index = words[0] switch
                            {
                                "Sum" => 0,
                                "Avg" => 1,
                                "Count" => 2,
                                _ => throw new ArgumentException($"Invalid name: {name}")
                            };
                            string value = words[^1];
                            return (index, value);
                        });

                sheets[
                    sheets.FindIndex(nameAndSheet => nameAndSheet.sheetName == worksheet.Name)
                ].sheet.pivotHandler.ImportPivotTable(
                    dataRange,
                    location,
                    selectedRow,
                    formulaAndSelectedColumn
                );
            }

            currentSheetIndex = 0;
            currentSheetName = sheets[currentSheetIndex].sheetName;
        }

        private bool GetPath(out string path, bool open = false)
        {
            Console.Clear();
            Console.TreatControlCAsInput = false;
            Console.CursorVisible = true;

            if (open)
            {
                Console.WriteLine(
                    "OPEN:\nPlease enter a path or file name if the file is local or type \"EXIT\" "
                    + "to return to the sheet:\n(Example: C:\\Users\\User\\Desktop\\FileName.xlsx or FileName.xlsx)"
                );
            }
            else
            {
                Console.WriteLine(
                    "SAVE:\nPlease enter a path or file name if the file is local or type \"EXIT\" "
                    + "to return to the sheet:\n(Example: C:\\Users\\User\\Desktop\\FileName.xlsx or FileName.xlsx)"
                );
            }

            path = Console.ReadLine();
            path = ValidatePath(path);

            return path is null ? false : true;
        }

        string ValidatePath(string path)
        {
            if (!path.EndsWith(".xlsx"))
            {
                return null;
            }

            if (path.Contains("\\"))
            {
                return path;
            }
            else
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                path = Path.Combine(desktopPath, path);
            }

            return path;
        }

        private int GetSheetIndexFromName(string sheetNameToFind)
        {
            return sheets.FindIndex(name => name.sheetName == sheetNameToFind);
        }

        private void UpdateSheetReferences(string currentSheetName)
        {
            int sheetIndex = GetSheetIndexFromName(currentSheetName);

            sheets[sheetIndex].sheet.sheets = sheets;

            sheets[sheetIndex].sheet.currentSheetName = currentSheetName;

            sheets[sheetIndex].sheet.currentSheetIndex = sheetIndex;

            sheets[sheetIndex].sheet.isTesting = true;
        }

        /// <summary>
        /// Initializes all the handlers by calling the constructors (Creation, Formula, Print).
        /// </summary>
        private void InitializeSheetHandlers()
        {
            int sheetIndex = GetSheetIndexFromName(currentSheetName);

            sheets[sheetIndex].sheet.creationHandler = new CreationHandler(
                sheets[sheetIndex].sheet
            );

            sheets[sheetIndex].sheet.formulaHandler = new FormulaHandler(sheets[sheetIndex].sheet);

            sheets[sheetIndex].sheet.printHandler = new PrintHandler(sheets[sheetIndex].sheet);

            sheets[sheetIndex].sheet.pivotHandler = new PivotHandler(sheets[sheetIndex].sheet);
        }

        /// <summary>
        /// Passes the SaveSheet, OpenSheet, NewSheet delegates to the sheet. This way all of them can be called from the sheet.Execute() method.
        /// </summary>
        private void PassImportExportDelegates()
        {
            int sheetIndex = GetSheetIndexFromName(currentSheetName);

            sheets[sheetIndex].sheet.saveDelegate = SaveWorkbook;

            sheets[sheetIndex].sheet.openDelegate = OpenWorkbook;

            sheets[sheetIndex].sheet.newDelegate = NewSheet;
        }

        /// <summary>
        /// Passes the SaveSheet, OpenSheet, NewSheet delegates to the sheet. This way all of them can be called from the sheet.Execute() method.
        /// </summary>
        private void PassMultipleSheetsDelegates()
        {
            int sheetIndex = GetSheetIndexFromName(currentSheetName);

            sheets[sheetIndex].sheet.InitializeGlobalCopyCutVariable(
                GlobalCopyCutVariable.GetInstance()
            );

            sheets[sheetIndex].sheet.addSheetDelegate = AddNewSheet;

            sheets[sheetIndex].sheet.sheetNavigationDelegate = SheetNavigation;

            sheets[sheetIndex].sheet.deleteSheetDelegate = DeleteSheet;

            sheets[sheetIndex].sheet.renameSheetDelegate = RenameSheet;
        }

        //XUNIT TESTING ------------------------------------------------------------------------------------

        public void TestingRunNewSheet()
        {
            sheets.Add((GenerateNewSheetName(), new Sheet()));

            UpdateSheetReferences(currentSheetName);

            PassImportExportDelegates();

            PassMultipleSheetsDelegates();
        }

        public void RegisterAndExecuteActionsTesting(ConsoleKeyInfo key)
        {
            sheets[currentSheetIndex].sheet.Execute(key);
        }

        public void CreatePivotTableTesting(
            string dataRange,
            string location,
            string selectedRow,
            IEnumerable<(int formula, string valueName)> formulaAndSelectedColumn
        )
        {
            sheets[currentSheetIndex].sheet.pivotHandler.
                ImportPivotTable(dataRange, location, selectedRow, formulaAndSelectedColumn);
        }

        public (IValue Content, string? Formula) CellContentAndFormulaAt(int x, int y)
        {
            return sheets[currentSheetIndex].sheet.cellData.ContainsKey((x, y)) 
                ? sheets[currentSheetIndex].sheet.CellContentAndFormulaAt(x, y) 
                : (null, null);
        }

        public (int X, int Y) GetCurrentPosition()
        {
            return sheets[currentSheetIndex].sheet.currentPosition;
        }

        public (int currentSheetIndex, List<(string sheetName, Sheet sheet)> nameAndSheet) GetSheetsInformation()
        {
            return (currentSheetIndex, sheets);
        }

    }
}
