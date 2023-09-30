using Newtonsoft.Json;
using OfficeOpenXml;
using Spreadsheet_Project.Formulas;

namespace Spreadsheet_Project
{
    public class Sheet
    {
        internal Dictionary<(int X, int Y), (IValue Content, string? Formula)> cellData;
        public (int X, int Y) currentPosition;

        private List<string> legendInfo;
        private int[] movementKeys;
        internal (bool, (int, int)) movedToAnotherCell;
        GlobalCopyCutVariable cutCopyData;

        [JsonIgnore]
        public List<string> Alphabet;

        // NAVIGATION AND MULTISHEET
        //-------------------------------------------------------------------------------------
        internal List<(string sheetName, Sheet sheet)> sheets;
        internal string currentSheetName;
        internal int currentSheetIndex;
        internal Action newDelegate;
        internal Action<ExcelPackage> openDelegate;
        internal Func<bool, ExcelPackage> saveDelegate;
        internal Action<string> addSheetDelegate;
        internal Action<int> sheetNavigationDelegate;
        internal Action deleteSheetDelegate;
        internal Action<string> renameSheetDelegate;
        //-------------------------------------------------------------------------------------

        // HANDLERS
        //-------------------------------------------------------------------------------------
        internal FormulaHandler formulaHandler;
        internal CreationHandler creationHandler;
        internal PrintHandler printHandler;
        internal PivotHandler pivotHandler;
        //-------------------------------------------------------------------------------------

        // PIVOT
        //-------------------------------------------------------------------------------------
        internal bool pivotMenuActive;
        internal List<(
            (int sheetRefIndex, string cellRange) dataRefIndexAndRange,
            (
                int sheetRefIndex,
                string startingCell,
                (int X, int Y) positionKey
            ) locationRefIndexAndStartingCell,
            List<(string headerName, int sheetRefIndex)> headersAndFormulas,
            List<(int, (int, int))> pivotCellKeysToNotBeExported
        )> PivotTableData;
        //-------------------------------------------------------------------------------------

        //TESTING
        //-------------------------------------------------------------------------------------
        internal bool isTesting = false;

        // CONSTRUCTOR
        public Sheet()
        {
            currentPosition = (1, 1);
            movedToAnotherCell = (false, (0,0));
            pivotMenuActive = false;

            cellData = new();
            legendInfo = new()
            {
                "  Keys Legend:",
                "------------------------",
                "  CTRL + N - New sheet",
                "  CTRL + S - Save sheet",
                "  CTRL + O - Open sheet",
                "  CTRL + X - Cut",
                "  CTRL + C - Copy",
                "  CTRL + V - Paste",
                "  F2 - Edit existing cell",
                "  F5 - New Sheet",
                "  F6 - Previous Sheet",
                "  F7 - Next Sheet",
                "  F8 - Delete Sheet",
                "------------------------",
                "\n  Formulas:",
                "---------------------------------------------------------------------------",
                "  Cell Reference \"=A4\"",
                "  Sum            \"=SUM(A1,B1,C1)\"",
                "  Average        \"=AVERAGE(A1,B1,C1)\"",
                "  Count          \"=COUNT(A1,B1,C1)\"",
                "  Mod            \"=MOD(number, divisor)\"",
                "  Power          \"=POWER(number, power)\"",
                "  Ceiling        \"=CEILING(number, significance)\"",
                "  Floor          \"=FLOOR(number, significance)\"",
                "  Concat         \"=CONCATENATE(A1, B1, C1)\"",
                "  Length         \"=LEN(A1)\"",
                "  Replace        \"=REPLACE(old_text, start_index, num_chars, new_text)\"",
                "  Substitute     \"=SUBSTITUTE(text, old_text, new_text, instance_num (optional))\"",
                "  Now            \"=NOW()\"",
                "  Today          \"=TODAY()\"",
                "  Vlookup        \"=VLOOKUP(lookup_value, table_array, col_index_num)\"",
                "  Subtotal       \"=SUBTOTAL(func_index, A1,B1,C1)\"",
                "  -----",
                "  func_index:",
                "  1 - AVG",
                "  2 - COUNT",
                "  9 - SUM",
                "  -----",
                "---------------------------------------------------------------------------",
                "\n  \"Enter\" to return to sheet."
            };

            movementKeys = new int[]
            {
                (int)ConsoleKey.LeftArrow,
                (int)ConsoleKey.RightArrow,
                (int)ConsoleKey.DownArrow,
                (int)ConsoleKey.UpArrow,
                (int)ConsoleKey.Tab,
                (int)ConsoleKey.Enter
            };

            creationHandler = new CreationHandler(this);
            printHandler = new PrintHandler(this);
            formulaHandler = new FormulaHandler(this);
            pivotHandler = new PivotHandler(this);
            PivotTableData = new();
            try
            {
                Console.CursorVisible = false;
            }
            catch
            {

            }
            
        }

        public void InitializeGlobalCopyCutVariable(GlobalCopyCutVariable globalCopyCutVariable)
        {
            cutCopyData = globalCopyCutVariable;
        }

        /// <summary>
        /// Prints the legend that includes all the implemented functionalities.
        /// </summary>
        public void Legend()
        {
            Console.Clear();
            legendInfo.ForEach(info => Console.WriteLine(info));
            Console.TreatControlCAsInput = false;
            Console.ReadLine();
            Console.TreatControlCAsInput = true;
            Console.Clear();
            printHandler.PrintSheet();
        }

        /// <summary>
        /// Reacts based on the user inputted key by running over all the possible key combinations.
        /// </summary>
        public void Execute(ConsoleKeyInfo key)
        {
            switch (key.Key, key.Modifiers)
            {
                case (ConsoleKey.L, ConsoleModifiers.Control):
                    Legend();
                    break;
                case (ConsoleKey.P, ConsoleModifiers.Control):
                    pivotHandler.RunPivot();
                    break;
                case (ConsoleKey.LeftArrow, 0):
                    UpdateCursorPosition(0, -1);
                    break;
                case (ConsoleKey.RightArrow or ConsoleKey.Tab, 0):
                    UpdateCursorPosition(0, 1);
                    break;
                case (ConsoleKey.UpArrow, 0):
                    UpdateCursorPosition(-1, 0);
                    break;
                case (ConsoleKey.DownArrow or ConsoleKey.Enter, 0):
                    UpdateCursorPosition(1, 0);
                    break;
                case (ConsoleKey.Backspace, 0):
                    Backspace();
                    break;
                case (ConsoleKey.Delete, 0):
                    Delete();
                    break;
                case (ConsoleKey.S, ConsoleModifiers.Control):
                    SaveSheet();
                    break;
                case (ConsoleKey.O, ConsoleModifiers.Control):
                    OpenSheet();
                    break;
                case (ConsoleKey.N, ConsoleModifiers.Control):
                    NewSheet();
                    break;
                case (ConsoleKey.X, ConsoleModifiers.Control):
                    Cut();
                    break;
                case (ConsoleKey.C, ConsoleModifiers.Control):
                    Copy();
                    break;
                case (ConsoleKey.V, ConsoleModifiers.Control):
                    Paste();
                    break;
                case (ConsoleKey.F2, 0):
                    movedToAnotherCell = (false, (0,0));
                    break;
                case (ConsoleKey.F5, 0):
                    AddSheet();
                    break;
                case (ConsoleKey.F6, 0):
                    PreviousSheet();
                    break;
                case (ConsoleKey.F7, 0):
                    NextSheet();
                    break;
                case (ConsoleKey.F8, 0):
                    DeleteSheet();
                    break;
                case (ConsoleKey.F9, 0):
                    RenameSheet();
                    break;
                default:
                    AppendKeyToCell(key.KeyChar);
                    break;
            }
        }

        // ConsoleKeyInfo ACTIONS
        //-------------------------------------------------------------------------------------

        private void Paste()
        {
            var (Contents, cut, sheetReferenceIndex) = this.cutCopyData.GetValue();

            if (Contents.Content is null)
                return;

            if (cut && sheetReferenceIndex != currentSheetIndex && Contents.Formula != null)
            {
                var sheetName = sheets[sheetReferenceIndex].sheetName;
                Contents.Formula = Contents.Formula.Replace("(", $"({sheetName}!");
            }

            cellData[currentPosition] = Contents;

            if (cut)
                cutCopyData.SetValue(((null, null), false, -1));
        }

        private void Cut()
        {
            if (cellData.ContainsKey(currentPosition))
            {
                cutCopyData.SetValue((cellData[currentPosition], true, currentSheetIndex));
                cellData.Remove(currentPosition);
            }
        }

        private void Copy()
        {
            if (cellData.ContainsKey(currentPosition))
            {
                cutCopyData.SetValue((cellData[currentPosition], false, currentSheetIndex));
            }
        }

        private void Backspace()
        {
            if (!cellData.ContainsKey(currentPosition))
                return;

            string curCellValue = cellData[currentPosition].Formula is null
                ? cellData[currentPosition].Content.TypeValue.ToString()
                : cellData[currentPosition].Formula;

            if (curCellValue.Length > 0)
                cellData[currentPosition] = (CreationHandler.FindValueType(curCellValue[..^1]), null);
        }

        private void Delete()
        {
            if (cellData.ContainsKey(currentPosition))
            {
                if (PivotTableData.Count > 0)
                {
                    pivotHandler.DeletePreviousPivotOnPosition(currentSheetIndex, currentPosition);
                }
                else
                {
                    cellData.Remove(currentPosition);
                }
            }
        }

        private void AppendKeyToCell(char key)
        {
            ResetCellIfMovedTo();

            string curCellValue = cellData.ContainsKey(currentPosition)
                ? cellData[currentPosition].Formula ?? cellData[currentPosition].Content.TypeValue.ToString()
                : string.Empty;

            cellData[currentPosition] = (
                CreationHandler.FindValueType(curCellValue += key),
                null
            );
        }

        // * Movement
        private void UpdateCursorPosition(int X, int Y)
        {
            if (cellData.ContainsKey(currentPosition))
            {
                if (cellData[currentPosition].Formula is null)
                {
                    ApplyFormula();
                }
                else
                {
                    printHandler.UpdateContentBasedOnFormula(currentPosition.X, currentPosition.Y);
                }
            }

            if (currentPosition.X + X > 0 && currentPosition.Y + Y > 0)
            {
                currentPosition.X += X;
                currentPosition.Y += Y;
                movedToAnotherCell = (true, currentPosition);
            }
        }

        private void ApplyFormula()
        {
            var formulaCheck = formulaHandler.TryFormulas(currentPosition);

            if (formulaCheck is not null)
            {
                cellData[currentPosition] = (
                    CreationHandler.FindValueType(formulaCheck),
                    cellData[currentPosition].Content.TypeValue.ToString()
                );
            }
        }

        /// <summary>
        /// Deletes the cellData Dictionary KVP in the case of the currentPosition cursor just arriving on the position.
        /// </summary>
        private void ResetCellIfMovedTo()
        {
            if (movedToAnotherCell.Item1 && movedToAnotherCell.Item2 == currentPosition)
            {
                cellData.Remove(currentPosition);
                movedToAnotherCell = (false, (0, 0));
            }
        }

        // NEW / IMPORT / EXPORT
        //-------------------------------------------------------------------------------------
        private void NewSheet()
        {
            if (!isTesting)
            Console.Clear();

            newDelegate.Invoke();
        }

        private void OpenSheet()
        {
            openDelegate.Invoke(null);
        }

        private void SaveSheet()
        {
            saveDelegate.Invoke(false);
            printHandler.PrintSheet();
        }

        // MULTIPLE SHEET MANAGEMENT
        //-------------------------------------------------------------------------------------
        private void AddSheet()
        {
            addSheetDelegate.Invoke(null);
        }

        private void PreviousSheet()
        {
            sheetNavigationDelegate.Invoke(-1);
        }

        private void NextSheet()
        {
            sheetNavigationDelegate.Invoke(1);
        }

        private void DeleteSheet()
        {
            deleteSheetDelegate.Invoke();
        }

        private void RenameSheet()
        {
            if (isTesting)
            {
                renameSheetDelegate.Invoke("Testing");
                return;
            }

            Console.Clear();
            Console.WriteLine("Please type a new name for this sheet:");

            Console.TreatControlCAsInput = false;
            string newSheetName = Console.ReadLine();
            Console.TreatControlCAsInput = true;

            renameSheetDelegate.Invoke(newSheetName);
        }

        // XUNIT TESTING
        //-------------------------------------------------------------------------------------

        public (IValue Content, string? Formula) CellContentAndFormulaAt(int x, int y)
        {
            if (cellData.ContainsKey((x, y)))
            {
                return cellData[(x, y)];
            }
            else
            {
                return (null, null);
            }
        }
    }
}
