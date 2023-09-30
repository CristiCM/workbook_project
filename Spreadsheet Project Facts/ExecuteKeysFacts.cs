using Spreadsheet_Project;
using static Spreadsheet_Project_Facts.ConsoleKeyInfosFeeder;

namespace Spreadsheet_Project_Facts
{
    public class ExecuteKeysFacts
    {
        ConsoleKeyInfosFeeder keyFeeder = new();

        [Fact]
        public void NavigateToNextRightCell_ShouldMoveCursorToTheNextRightCellOnRightArrowPress()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            Assert.Equal((1, 1), test.GetCurrentPosition());

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));

            Assert.Equal((1, 3), test.GetCurrentPosition());
        }

        [Fact]
        public void NavigateToNextLeftCell_ShouldMoveCursorToTheNextLeftCellOnLeftArrowPress()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));

            Assert.Equal((1, 3), test.GetCurrentPosition());

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Left));

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Left));

            Assert.Equal((1, 1), test.GetCurrentPosition());
        }

        [Fact]
        public void NavigateToNextBelowCell_ShouldMoveCursorToTheNextBelowCellOnDownArrowPress()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            Assert.Equal((1, 1), test.GetCurrentPosition());

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal((3, 1), test.GetCurrentPosition());
        }

        [Fact]
        public void NavigateToNextAboveCell_ShouldMoveCursorToTheNextAboveCellOnUpArrowPress()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal((3, 1), test.GetCurrentPosition());

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Up));

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Up));

            Assert.Equal((1, 1), test.GetCurrentPosition());
        }

        [Fact]
        public void Backspace_ShouldTrimOffTheLastCharFromTheCellIfAnyExist()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");

            foreach (var key in testType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }


            Assert.Equal("testing string", test.CellContentAndFormulaAt(1,1).Content.TypeValue);

            for (int i = 0; i < 7; i++)
            {
                test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Backspace));
            }

            Assert.Equal("testing", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void Backspace_ShouldTrimOffTheLastCharFromTheCellWhileCharsExist()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");

            foreach (var key in testType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }


            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);

            for (int i = 0; i < 20; i++)
            {
                test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Backspace));
            }

            Assert.Equal("", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void NewSheet_ShouldInitializeANewSheetDeletingAllCurrentContents()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");

            for (int i = 0; i < 10; i++)
            {
                foreach (var key in testType)
                {
                    test.RegisterAndExecuteActionsTesting(key);
                }

                Assert.Equal("testing string", test.CellContentAndFormulaAt(1, i + 1).Content.TypeValue);

                test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));
            }


            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.NewSheet));

            for (int i = 0; i < 10; i++)
            {
                Assert.Equal((null, null), test.CellContentAndFormulaAt(1, i + 1));
            }
        }

        [Fact]
        public void CopyPasteSameSheet_ShouldCopyTheItemAndBeAbleToPasteIt()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");
            
            foreach (var key in testType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Copy));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Paste));

            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 2).Content.TypeValue);
        }

        [Fact]
        public void CopyPasteCrossSheet_ShouldCopyTheItemAndBeAbleToPasteIt()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");

            foreach (var key in testType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Copy));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Paste));

            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.PrevSheet));

            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void CutPasteSameSheet_ShouldCopyTheItemAndBeAbleToPasteIt()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");

            foreach (var key in testType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Cut));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Paste));

            Assert.Equal((null, null), test.CellContentAndFormulaAt(1, 1));

            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 2).Content.TypeValue);
        }

        [Fact]
        public void CutPasteCrossSheet_ShouldCopyTheItemAndBeAbleToPasteIt()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");

            foreach (var key in testType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Cut));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Paste));

            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.PrevSheet));

            Assert.Equal((null, null), test.CellContentAndFormulaAt(1, 1));
        }

        [Fact]
        public void Delete_ShouldRemoveTheCellFromTheSheet()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");

            foreach (var key in testType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Delete));

            Assert.Equal((null, null), test.CellContentAndFormulaAt(1, 1));
        }

        [Fact]
        public void TypeOnMovedToCell_ShouldDeleteTheCellBeforeAddingToIt()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");

            var secondTestType = keyFeeder.CreateKeyFeedFromString("second testing string");

            foreach (var key in testType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Up));

            foreach (var key in secondTestType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            Assert.Equal("second testing string", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void TypeOnMovedToCellWithF2_ShouldEnterTheEditModeAndAppendToCell()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var testType = keyFeeder.CreateKeyFeedFromString("testing string");

            var appendString = keyFeeder.CreateKeyFeedFromString(" append");

            foreach (var key in testType)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            Assert.Equal("testing string", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Up));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.F2));

            foreach (var key in appendString)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            Assert.Equal("testing string append", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void AddNewSheet_ShouldAddTwoMoreSheetsToTheExistingOne()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            AddSheets(test, 2);

            var sheetNames = test.GetSheetsInformation().nameAndSheet.Select(sheetInfo => sheetInfo.sheetName).ToArray();

            Assert.Equal(new string[] { "Sheet1", "Sheet2", "Sheet3" }, sheetNames);
        }

        [Fact]
        public void SheetNavigation_ShouldAllowChangingTheShownSheet()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            AddSheets(test, 2);

            var sheetIndex = test.GetSheetsInformation().currentSheetIndex;

            Assert.Equal(2, sheetIndex);

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.PrevSheet));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.PrevSheet));

            sheetIndex = test.GetSheetsInformation().currentSheetIndex;

            Assert.Equal(0, sheetIndex);

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.NextSheet));

            sheetIndex = test.GetSheetsInformation().currentSheetIndex;

            Assert.Equal(1, sheetIndex);
        }

        [Fact]
        public void DeleteSheet_ShouldAllowCurrentSheetDeletion()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            AddSheets(test, 2);

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.PrevSheet));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.PrevSheet));

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.DeleteSheet));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.DeleteSheet));

            var sheetNames = test.GetSheetsInformation().nameAndSheet.Select(sheetInfo => sheetInfo.sheetName).ToArray();

            Assert.Equal(new string[] { "Sheet3" }, sheetNames);
        }

        [Fact]
        public void RenameSheet_ShouldAllowSheetRenaming()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var sheetNames = test.GetSheetsInformation().nameAndSheet.Select(sheetInfo => sheetInfo.sheetName).ToArray();

            Assert.Equal(new string[] { "Sheet1" }, sheetNames);

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.RenameSheet));

            sheetNames = test.GetSheetsInformation().nameAndSheet.Select(sheetInfo => sheetInfo.sheetName).ToArray();

            Assert.Equal(new string[] { "Testing" }, sheetNames);
        }

        private void AddSheets(ApplicationInitializer test, int numberOfSheets)
        {
            for (int i = 0; i < numberOfSheets; i++)
            {
                test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));
            }
        }
    }
}
