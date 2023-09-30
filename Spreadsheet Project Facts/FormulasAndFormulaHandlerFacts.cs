using Spreadsheet_Project;
using static Spreadsheet_Project_Facts.ConsoleKeyInfosFeeder;

namespace Spreadsheet_Project_Facts
{
    public class FormulasAndFormulaHandlerFacts
    {
        ConsoleKeyInfosFeeder keyFeeder = new();

        //VLOOKUP
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void VlookupSameSheet_ShouldDisplayFirstMatchValueFromIndicatedColumn()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=VLOOKUP(\"bob\", A1:D8, 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal(100, test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void VlookupCrossSheet_ShouldDisplayFirstMatchValueFromIndicatedColumn()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=VLOOKUP(\"bob\", Sheet1!A1:D8, 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal(100, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void VlookupRecursion_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=VLOOKUP(\"bob\", A1:D9, 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(9,1).Content.TypeValue);
        }

        [Fact]
        public void VlookupTextElementsFormat_ShouldShowAnErrorIfTextNotInQuotations()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=VLOOKUP(bob, A1:D8, 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("#NAME", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void VlookupIndexing_ShouldShowAnErrorIfTheIndexIsOutOfTheRangeBounds()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=VLOOKUP(\"bob\", A1:D8, 6)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("#REF!", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }


        //SUBSTITUTE
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void SubstituteSameSheet_ShouldReplaceaSpecifiedStringWithAnother()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBSTITUTE(A2, \"b\", \"x\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("xox", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void SubstituteCrossSheet_ShouldReplaceaSpecifiedStringWithAnother()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBSTITUTE(Sheet1!A2, \"b\", \"x\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("xox", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void SubstituteSameSheetFirstOccurence_ShouldReplaceaSpecifiedStringWithAnother()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBSTITUTE(A2, \"b\", \"x\", 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("xob", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void SubstituteCrossSheetFirstOccurence_ShouldReplaceaSpecifiedStringWithAnother()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBSTITUTE(Sheet1!A2, \"b\", \"x\", 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("xob", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void SubstituteSameSheetSecondOccurence_ShouldReplaceaSpecifiedStringWithAnother()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBSTITUTE(A2, \"b\", \"x\", 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("box", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void SubstituteCrossSheetSecondOccurence_ShouldReplaceaSpecifiedStringWithAnother()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBSTITUTE(Sheet1!A2, \"b\", \"x\", 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("box", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void SubstituteSameSheetFirst_ShouldReplaceaSpecifiedStringWithAnother()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBSTITUTE(A2, \"b\", \"x\", 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("xob", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void SubstituteRecursion_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBSTITUTE(A2, A9, \"x\", 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void SubstituteTextElementFormat_ShouldShowAnErrorIfTextNotInQuotations()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBSTITUTE(A2, b, \"x\", 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("#NAME", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }


        //REPLACE
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void ReplaceSameSheet_ShouldReplaceFirstCharacter()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=REPLACE(A2, 1, 1, \"x\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("xob", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void ReplaceSameSheet_ShouldReplaceAllCharacters()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=REPLACE(A2, 1, 3, \"z\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("z", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void ReplaceCrossSheet_ShouldReplaceFirstCharacter()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));
            
            var formula = keyFeeder.CreateKeyFeedFromString("=REPLACE(Sheet1!A2, 1, 1, \"x\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("xob", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void ReplaceCrossSheet_ShouldReplaceAllCharacters()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=REPLACE(Sheet1!A2, 1, 3, \"z\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("z", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void ReplaceRecursive_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=REPLACE(A9, 1, 1, \"x\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void ReplaceIntElementFormat_ShouldShowAnErrorIfElementNotInt()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=REPLACE(A2, z, 1, \"x\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("#NAME", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void ReplaceTextElementFormat_ShouldShowAnErrorIfTextNotInQuotations()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=REPLACE(bob, 1, 1, \"x\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("#NAME", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        //CONCATENATE
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void ConcatenateSameSheet_ShouldConcatenateCellValues()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=CONCATENATE(A1, B1, C1:D1, \"123\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("userstakeprofitroi123", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void ConcatenateCrossSheet_ShouldConcatenateCellValues()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=CONCATENATE(Sheet1!A1, Sheet1!B1, Sheet1!C1:D1, \"123\")");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("userstakeprofitroi123", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void ConcatenateTextElementFormat_ShouldShowAnErrorIfTextNotInQuotations()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=CONCATENATE(A1, B1, C1:D1, text)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("#NAME", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        //LEN
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void LenSameSheet_ReturnLength()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula1 = keyFeeder.CreateKeyFeedFromString("=LEN(A2)");

            foreach (var key in formula1)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula2 = keyFeeder.CreateKeyFeedFromString("=LEN(\"bob\")");

            foreach (var key in formula2)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(3, test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
            Assert.Equal(3, test.CellContentAndFormulaAt(10, 1).Content.TypeValue);
        }

        [Fact]
        public void LenCrossSheet_ReturnLength()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula1 = keyFeeder.CreateKeyFeedFromString("=LEN(Sheet1!A2)");

            foreach (var key in formula1)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(3, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void LenRecursive_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=LEN(A9)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        [Fact]
        public void LenTextElementFormat_ShouldShowAnErrorIfTextNotInQuotations()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=LEN(bob)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));


            Assert.Equal("#NAME", test.CellContentAndFormulaAt(9, 1).Content.TypeValue);
        }

        //FLOOR
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void FloorSameSheet_ShouldRoundDown()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("1.7");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula = keyFeeder.CreateKeyFeedFromString("=FLOOR(A1, 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(1, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void FloorCrossSheet_ShouldRoundDown()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("1.7");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=FLOOR(Sheet1!A1, 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(1, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void FloorRecursive_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=FLOOR(A1, 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void FloorNumberFormat_ShouldShowAnErrorIfElementsNotNumber()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=FLOOR(z, 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula1 = keyFeeder.CreateKeyFeedFromString("=FLOOR(1.7, z)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("#VALUE!", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
            Assert.Equal("#VALUE!", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        //CEILING
        //------------------------------------------------------------------------------------------------

        [Fact]
        public void CeilingSameSheet_ShouldRoundUp()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("1.2");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula = keyFeeder.CreateKeyFeedFromString("=CEILING(A1, 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(2, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void CeilingCrossSheet_ShouldRoundUp()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("1.2");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=CEILING(Sheet1!A1, 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(2, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void CeilingRecursive_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=CEILING(A1, 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void CeilingNumberFormat_ShouldShowAnErrorIfElementsNotNumber()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=CEILING(z, 1)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula1 = keyFeeder.CreateKeyFeedFromString("=CEILING(1.7, z)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("#VALUE!", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
            Assert.Equal("#VALUE!", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        //POWER
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void PowerSameSheet_ShouldRaiseToPower()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("2");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula = keyFeeder.CreateKeyFeedFromString("=POWER(A1, 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(4, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void PowerCrossSheet_ShouldRaiseToPower()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("2");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=POWER(Sheet1!A1, 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(4, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void PowerRecursive_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=POWER(A1, 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void PowerNumberFormat_ShouldShowAnErrorIfElementsNotNumber()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=POWER(z, 2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula1 = keyFeeder.CreateKeyFeedFromString("=POWER(2, z)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("#VALUE!", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
            Assert.Equal("#VALUE!", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        //MOD
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void ModSameSheet_ShouldReturnTheRemining()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("10");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula = keyFeeder.CreateKeyFeedFromString("=MOD(A1, 9)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(1, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void ModCrossSheet_ShouldReturnTheRemining()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("10");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=MOD(Sheet1!A1, 9)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(1, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void ModRecursive_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=POWER(A1, 9)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void ModNumberFormat_ShouldShowAnErrorIfElementsNotNumber()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=POWER(z, 9)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula1 = keyFeeder.CreateKeyFeedFromString("=POWER(10, z)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("#VALUE!", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
            Assert.Equal("#VALUE!", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        //SUM
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void SumSameSheet_ShouldAddTheElements()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUM(A1, B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(6, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void SumCrossSheet_ShouldAddTheElements()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=SUM(Sheet1!A1, Sheet1!B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(6, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void SumRecursive_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUM(A1, B1:C1, 3, A2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        //AVERAGE
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void AverageSameSheet_ShouldReturnTheAvg()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=AVERAGE(A1, B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(2, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void AverageCrossSheet_ShouldReturnTheAvg()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=AVERAGE(Sheet1!A1, Sheet1!B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(2, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void AverageRecursive_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=AVERAGE(A1, B1:C1, 3, A2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        //COUNT
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void CountSameSheet_ShouldReturnTheElementCount()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=COUNT(A1, B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(3, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void CountCrossSheet_ShouldReturnTheElementCount()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=COUNT(Sheet1!A1, Sheet1!B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(3, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void CountRecursive_ShouldShowARecursiveErrorIfRangesContainTheFormulaCell()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=COUNT(A1, B1:C1, 3, A2)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        //SUBTOTAL
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void SubtotalSumSameSheet_ShouldReturnAddedElements()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBTOTAL(9, A1, B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(6, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void SubtotalSumCrossSheet_ShouldReturnAddedElements()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBTOTAL(9, Sheet1!A1, Sheet1!B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(6, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void SubtotalAverageSameSheet_ShouldReturnAvgOfElements()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBTOTAL(1, A1, B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(2, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void SubtotalAverageCrossSheet_ShouldReturnAvgOfElements()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBTOTAL(1, Sheet1!A1, Sheet1!B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(2, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void SubtotalCountSameSheet_ShouldReturnCountOfElements()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBTOTAL(2, A1, B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(3, test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void SubtotalCountCrossSheet_ShouldReturnCountOfElements()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBTOTAL(2, Sheet1!A1, Sheet1!B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal(3, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void SubtotalWrongIndexSheet_ShouldNotReturnAnything()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBTOTAL(5, A1, B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("=SUBTOTAL(5, A1, B1:C1, 3)", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void SubtotalWrongIndexCrossSheet_ShouldNotReturnAnything()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateInitialDataKeyFeed(new string[][] { new[] { "1", "2", "text" } });

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=SUBTOTAL(5, Sheet1!A1, Sheet1!B1:C1, 3)");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("=SUBTOTAL(5, Sheet1!A1, Sheet1!B1:C1, 3)", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        //STRINGOVERRIDE
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void StringOverride_ShouldRoundDown()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("=\"123tt\"");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("123tt", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        //NOW
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void Now_ShouldReturnCurrentTime()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("=NOW()");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            DateTime currentTime = DateTime.Now;

            Assert.Equal(currentTime.ToString("HH:mm:ss"), test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        //TODAY
        //------------------------------------------------------------------------------------------------
        [Fact]
        public void TODAY_ShouldReturnCurrentDate()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("=TODAY()");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            DateTime currentTime = DateTime.Today;

            Assert.Equal(currentTime.ToString("dd-MM-yyyy"), test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        //CELL REFERENCE
        //------------------------------------------------------------------------------------------------

        [Fact]
        public void CellReferenceCurrentSheet_ShouldReturnCellValue()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("testtext");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula = keyFeeder.CreateKeyFeedFromString("=A1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("testtext", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void CellReferenceCrossSheet_ShouldReturnCellValue()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("testtext");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }


            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=Sheet1!A1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("testtext", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void CellReferenceEmptyCellCurrentSheet_ShouldReturnCellValue()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula = keyFeeder.CreateKeyFeedFromString("=A1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void CellReferenceEmptyCellCrossSheet_ShouldReturnCellValue()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=Sheet1!A1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void CellReferenceChainCurrentSheet_ShouldReturnCellValue()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("testtext");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula = keyFeeder.CreateKeyFeedFromString("=A1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula1 = keyFeeder.CreateKeyFeedFromString("=A2");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula2 = keyFeeder.CreateKeyFeedFromString("=A3");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("testtext", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
            Assert.Equal("testtext", test.CellContentAndFormulaAt(3, 1).Content.TypeValue);
            Assert.Equal("testtext", test.CellContentAndFormulaAt(4, 1).Content.TypeValue);
        }

        [Fact]
        public void CellReferenceChainCrossSheet_ShouldReturnCellValue()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var value = keyFeeder.CreateKeyFeedFromString("testtext");

            foreach (var key in value)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }


            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            var formula = keyFeeder.CreateKeyFeedFromString("=Sheet1!A1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula1 = keyFeeder.CreateKeyFeedFromString("=A1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula2 = keyFeeder.CreateKeyFeedFromString("=A2");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("testtext", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
            Assert.Equal("testtext", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
            Assert.Equal("testtext", test.CellContentAndFormulaAt(3, 1).Content.TypeValue);
        }

        [Fact]
        public void CellReferenceRecursion_ShouldReturnRecursionError()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=A1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
        }

        [Fact]
        public void CellReferenceSmallChainRecursion_ShouldReturnRecursionError()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("=A2");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            var formula1 = keyFeeder.CreateKeyFeedFromString("=A1");

            foreach (var key in formula1)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            //In he ConsoleApplication the formulas are updated during printing but as here printing
            //is out of the question, i need to hover again over the cells so the contents update and the recursion error shows.
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Up));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Down));

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(2, 1).Content.TypeValue);
        }

        [Fact]
        public void CellReferenceLargeChainRecursion_ShouldReturnRecursionError()
        {
            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            var formula = keyFeeder.CreateKeyFeedFromString("1234");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));

            formula = keyFeeder.CreateKeyFeedFromString("=A1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));

            formula = keyFeeder.CreateKeyFeedFromString("=B1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));

            formula = keyFeeder.CreateKeyFeedFromString("=C1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));

            formula = keyFeeder.CreateKeyFeedFromString("=D1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));

            Assert.Equal(1234, test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
            Assert.Equal(1234, test.CellContentAndFormulaAt(1, 2).Content.TypeValue);
            Assert.Equal(1234, test.CellContentAndFormulaAt(1, 3).Content.TypeValue);
            Assert.Equal(1234, test.CellContentAndFormulaAt(1, 4).Content.TypeValue);
            Assert.Equal(1234, test.CellContentAndFormulaAt(1, 5).Content.TypeValue);


            for (int i = 0; i < 5; i++)
            {
                test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Left));
            }

            formula = keyFeeder.CreateKeyFeedFromString("=E1");

            foreach (var key in formula)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            //In he ConsoleApplication the formulas are updated during printing but as here printing
            //is out of the question, i need to hover again over the cells so the contents update and the recursion error shows.
            for (int i = 0; i < 7; i++)
            {
                test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));
            }

            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 1).Content.TypeValue);
            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 2).Content.TypeValue);
            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 3).Content.TypeValue);
            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 4).Content.TypeValue);
            Assert.Equal("RecursErr", test.CellContentAndFormulaAt(1, 5).Content.TypeValue);
        }
    }
}
