using Spreadsheet_Project;
using static Spreadsheet_Project_Facts.ConsoleKeyInfosFeeder;

namespace Spreadsheet_Project_Facts
{
    public class PivotTableFacts
    {
        ConsoleKeyInfosFeeder keyFeeder = new();

        [Fact]
        public void PivotTableSumFunction_ShouldCreateAPivotTableAtTheLocation()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.CreatePivotTableTesting("A1:D8", "G1", "user", new List<(int, string)>()
            {
                (0, "stake"),
            });

            string[][] resultedPivot = new[] {
                new [] { "user", "stake"},
                new [] { "bill", "400"},
                new [] {"bob", "600" },
                new [] {"sam", "800" },
                new [] {"tom", "1000" },
                new [] {"Total", "2800" }
            };

            for (int j = 0; j < resultedPivot.Length; j++)
            {
                for (int i = 0; i < resultedPivot[i].Length; i++)
                {
                    Assert.Equal(resultedPivot[j][i], test.CellContentAndFormulaAt(j + 1, i + 7).Content.TypeValue.ToString());
                }
            }
        }

        [Fact]
        public void PivotTableAvgFunction_ShouldCreateAPivotTableAtTheLocation()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.CreatePivotTableTesting("A1:D8", "G1", "user", new List<(int, string)>()
            {
                (1, "profit"),
            });

            string[][] resultedPivot = new[] {
                new [] { "user", "profit" },
                new [] { "bill", "800" },
                new [] {"bob", "600" },
                new [] {"sam", "800" },
                new [] {"tom", "1000" },
                new [] {"Total", "800" }
            };

            for (int j = 0; j < resultedPivot.Length; j++)
            {
                for (int i = 0; i < resultedPivot[i].Length; i++)
                {
                    Assert.Equal(resultedPivot[j][i], test.CellContentAndFormulaAt(j + 1, i + 7).Content.TypeValue.ToString());
                }
            }
        }

        [Fact]
        public void PivotTableCountFunction_ShouldCreateAPivotTableAtTheLocation()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.CreatePivotTableTesting("A1:D8", "G1", "user", new List<(int, string)>()
            {
                (2, "roi"),
            });

            string[][] resultedPivot = new[] {
                new [] { "user", "roi" },
                new [] { "bill", "1" },
                new [] {"bob", "2" },
                new [] {"sam", "2" },
                new [] {"tom", "2" },
                new [] {"Total", "7" }
            };

            for (int j = 0; j < resultedPivot.Length; j++)
            {
                for (int i = 0; i < resultedPivot[i].Length; i++)
                {
                    Assert.Equal(resultedPivot[j][i], test.CellContentAndFormulaAt(j + 1, i + 7).Content.TypeValue.ToString());
                }
            }
        }

        [Fact]
        public void PivotTableAllFunction_ShouldCreateAPivotTableAtTheLocation()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.CreatePivotTableTesting("A1:D8", "G1", "user", new List<(int, string)>()
            {
                (0, "stake"),
                (1, "profit"),
                (2, "roi"),
            });

            string[][] resultedPivot = new[] {
                new [] { "user", "stake", "profit", "roi" },
                new [] { "bill", "400", "800", "1" },
                new [] { "bob", "600", "600", "2" },
                new [] { "sam", "800", "800", "2" },
                new [] { "tom", "1000", "1000", "2" },
                new [] { "Total", "2800", "800", "7" }
            };

            for (int j = 0; j < resultedPivot.Length; j++)
            {
                for (int i = 0; i < resultedPivot[i].Length; i++)
                {
                    Assert.Equal(resultedPivot[j][i], test.CellContentAndFormulaAt(j + 1, i + 7).Content.TypeValue.ToString());
                }
            }
        }

        [Fact]
        public void PivotTableAllFunctionCrossSheet_ShouldCreateAPivotTableOnCurrentSheetWithDataFromAnotherSheet()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));

            test.CreatePivotTableTesting("Sheet1!A1:D8", "G1", "user", new List<(int, string)>()
            {
                (0, "stake"),
                (1, "profit"),
                (2, "roi"),
            });

            string[][] resultedPivot = new[] {
                new [] { "user", "stake", "profit", "roi" },
                new [] { "bill", "400", "800", "1" },
                new [] { "bob", "600", "600", "2" },
                new [] { "sam", "800", "800", "2" },
                new [] { "tom", "1000", "1000", "2" },
                new [] { "Total", "2800", "800", "7" }
            };

            for (int j = 0; j < resultedPivot.Length; j++)
            {
                for (int i = 0; i < resultedPivot[i].Length; i++)
                {
                    Assert.Equal(resultedPivot[j][i], test.CellContentAndFormulaAt(j + 1, i + 7).Content.TypeValue.ToString());
                }
            }
        }

        [Fact]
        public void PivotTableAllFunctionCrossSheet_ShouldCreateAPivotTableOnAnotherSheetWithDataFromCurrentSheet()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.AddSheet));
            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.PrevSheet));

            test.CreatePivotTableTesting("A1:D8", "Sheet2!G1", "user", new List<(int, string)>()
            {
                (0, "stake"),
                (1, "profit"),
                (2, "roi"),
            });

            string[][] resultedPivot = new[] {
                new [] { "user", "stake", "profit", "roi" },
                new [] { "bill", "400", "800", "1" },
                new [] { "bob", "600", "600", "2" },
                new [] { "sam", "800", "800", "2" },
                new [] { "tom", "1000", "1000", "2" },
                new [] { "Total", "2800", "800", "7" }
            };

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.NextSheet));

            for (int j = 0; j < resultedPivot.Length; j++)
            {
                for (int i = 0; i < resultedPivot[i].Length; i++)
                {
                    Assert.Equal(resultedPivot[j][i], test.CellContentAndFormulaAt(j + 1, i + 7).Content.TypeValue.ToString());
                }
            }
        }

        [Fact]
        public void PivotTableCreationOnTopOfExistingPivot_ShouldDeleteTheOldPivotAndCreateTheNewOne()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            //Initial Pivot
            test.CreatePivotTableTesting("A1:D8", "G1", "user", new List<(int, string)>()
            {
                (2, "roi"),
            });

            string[][] resultInitialPivot = new[] {
                new [] { "user", "roi" },
                new [] { "bill", "1" },
                new [] {"bob", "2" },
                new [] {"sam", "2" },
                new [] {"tom", "2" },
                new [] {"Total", "7" }
            };

            for (int j = 0; j < resultInitialPivot.Length; j++)
            {
                for (int i = 0; i < resultInitialPivot[i].Length; i++)
                {
                    Assert.Equal(resultInitialPivot[j][i], test.CellContentAndFormulaAt(j + 1, i + 7).Content.TypeValue.ToString());
                }
            }

            //Folowing Pivot
            test.CreatePivotTableTesting("A1:D8", "G1", "user", new List<(int, string)>()
            {
                (0, "stake"),
                (1, "profit"),
                (2, "roi"),
            });

            string[][] resultSecondPivot = new[] {
                new [] { "user", "stake", "profit", "roi" },
                new [] { "bill", "400", "800", "1" },
                new [] { "bob", "600", "600", "2" },
                new [] { "sam", "800", "800", "2" },
                new [] { "tom", "1000", "1000", "2" },
                new [] { "Total", "2800", "800", "7" }
            };

            for (int j = 0; j < resultSecondPivot.Length; j++)
            {
                for (int i = 0; i < resultSecondPivot[i].Length; i++)
                {
                    Assert.Equal(resultSecondPivot[j][i], test.CellContentAndFormulaAt(j + 1, i + 7).Content.TypeValue.ToString());
                }
            }
        }

        [Fact]
        public void PivotTableDeleteKeyOnPivotLocation_ShouldDeleteThePivotTable()
        {
            var initialDataKeys = keyFeeder.GetStandardDataConsoleKeyInfosList();

            ApplicationInitializer test = new();

            test.TestingRunNewSheet();

            foreach (var key in initialDataKeys)
            {
                test.RegisterAndExecuteActionsTesting(key);
            }

            test.CreatePivotTableTesting("A1:D8", "G1", "user", new List<(int, string)>()
            {
                (0, "stake"),
                (1, "profit"),
                (2, "roi"),
            });

            string[][] resultPivot = new[] {
                new [] { "user", "stake", "profit", "roi" },
                new [] { "bill", "400", "800", "1" },
                new [] { "bob", "600", "600", "2" },
                new [] { "sam", "800", "800", "2" },
                new [] { "tom", "1000", "1000", "2" },
                new [] { "Total", "2800", "800", "7" }
            };

            for (int j = 0; j < resultPivot.Length; j++)
            {
                for (int i = 0; i < resultPivot[i].Length; i++)
                {
                    Assert.Equal(resultPivot[j][i], test.CellContentAndFormulaAt(j + 1, i + 7).Content.TypeValue.ToString());
                }
            }

            while (test.GetCurrentPosition() != (1, 7))
            {
                if (test.GetCurrentPosition().X > 1)
                    test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Up));

                if (test.GetCurrentPosition().Y < 7)
                    test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Right));
            }

            test.RegisterAndExecuteActionsTesting(keyFeeder.GetExecuteKey(ExecuteKeys.Delete));

            for (int j = 0; j < resultPivot.Length; j++)
            {
                for (int i = 0; i < resultPivot[i].Length; i++)
                {
                    Assert.Equal((null, null), test.CellContentAndFormulaAt(j + 1, i + 7));
                }
            }
        }
    }
}
