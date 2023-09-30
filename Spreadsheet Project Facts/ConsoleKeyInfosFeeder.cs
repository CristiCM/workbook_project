namespace Spreadsheet_Project_Facts
{
    public class ConsoleKeyInfosFeeder
    {
        readonly string[][] initailData = new[]
        {
            new [] { "user", "stake", "profit", "roi" },
            new [] { "bob", "100", "200", "1" },
            new [] { "sam", "200", "400", "2" },
            new [] { "tom", "300", "600", "3" },
            new [] { "bill", "400", "800", "4" },
            new [] { "bob", "500", "1000", "5" },
            new [] { "sam", "600", "1200", "6" },
            new [] { "tom", "700", "1400", "7" },
        };

        public enum ExecuteKeys
        {
            Left,
            Right,
            Up,
            Down,
            AddSheet,
            NextSheet,
            PrevSheet,
            RenameSheet,
            DeleteSheet,
            Delete,
            SaveSheet,
            OpenSheet,
            NewSheet,
            Backspace,
            Cut,
            Copy,
            Paste,
            F2
        }



        public ConsoleKeyInfo GetExecuteKey(ExecuteKeys direction)
        {
            return direction switch
            {
                ExecuteKeys.Left => new ConsoleKeyInfo('\0', ConsoleKey.LeftArrow, false, false, false),
                ExecuteKeys.Right => new ConsoleKeyInfo('\0', ConsoleKey.RightArrow, false, false, false),
                ExecuteKeys.Up => new ConsoleKeyInfo('\0', ConsoleKey.UpArrow, false, false, false),
                ExecuteKeys.Down => new ConsoleKeyInfo('\0', ConsoleKey.DownArrow, false, false, false),
                ExecuteKeys.AddSheet => new ConsoleKeyInfo((char)0, ConsoleKey.F5, false, false, false),
                ExecuteKeys.NextSheet => new ConsoleKeyInfo((char)0, ConsoleKey.F7, false, false, false),
                ExecuteKeys.PrevSheet => new ConsoleKeyInfo((char)0, ConsoleKey.F6, false, false, false),
                ExecuteKeys.RenameSheet => new ConsoleKeyInfo((char)0, ConsoleKey.F9, false, false, false),
                ExecuteKeys.DeleteSheet => new ConsoleKeyInfo((char)0, ConsoleKey.F8, false, false, false),
                ExecuteKeys.Delete => new ConsoleKeyInfo('\0', ConsoleKey.Delete, false, false, false),
                ExecuteKeys.SaveSheet => new ConsoleKeyInfo((char)0, ConsoleKey.S, true, false, false),
                ExecuteKeys.OpenSheet => new ConsoleKeyInfo((char)0, ConsoleKey.O, true, false, false),
                ExecuteKeys.NewSheet => new ConsoleKeyInfo('\u000e', ConsoleKey.N, false, false, true),
                ExecuteKeys.Backspace => new ConsoleKeyInfo((char)0, ConsoleKey.Backspace, false, false, false),
                ExecuteKeys.Cut => new ConsoleKeyInfo('\u0018', ConsoleKey.X, false, false, true),
                ExecuteKeys.Copy => new ConsoleKeyInfo('\u0003', ConsoleKey.C, false, false, true),
                ExecuteKeys.Paste => new ConsoleKeyInfo('\u0016', ConsoleKey.V, false, false, true),
                ExecuteKeys.F2 => new ConsoleKeyInfo((char)0, ConsoleKey.F2, false, false, false),
                _ => throw new ArgumentException("Invalid direction."),
            };
        }

        private ConsoleKey GetConsoleKeyFromChar(char inputChar)
        {
            inputChar = char.ToLower(inputChar);

            Dictionary<char, ConsoleKey> characterToKeyMapping = new Dictionary<char, ConsoleKey>
            {
                { 'a', ConsoleKey.A },
                { 'b', ConsoleKey.B },
                { 'c', ConsoleKey.C },
                { 'd', ConsoleKey.D },
                { 'e', ConsoleKey.E },
                { 'f', ConsoleKey.F },
                { 'g', ConsoleKey.G },
                { 'h', ConsoleKey.H },
                { 'i', ConsoleKey.I },
                { 'j', ConsoleKey.J },
                { 'k', ConsoleKey.K },
                { 'l', ConsoleKey.L },
                { 'm', ConsoleKey.M },
                { 'n', ConsoleKey.N },
                { 'o', ConsoleKey.O },
                { 'p', ConsoleKey.P },
                { 'q', ConsoleKey.Q },
                { 'r', ConsoleKey.R },
                { 's', ConsoleKey.S },
                { 't', ConsoleKey.T },
                { 'u', ConsoleKey.U },
                { 'v', ConsoleKey.V },
                { 'w', ConsoleKey.W },
                { 'x', ConsoleKey.X },
                { 'y', ConsoleKey.Y },
                { 'z', ConsoleKey.Z },
                { '=', ConsoleKey.OemPlus },
                { '0', ConsoleKey.D0 },
                { '1', ConsoleKey.D1 },
                { '2', ConsoleKey.D2 },
                { '3', ConsoleKey.D3 },
                { '4', ConsoleKey.D4 },
                { '5', ConsoleKey.D5 },
                { '6', ConsoleKey.D6 },
                { '7', ConsoleKey.D7 },
                { '8', ConsoleKey.D8 },
                { '9', ConsoleKey.D9 },
                { ' ', ConsoleKey.Spacebar },
                { '(', ConsoleKey.D9 },
                { ')', ConsoleKey.D0 },
                { '"', ConsoleKey.Oem7 },
                { ':', ConsoleKey.Oem1 },
                { ',', ConsoleKey.OemComma }
            };


            if (characterToKeyMapping.ContainsKey(inputChar))
            {
                return characterToKeyMapping[inputChar];
            }

            return ConsoleKey.NoName;
        }


        public List<ConsoleKeyInfo> CreateInitialDataKeyFeed(string[][] cells)
        {
            List<ConsoleKeyInfo> dataConsoleKeyInfos = new();

            foreach (var cellRow in cells)
            {
                foreach (var cell in cellRow)
                {
                    foreach (var character in cell)
                    {
                        dataConsoleKeyInfos.Add(new ConsoleKeyInfo(character, GetConsoleKeyFromChar(character), false, false, false));
                    }

                    dataConsoleKeyInfos.Add(GetExecuteKey(ExecuteKeys.Right));
                }


                dataConsoleKeyInfos.Add(GetExecuteKey(ExecuteKeys.Down));
                for (int i = 0; i < cellRow.Length; i++)
                {
                    dataConsoleKeyInfos.Add(GetExecuteKey(ExecuteKeys.Left));
                }
            }

            return dataConsoleKeyInfos;
        }

        public List<ConsoleKeyInfo> CreateKeyFeedFromString(string stringInput)
        {
            List<ConsoleKeyInfo> dataConsoleKeyInfos = new();
            
            foreach (var character in stringInput)
            {
                dataConsoleKeyInfos.Add(new ConsoleKeyInfo(character, GetConsoleKeyFromChar(character), false, false, false));
            }

            return dataConsoleKeyInfos;
        }

        public List<ConsoleKeyInfo> GetStandardDataConsoleKeyInfosList()
        {
            return CreateInitialDataKeyFeed(initailData);
        }

        public ConsoleKeyInfo Escape()
        {
            return new ConsoleKeyInfo((char)0, ConsoleKey.Escape, false, false, false);
        }
    }
}
