using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spreadsheet_Project.Network
{
    public class ConsoleKeyDataTransferObject
    {
        public char KeyChar { get; set; }
        public int Key { get; set; }
        public bool Shift { get; set; }
        public bool Alt { get; set; }
        public bool Control { get; set; }
        public int CurrentSheetIndex { get; set; }
        public (int, int) CurrentPosition { get; set; }

        public static string Serialize(ConsoleKeyInfo consoleKeyInfo, int currentSheetIndex, (int, int) currentPosition)
        {
            var dto = new ConsoleKeyDataTransferObject
            {
                KeyChar = consoleKeyInfo.KeyChar,
                Key = (int)consoleKeyInfo.Key,
                Shift = (consoleKeyInfo.Modifiers & ConsoleModifiers.Shift) != 0,
                Alt = (consoleKeyInfo.Modifiers & ConsoleModifiers.Alt) != 0,
                Control = (consoleKeyInfo.Modifiers & ConsoleModifiers.Control) != 0,
                CurrentSheetIndex = currentSheetIndex,
                CurrentPosition = currentPosition
            };

            return JsonConvert.SerializeObject(dto);
        }

        public static (ConsoleKeyInfo, int, (int, int)) Deserialize(string json)
        {
            var dto = JsonConvert.DeserializeObject<ConsoleKeyDataTransferObject>(json);

            ConsoleModifiers modifiers = 0;
            if (dto.Shift) modifiers |= ConsoleModifiers.Shift;
            if (dto.Alt) modifiers |= ConsoleModifiers.Alt;
            if (dto.Control) modifiers |= ConsoleModifiers.Control;

            return (new ConsoleKeyInfo(dto.KeyChar, (ConsoleKey)dto.Key, dto.Shift, dto.Alt, dto.Control), dto.CurrentSheetIndex, dto.CurrentPosition);
        }
    }
}
