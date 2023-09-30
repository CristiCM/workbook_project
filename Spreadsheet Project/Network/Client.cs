using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Sockets;
using Newtonsoft.Json;
using HonkSharp.Fluency;
using OfficeOpenXml;

namespace Spreadsheet_Project.Network
{
    public class Client
    {
        Socket clientSocket;
        Dictionary<int, object> locks = new();
        List<(string sheetName, Sheet sheet)> sheets;
        int currentSheetIndex;
        ApplicationInitializer appInitializer;
        bool movement = false;

        int[] movementKeys = new int[]
        {
            (int)ConsoleKey.LeftArrow,
            (int)ConsoleKey.RightArrow,
            (int)ConsoleKey.DownArrow,
            (int)ConsoleKey.UpArrow,
            (int)ConsoleKey.Tab,
            (int)ConsoleKey.Enter
        };

        public Client(List<(string sheetName, Sheet sheet)> sheets, int currentSheetIndex, ApplicationInitializer appInitializer)
        {
            this.sheets = sheets;
            this.currentSheetIndex = currentSheetIndex;
            this.appInitializer = appInitializer;
        }

        public async Task Initialize(string ipAdress, int port)
        {
            clientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            IPEndPoint serverEndPoint = new IPEndPoint(IPAddress.Parse(ipAdress), port);

            Console.Clear();

            Console.WriteLine("Connecting to the server...");

            Thread.Sleep(1000);

            await clientSocket.ConnectAsync(serverEndPoint);

            await ReceiveInitialDictionary();

            await RegisterAndExecuteKeysFromServerAsync();
        }

        private async Task ReceiveInitialDictionary()
        {
            byte[] buffer = new byte[8192];

            Thread.Sleep(5000);
            int bytesRead = await clientSocket.ReceiveAsync(buffer, SocketFlags.None);

            using (var stream = new MemoryStream(buffer, 0, bytesRead))
            {
                ExcelPackage package = new ExcelPackage(stream);

                appInitializer.OpenWorkbook(package);
            }

            sheets[currentSheetIndex].sheet.printHandler.PrintSheet();
        }

        private async Task RegisterAndExecuteKeysFromServerAsync()
        {
            while (true)
            {
                await HandleKeyPressFromServerAsync();
            }

        }

        public async Task HandleKeyPressFromClientLocalConsoleAsync(ConsoleKeyInfo pressedKey)
        {
            await ProcessKeyAsync(pressedKey, currentSheetIndex, sheets[currentSheetIndex].sheet.currentPosition);

            string json = ConsoleKeyDataTransferObject.Serialize(pressedKey, currentSheetIndex, sheets[currentSheetIndex].sheet.currentPosition);

            byte[] responseBytes = Encoding.ASCII.GetBytes(json);
            
            await clientSocket.SendAsync(responseBytes, SocketFlags.None);
        }

        private async Task HandleKeyPressFromServerAsync()
        {
            try
            {
                byte[] buffer = new byte[1024];

                int bytesRead = await clientSocket.ReceiveAsync(buffer, SocketFlags.None);
                if (bytesRead == 0)
                    return;

                string jsonString = Encoding.UTF8.GetString(buffer, 0, bytesRead);

                (ConsoleKeyInfo keyInfo, int correspondingSheetIndex, (int, int) correspondingCurrentPosition) = ConsoleKeyDataTransferObject.Deserialize(jsonString);

                await ProcessKeyAsync(keyInfo, correspondingSheetIndex, correspondingCurrentPosition);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

        }

        private async Task ProcessKeyAsync(ConsoleKeyInfo pressedKey, int correspondingSheetIndex, (int, int) correspondingCurrentPosition)
        {
            if (!movementKeys.Contains((int)pressedKey.Key) && movement)
                await SendCellDeletionInstructionsToServerDueToMovement();
            

            lock (locks)
            {
                if (!locks.ContainsKey(correspondingSheetIndex))
                    locks.Add(correspondingSheetIndex, correspondingCurrentPosition);
            }

            lock (locks[correspondingSheetIndex])
            {
                movement = movementKeys.Contains((int)pressedKey.Key);
                
                var oldCurrentPosition = sheets[correspondingSheetIndex].sheet.currentPosition;

                sheets[correspondingSheetIndex].sheet.currentPosition = correspondingCurrentPosition;

                sheets[correspondingSheetIndex].sheet.Execute(pressedKey);

                if (!movementKeys.Contains((int)pressedKey.Key))
                {
                    sheets[correspondingSheetIndex].sheet.currentPosition = oldCurrentPosition;
                }
                    
                sheets[correspondingSheetIndex].sheet.printHandler.PrintSheet();
            }
        }

        private async Task SendCellDeletionInstructionsToServerDueToMovement()
        {
            ConsoleKeyInfo deleteKeyInfo = new((char)0, ConsoleKey.Delete, false, false, false);

            string json = ConsoleKeyDataTransferObject.Serialize(deleteKeyInfo, currentSheetIndex, sheets[currentSheetIndex].sheet.currentPosition);

            byte[] responseBytes = Encoding.ASCII.GetBytes(json);

            await clientSocket.SendAsync(responseBytes, SocketFlags.None);
        }
    }
}
