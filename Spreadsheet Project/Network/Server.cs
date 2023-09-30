using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Newtonsoft.Json;
using OfficeOpenXml;
using static AngouriMath.Entity;

namespace Spreadsheet_Project.Network
{
    //TODO: An existing connection was forcibly closed by the remote host.

    public class Server
    {
        Socket serverSocket;
        Dictionary<string, (Socket clientSocket, (int, int) currentPosition, int currentSheetIndex)> clientSockets;
        Dictionary<int, object> locks = new();
        List<(string sheetName, Sheet sheet)> sheets;
        int currentSheetIndex;
        ApplicationInitializer appInitializer;
        bool movement = false;
        bool clientConnected = false;

        int[] movementKeys = new int[]
            {
                (int)ConsoleKey.LeftArrow,
                (int)ConsoleKey.RightArrow,
                (int)ConsoleKey.DownArrow,
                (int)ConsoleKey.UpArrow,
                (int)ConsoleKey.Tab,
                (int)ConsoleKey.Enter
            };

        public Server(List<(string sheetName, Sheet sheet)> sheets, int currentSheetIndex, ApplicationInitializer appInitializer)
        {
            this.sheets = sheets;
            this.currentSheetIndex = currentSheetIndex;
            this.appInitializer = appInitializer;
        }

        public async Task Initialize()
        {
            serverSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            IPAddress? ipAdress = IPAddress.Any;
            IPEndPoint? iPEndPoint = new(ipAdress, 8080);
            serverSocket.Bind(iPEndPoint);
            serverSocket.Listen();

            clientSockets = new();

            Console.Clear();

            Console.WriteLine("Starting server, waiting for someone to connect...");

            Task acceptClientsTask = AcceptClientsAsync();

            Task handleKeysAsyncTask = HandleKeysAsync();

            await Task.WhenAll(acceptClientsTask, handleKeysAsyncTask);
        }

        private async Task AcceptClientsAsync()
        {
            while (true)
            {
                Socket clientSocket = await serverSocket.AcceptAsync();
                string identifier = Guid.NewGuid().ToString();
                clientSockets.Add(identifier, (clientSocket, (1, 1), 0));
                await SendInitialDictionaryAndReceiveCurrentPosition(clientSocket, identifier);

                clientConnected = true;
            }
        }

        public async Task HandleKeysAsync()
        {
            List<Task> clientTasks = new List<Task>();

            while (!clientConnected)
            {
                if (clientConnected)
                {
                    break;
                }
            }

            while (true)
            {
                foreach (var clientEntry in clientSockets)
                {
                    var clientTask = HandleKeyPressFromClientsAsync(clientEntry.Value.clientSocket, clientEntry.Key);

                    if (!clientTasks.Contains(clientTask))
                        clientTasks.Add(clientTask);
                }

                await Task.WhenAll(clientTasks);
            }
        }


        public async Task SendInitialDictionaryAndReceiveCurrentPosition(Socket clientSocket, string identifier)
        {
            var package = appInitializer.SaveWorkbook(networkAction: true);

            using (var stream = new MemoryStream())
            {
                package.SaveAs(stream);

                stream.Seek(0, SeekOrigin.Begin);

                byte[] buffer = stream.ToArray();
                await clientSocket.SendAsync(buffer, SocketFlags.None);
            }
        }

        public async Task HandleKeyPressFromServerLocalConsoleAsync(ConsoleKeyInfo presedKey)
        {
            await ProcessKeyAsync(presedKey, currentSheetIndex, sheets[currentSheetIndex].sheet.currentPosition);

            if (!movementKeys.Contains((int)presedKey.Key))
            {
                string json = ConsoleKeyDataTransferObject.Serialize(presedKey, currentSheetIndex, sheets[currentSheetIndex].sheet.currentPosition);

                byte[] responseBytes = Encoding.ASCII.GetBytes(json);

                foreach (var client in clientSockets)
                {
                    await client.Value.clientSocket.SendAsync(responseBytes, SocketFlags.None);
                }
            }
        }

        private async Task HandleKeyPressFromClientsAsync(Socket clientSocket, string identifier)
        {
            try
            {
                byte[] buffer = new byte[1024];
                int bytesRead = await clientSocket.ReceiveAsync(buffer, SocketFlags.None);
                if (bytesRead == 0)
                    return;

                string jsonString = Encoding.UTF8.GetString(buffer);

                (ConsoleKeyInfo keyInfo, int correspondingSheetIndex, (int, int) correspondingCurrentPosition) =
                    ConsoleKeyDataTransferObject.Deserialize(jsonString);

                if (!movementKeys.Contains((int)keyInfo.Key))
                {
                    await ProcessKeyAsync(keyInfo, correspondingSheetIndex, correspondingCurrentPosition);

                    string json = ConsoleKeyDataTransferObject.Serialize(keyInfo, correspondingSheetIndex, correspondingCurrentPosition);
                    byte[] responseBytes = Encoding.ASCII.GetBytes(json);

                    foreach (var client in clientSockets)
                    {
                        if (client.Key != identifier)
                        {
                            await client.Value.clientSocket.SendAsync(responseBytes, SocketFlags.None);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

        }

        private async Task ProcessKeyAsync(ConsoleKeyInfo pressedKey, int correspondingSheetIndex, (int, int) correspondingCurrentPosition)
        {
            if (!movementKeys.Contains((int)pressedKey.Key) && movement)
                await SendCellDeletionInstructionsToClientsDueToMovement();


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

        private async Task SendCellDeletionInstructionsToClientsDueToMovement()
        {
            ConsoleKeyInfo deleteKeyInfo = new((char)0, ConsoleKey.Delete, false, false, false);

            string json = ConsoleKeyDataTransferObject.Serialize(deleteKeyInfo, currentSheetIndex, sheets[currentSheetIndex].sheet.currentPosition);

            byte[] responseBytes = Encoding.ASCII.GetBytes(json);

            foreach (var item in clientSockets)
            {
                await item.Value.clientSocket.SendAsync(responseBytes, SocketFlags.None);
            }
        }
    }
}
