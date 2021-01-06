using OpenMcdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace RFEM
{
    public class Program
    {
        private static readonly object Padlock = new object();
        private static readonly List<Exception> Exceptions = new List<Exception>();
        private static readonly List<NonDeletableObject> NonDeletableObjects = new List<NonDeletableObject>(); // container entries that couldnt be deleted
        private static readonly Stopwatch StopWatch = new Stopwatch();
        private static string[] _fileNames = { }; // backups deleted file count
        private static int _index; // processed file count
        private static int _backupDeleteCount; // backups deleted file count

        private class DeletableObject
        {
            public string Name { get; set; }

            public List<DeletableObject> Children { get; set; } = new List<DeletableObject>();

            public bool HasDeletableChildren // if this is true, only some children get deleted and not the parent storage itself
            {
                get { return Children.Any(); }
            }
        }

        private class NonDeletableObject
        {
            public string Path { get; set; }
            public string FileName { get; set; }
        }

        public static void Main(string[] args)
        {
            PrintConsoleHeader();
            SetFileNamesFromDirectory();
            TriggerStartExecution();
            StopWatch.Start();
            ProcessFiles();
            StopWatch.Stop();
            PrintSummary();
            PrintExitProgram();
        }

        private static void PrintConsoleHeader()
        {
            Console.SetWindowSize(155, 30);
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            string ascii = @"
════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════

██████╗ ███████╗███████╗███╗   ███╗    ██████╗ ███████╗███████╗██╗   ██╗██╗  ████████╗    ██████╗ ███████╗███╗   ███╗ ██████╗ ██╗   ██╗███████╗██████╗ 
██╔══██╗██╔════╝██╔════╝████╗ ████║    ██╔══██╗██╔════╝██╔════╝██║   ██║██║  ╚══██╔══╝    ██╔══██╗██╔════╝████╗ ████║██╔═══██╗██║   ██║██╔════╝██╔══██╗
██████╔╝█████╗  █████╗  ██╔████╔██║    ██████╔╝█████╗  ███████╗██║   ██║██║     ██║       ██████╔╝█████╗  ██╔████╔██║██║   ██║██║   ██║█████╗  ██████╔╝
██╔══██╗██╔══╝  ██╔══╝  ██║╚██╔╝██║    ██╔══██╗██╔══╝  ╚════██║██║   ██║██║     ██║       ██╔══██╗██╔══╝  ██║╚██╔╝██║██║   ██║╚██╗ ██╔╝██╔══╝  ██╔══██╗
██║  ██║██║     ███████╗██║ ╚═╝ ██║    ██║  ██║███████╗███████║╚██████╔╝███████╗██║       ██║  ██║███████╗██║ ╚═╝ ██║╚██████╔╝ ╚████╔╝ ███████╗██║  ██║
╚═╝  ╚═╝╚═╝     ╚══════╝╚═╝     ╚═╝    ╚═╝  ╚═╝╚══════╝╚══════╝ ╚═════╝ ╚══════╝╚═╝       ╚═╝  ╚═╝╚══════╝╚═╝     ╚═╝ ╚═════╝   ╚═══╝  ╚══════╝╚═╝  ╚═╝

════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════";

            Console.WriteLine(ascii);

            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);

            string authorAndVersion = "v." + fvi.FileVersion + " | Marinus S. & Stefan R.   ";
            Console.CursorLeft = Console.BufferWidth - authorAndVersion.Length;
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.Write(authorAndVersion);
            Console.WriteLine();
            Console.WriteLine();
            Console.ResetColor();
        }

        private static void SetFileNamesFromDirectory()
        {
            Console.WriteLine("[ Please enter the directory path with the .rf5 files: ]");
            Console.WriteLine();

            string path = Console.ReadLine();

            while (!Directory.Exists(path))
            {
                Console.WriteLine("[ Directory does not exist, please enter the correct path. ]");
                path = Console.ReadLine();
            }

            _fileNames = new string[] { };

            while (!_fileNames.Any())
            {
                _fileNames = Directory.GetFiles(path, "*.rf5", SearchOption.AllDirectories);

                if (_fileNames.Any())
                    break;

                Console.WriteLine("Directory does not contain any .rf5 files, please enter the correct path.");
                path = Console.ReadLine();
            }
        }

        private static void TriggerStartExecution()
        {
            Console.WriteLine();
            Console.WriteLine($"[ {_fileNames.Length} files found. Are you sure you want to start the conversion? Type 'start' to execute. ]");

            // blocking start wait
            while (Console.ReadLine()?.ToLower() != "start")
            {
            }

            Console.WriteLine($"[ Program started. Converting files... ]");
            Console.CursorVisible = false;
        }

        private static void ProcessFiles()
        {
            // process loop
            Parallel.ForEach(_fileNames, (fileName) =>
            {
                try
                {
                    RemoveObsoleteObjectsFromFile(fileName);

                    if (RemoveIndividualBackup(fileName))
                        Interlocked.Increment(ref _backupDeleteCount);
                }
                catch (Exception ex)
                {
                    ex.Data["fileName"] = fileName;

                    lock (Padlock)
                    {
                        Exceptions.Add(ex);
                    }
                }

                Interlocked.Increment(ref _index);

                lock (Padlock)
                {
                    int percentage = _index * 100 / _fileNames.Length;
                    ClearCurrentConsoleLine();
                    Console.Write($"[ {_index}/{_fileNames.Length} processed [{percentage}%] ]");
                }
            });
        }

        private static void PrintSummary()
        {
            // Print info
            Console.CursorVisible = true;
            Console.WriteLine();
            Console.WriteLine($"[ Program finished. ]");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"[ Successfully converted {(_index - Exceptions.Count)} out of {_fileNames.Length} files, with {Exceptions.Count} errors (listed below). ]");
            Console.WriteLine($"[ Successfully removed {_backupDeleteCount} backup files ]");
            Console.WriteLine();
            Console.WriteLine($"Total time used: {StopWatch.Elapsed.ToString(@"hh\:mm\:ss")}");
            Console.ResetColor();

            Console.WriteLine();

            // write errors
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[ {Exceptions.Count} ERROR(s) found: ]");
            Console.WriteLine($"[ These files have not been processed ]");
            Console.WriteLine();

            foreach (var exception in Exceptions)
            {
                Console.WriteLine("----------------------------------------------");
                Console.WriteLine();
                Console.WriteLine($"File Path:");
                Console.WriteLine($"{exception.Data["fileName"]}");
                if (exception.Data["corruptedEntryName"] != null)
                {
                    Console.WriteLine($"Corrupted Entry Name:");
                    Console.WriteLine($"{exception.Data["corruptedEntryName"]}");
                }
                Console.WriteLine();
                Console.WriteLine($"StackTrace:");
                Console.WriteLine();
                Console.WriteLine(exception.ToString());
                Console.WriteLine();
                Console.WriteLine();
            }

            Console.ResetColor();

            Console.WriteLine();
        }

        private static void PrintExitProgram()
        {
            Console.WriteLine($"[ Press any key to exit the program... ]");
            Console.ReadKey();
        }

        private static void ClearCurrentConsoleLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            for (int i = 0; i < Console.WindowWidth; i++)
                Console.Write(" ");
            Console.SetCursorPosition(0, currentLineCursor);
        }

        private static bool RemoveIndividualBackup(string path)
        {
            var backupPath = path + "bak";

            if (!File.Exists(backupPath))
                return false;

            File.Delete(backupPath);
            return true;
        }

        private static void RemoveObsoleteObjectsFromFile(string path)
        {
            using var compoundFile = new CompoundFile(path); //, CFSUpdateMode.ReadOnly,CFSConfiguration.SectorRecycle | CFSConfiguration.EraseFreeSectors);

            if (compoundFile.RootStorage == null)
                throw new InvalidOperationException($"RootStorage null for compoundFile: {path}");

            var deletableObjects = FindObjectsToDeleteRecursive(compoundFile.RootStorage);

            RemoveStorageRecursive(compoundFile.RootStorage, deletableObjects);

            using var newCompound = CopyCompoundFile(path, compoundFile);

            newCompound.Save(path);
        }

        private static List<DeletableObject> FindObjectsToDeleteRecursive(CFStorage storage)
        {
            var deletableObjects = new List<DeletableObject>();

            if (storage == null)
                return deletableObjects;

            storage.VisitEntries((item) =>
            {
                var itemName = item.Name;

                if (itemName.StartsWith("Results") ||
                    itemName.StartsWith("Protocols") ||
                    itemName.StartsWith("Mesh") ||
                    itemName.StartsWith("Exported Results") ||
                    itemName.StartsWith("STB") ||
                    itemName.StartsWith("STM"))
                {
                    //Console.WriteLine($"Found deletable entry {itemName}");

                    var deletableObject = new DeletableObject();
                    deletableObject.Name = itemName;

                    deletableObjects.Add(deletableObject);
                }
                else if (item.IsStorage)
                {
                    var deletableObject = new DeletableObject();
                    deletableObject.Name = itemName;
                    deletableObject.Children = FindObjectsToDeleteRecursive(item as CFStorage);

                    if (deletableObject.HasDeletableChildren)
                        deletableObjects.Add(deletableObject);
                }

            }, false);

            return deletableObjects;
        }

        private static void RemoveStorageRecursive(CFStorage storage, List<DeletableObject> deletableObjects)
        {
            foreach (var deletableObject in deletableObjects)
            {
                if (deletableObject.HasDeletableChildren && // if its a parent, only some children get removed and not the parent itself
                    storage.TryGetStorage(deletableObject.Name, out var newStorage))
                {
                    RemoveStorageRecursive(newStorage, deletableObject.Children);
                }
                else
                {
                    storage.Delete(deletableObject.Name);
                }
            }
        }

        private static CompoundFile CopyCompoundFile(string path, CompoundFile compoundFile)
        {
            var newCompoundFile = new CompoundFile(compoundFile.Version, compoundFile.Configuration);

            //Copy Root CLSID
            newCompoundFile.RootStorage.CLSID = new Guid(compoundFile.RootStorage.CLSID.ToByteArray());

            CopyCFStorageRecursive(path, compoundFile.RootStorage, newCompoundFile.RootStorage);

            return newCompoundFile;
        }

        private static void CopyCFStorageRecursive(string path, CFStorage sourceStorage, CFStorage destinationStorage)
        {
            void CopyAction(CFItem item)
            {
                if (item.IsStream)
                {
                    var itemAsStorage = item as CFStream;

                    try
                    {
                        var data = itemAsStorage.GetData();
                        CFStream stream = destinationStorage.AddStream(itemAsStorage.Name);
                        stream.SetData(data);
                    }
                    catch (CFCorruptedFileException ex)
                    {
                        ex.Data["corruptedEntryName"] = itemAsStorage.Name;
                        throw;
                    }
                }
                else if (item.IsStorage)
                {
                    var itemAsStorage = item as CFStorage;

                    CFStorage newStorage = destinationStorage.AddStorage(itemAsStorage.Name);
                    newStorage.CLSID = new Guid(itemAsStorage.CLSID.ToByteArray());
                    CopyCFStorageRecursive(path, itemAsStorage, newStorage); // recursion, one level deeper
                }
            }

            sourceStorage.VisitEntries(CopyAction, false);
        }
    }
}
