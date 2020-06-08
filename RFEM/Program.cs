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

        private class DeletableObject
        {
            public string Name { get; set; }
            
            public List<DeletableObject> Children { get; set; } = new List<DeletableObject>();

            public bool HasDeletableChildren // if this is true, only some children get deleted and not the parent storage itself
            {
                get { return Children.Any(); }
            }
        }

        public static void Main(string[] args)
        {
            try
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
                //string path = Directory.GetCurrentDirectory();
                Console.ResetColor();


                Console.WriteLine("[ Please enter the directory path with the .rf5 files: ]");
                Console.WriteLine();

                string path = Console.ReadLine();

                while (!Directory.Exists(path))
                {
                    Console.WriteLine("[ Directory does not exist, please enter the correct path. ]");
                    path = Console.ReadLine();
                }

                string[] fileNames = new string[] { };

                while (!fileNames.Any())
                {
                    fileNames = Directory.GetFiles(path, "*.rf5", SearchOption.AllDirectories);

                    if (fileNames.Any())
                        break;

                    Console.WriteLine("Directory does not contain any .rf5 files, please enter the correct path.");
                    path = Console.ReadLine();
                }

                int fileCount = fileNames.Length;

                Console.WriteLine();
                Console.WriteLine($"[ {fileCount} files found. Press any key to start the conversion... ]");
                Console.ReadKey();
                Console.WriteLine($"[ Program started. Converting files... ]");
                Console.CursorVisible = false;
                int index = 0;
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();
                Parallel.ForEach(fileNames, (fileName) =>
                {
                    try
                    {
                        RemoveObsoleteObjectsFromFile(fileName);
                        Interlocked.Increment(ref index);

                        lock (Padlock)
                        {
                            int percentage = index * 100 / fileCount;
                            ClearCurrentConsoleLine();
                            Console.Write($"[ {index}/{fileCount} processed [{percentage}%] ]");
                            //Console.Write($"{percentage}% [{index}/{fileCount}] {fileName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        ex.Data["fileName"] = fileName;
                        ex.Data["message"] = ex.Message;
                        ex.Data["stackTrace"] = ex.StackTrace;
                        throw ex;
                    }
                });

                //ClearCurrentConsoleLine();
                //Console.SetCursorPosition(0, Console.CursorTop - 1);
                //ClearCurrentConsoleLine();

                stopWatch.Stop();
                TimeSpan timeSpan = stopWatch.Elapsed;
                Console.CursorVisible = true;
                Console.WriteLine();
                Console.WriteLine($"[ Program finished. ]");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"[ Successfully converted {fileCount} files. Time used: {timeSpan.ToString(@"hh\:mm\:ss")}]");
                Console.ResetColor();
                Console.WriteLine($"[ Press any key to exit the program... ]");
                Console.ReadKey();
            }
            catch(Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine();
                Console.WriteLine($"[ Fatal error has occurred while executing the program: ]");
                Console.WriteLine($"[ FileName: {ex.InnerException.Data["fileName"]} ]");
                Console.WriteLine();
                Console.WriteLine(">> " + ex.InnerException.Data["message"]);
                Console.WriteLine();
                Console.WriteLine(ex.InnerException.Data["stackTrace"]);

                Console.ResetColor();
                Console.WriteLine($"[ Press any key to exit the program... ]");
                Console.ReadKey();

            }
        }

        public static void ClearCurrentConsoleLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            for (int i = 0; i < Console.WindowWidth; i++)
                Console.Write(" ");
            Console.SetCursorPosition(0, currentLineCursor);
        }

        private static void RemoveObsoleteObjectsFromFile(string path)
        {
            using (var compoundFile = new CompoundFile(path))//, CFSUpdateMode.ReadOnly,CFSConfiguration.SectorRecycle | CFSConfiguration.EraseFreeSectors);
            {
                if (compoundFile.RootStorage == null)
                    throw new InvalidOperationException($"RootStorage null for compoundFile: {path}");

                var deletableObjects = FindObjectsToDeleteRecursive(compoundFile.RootStorage);

                RemoveStorageRecursive(compoundFile.RootStorage, deletableObjects);
                
                using (var newCompound = CopyCompoundFile(compoundFile))
                {
                    newCompound.Save(path);
                }
            }
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
                     itemName.StartsWith("STM") )
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

        private static CompoundFile CopyCompoundFile(CompoundFile compoundFile)
        {
            var newCompoundFile = new CompoundFile(compoundFile.Version, compoundFile.Configuration);

            //Copy Root CLSID
            newCompoundFile.RootStorage.CLSID = new Guid(compoundFile.RootStorage.CLSID.ToByteArray());

            CopyCFStorageRecursive(compoundFile.RootStorage, newCompoundFile.RootStorage);

            return newCompoundFile;
        }

        private static void CopyCFStorageRecursive(CFStorage sourceStorage, CFStorage destinationStorage)
        {
            Action<CFItem> copyAction = delegate (CFItem item)
            {
                if (item.IsStream)
                {
                    var itemAsStorage = item as CFStream;

                    CFStream stream = destinationStorage.AddStream(itemAsStorage.Name);
                    stream.SetData(itemAsStorage.GetData());
                }
                else if (item.IsStorage)
                {
                    var itemAsStorage = item as CFStorage;

                    CFStorage newStorage = destinationStorage.AddStorage(itemAsStorage.Name);
                    newStorage.CLSID = new Guid(itemAsStorage.CLSID.ToByteArray());
                    CopyCFStorageRecursive(itemAsStorage, newStorage); // recursion, one level deeper
                }
            };

            sourceStorage.VisitEntries(copyAction, false);
        }
    }
}
