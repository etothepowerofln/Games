using System;
using System.IO;
using IWshRuntimeLibrary;
using File = System.IO.File;  // Ensure you add a COM reference to "Windows Script Host Object Model"

namespace GOG_Shorts
{
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Check that a file was drag‑and‑dropped.
            if (args.Length == 0)
            {
                Console.WriteLine("Please drag and drop a file named \"goggame-gameID.ico\" onto this application.");
                Console.WriteLine("You can locate it in the folder where the game was installed.");
                Console.WriteLine("If there is no file with this name, you should create one.");
                Console.WriteLine("Also remember to create a file MyDocuments>GOG_Shorts>config.txt with path to GOG Galaxy.");
                Console.WriteLine("For a common path, paste \"C:\\Program Files(x86)\\GOG Galaxy\\GalaxyClient.exe\".");
                PauseAndExit();
                return;
            }

            string droppedFile = args[0];
            if (!File.Exists(droppedFile))
            {
                Console.WriteLine("File not found: " + droppedFile);
                PauseAndExit();
                return;
            }
            
            Console.WriteLine("GOG_Shorts by Luiz Filipi Anderson de Sousa Moura");
            Console.WriteLine(droppedFile);
            Console.WriteLine("-------------------------------------------------");

            if (droppedFile.EndsWith(".ico", StringComparison.OrdinalIgnoreCase) == false)
            {
                Console.WriteLine("ERROR: The dropped file is not an icon file (.ico).");
                PauseAndExit();
                return;
            }

            // 2. Extract the game ID (N) from the file name.
            // Get the file name (without extension); should be "goggame-N" where N is a number.
            string fileName = Path.GetFileNameWithoutExtension(droppedFile);
            const string prefix = "goggame-";
            if (!fileName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("ERROR: The file name does not match the expected pattern \"goggame-gameID.ico\".");
                PauseAndExit();
                return;
            }
            string gameId = fileName.Substring(prefix.Length);
            if (string.IsNullOrEmpty(gameId))
            {
                Console.WriteLine("ERROR: Could not extract game id from the file name.");
                PauseAndExit();
                return;
            }

            // 3. Determine the parent folder of the dropped file.
            // Use the name of that folder as the shortcut's name.
            string parentDirectory = Path.GetDirectoryName(droppedFile);
            if (string.IsNullOrEmpty(parentDirectory))
            {
                Console.WriteLine("ERROR: Could not determine the parent directory of the dropped file.");
                PauseAndExit();
                return;
            }
            string shortcutName = new DirectoryInfo(parentDirectory).Name;

            // 4. Read the external value V from configuration.
            // The config.txt is now expected to be at "C:\ProgramData\GOG_Shorts\config.txt".
            string configFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "GOG_Shorts");
            string configFile = Path.Combine(configFolder, "config.txt");
            if (!File.Exists(configFile))
            {
                Console.WriteLine("ERROR: config.txt file not found at: " + configFile);
                try
                {
                    // Create the directory if it doesn't exist.
                    if (!Directory.Exists(configFolder))
                    {
                        Directory.CreateDirectory(configFolder);
                        Console.WriteLine("Folder created: " + configFolder);
                        Console.WriteLine("Now create a file named config.txt in this folder whose text is the path to GOG Galaxy executable.");
                    }
                    Console.WriteLine("If GOG Galaxy is installed in another location, edit config.txt.");
                    Console.WriteLine("Text should be between quotation marks: \"Full\\Path\\to\\Galaxy.exe\".");
                    Console.WriteLine("For a common path, paste \"C:\\Program Files(x86)\\GOG Galaxy\\GalaxyClient.exe\".");
                    PauseAndExit();
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR creating the config.txt: " + ex.Message);
                    PauseAndExit();
                    return;
                }
            }

            string V = File.ReadAllText(configFile).Trim();
            if (string.IsNullOrEmpty(V))
            {
                Console.WriteLine("ERROR: The value from config.txt is empty.");
                PauseAndExit();
                return;
            }

            // 5. Build the shortcut target command.
            // The target string should be: V /command=runGame /gameId=N
            string arguments = $"/command=runGame /gameId={gameId}";

            // 6. Create the shortcut in the same folder as the application.
            // The shortcut's file name is based on the parent folder name of the dropped file.
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string shortcutPath = Path.Combine(appDir, shortcutName + ".lnk");
            try
            {
                WshShell shell = new();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

                // Set the target executable (V) and pass the constructed arguments.
                shortcut.TargetPath = V;
                shortcut.Arguments = arguments;
                shortcut.WorkingDirectory = appDir;

                // 7. Set the shortcut's icon to the dropped file (goggame-N.ico).
                shortcut.IconLocation = $"{Path.GetFullPath(droppedFile)},0";

                shortcut.Save();

                Console.WriteLine("Shortcut created successfully at:");
                Console.WriteLine(shortcutPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR creating the shortcut: " + ex.Message);
            }

            PauseAndExit();
        }

        static void PauseAndExit()
        {
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
