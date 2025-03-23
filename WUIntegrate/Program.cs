using DiscUtils.Udf;
using Microsoft.Dism;

namespace WUIntegrate;

public class WUIntegrateRoot
{
    //yes 
    private static readonly bool SkipDism = false;
    private static readonly bool SkipUpdate = false;

    private static int WimIndex;
    private static MediumType Medium;
    private static WindowsVersion WindowsVersion;
    private static Architecture SystemArchitecture;
    private static string? ArgumentPath;
    private static List<DismImageInfo>? IntegratableImages;
    private static UpdateSystem? updateSystem;

    public static readonly Locations Directories = new()
    {
        SysTemp = Path.GetTempPath(),
        WuRoot = Path.Combine(Path.GetTempPath(), "WUIntegrate"),
        ScanCabExtPath = Path.Combine(Path.GetTempPath(), "WUIntegrate", "scancab"),
        MediumPath = null,
        DismMountPath = Path.Combine(Path.GetTempPath(), "WUIntegrate", "mount"),
        MediumExtractPath = Path.Combine(Path.GetTempPath(), "WUIntegrate", "extract"),
        DlUpdatesPath = Path.Combine(Path.GetTempPath(), "WUIntegrate", "updates")
    };

    public static void Main(string[] args)
    {
        Console.Title = "WUIntegrate";

        var imagePath = args.GetValue(0);
        var logBool = args.GetValue(1);
        if (imagePath is null)
        {
            Console.WriteLine("Usage: WUIntegrate.exe (wim / iso path) OPTIONAL: [log boolean]");
            return;
        }
        ArgumentPath = imagePath.ToString();

        if (logBool is null) // if not specified
        {
            Logger.EnableLogging = true;
            Logger.LogFile = Path.Combine(Directories.SysTemp, "wuintegrate.log");
            Logger.Msg("Logging enabled.");
        }
        else // if specified
        {
            if (bool.TryParse(logBool.ToString(), out var result))
            {
                Logger.EnableLogging = result;
                if (result)
                {
                    Logger.LogFile = Path.Combine(Directories.SysTemp, "wuintegrate.log");
                    Logger.Msg("Logging enabled.");
                }
            }
        }

        var bar = new string('-', 5);
        Logger.Log(bar + " STARTING LOG " + bar);
        Logger.Msg(Constants.Notices.Startup);

        if (!Helper.IsCurrentUserAdmin())
        {
            Logger.Error(Constants.Notices.Admin);
            return;
        }

        // Phase 1: Initialization
        Initialization();

        // Phase 2: DISM Choice
        if (SkipDism)
        {
            Logger.Error("DISM is disabled. Continuing is impossible.");
            CleanupPhase();
            return;
        }
        DismChoicePhase();

        // Phase 3: Update Phase
        UpdatePhase();

        // Phase 4: Cleanup
        CleanupPhase();

        Logger.Log(bar + " ENDING LOG " + bar);
        return;
    }

    private static void Initialization()
    {
        // Test Path
        TestPath(ArgumentPath!);

        // Test Space
        TestSpace(ArgumentPath!);

        // Initializing DISM API
        if (!SkipDism)
        {
            Logger.Msg("Initializing DISM");
            DismApi.Initialize(DismLogLevel.LogErrorsWarningsInfo);
        }
        else
        {
            Logger.Warn("DISM is disabled.");
        }

        // Create Paths
        Logger.Msg("Creating temporary directories");
        CreatePaths();

        // Identify Medium
        Logger.Msg("Identifying installation medium type");
        IdentifyMedium(ArgumentPath!);
        Logger.Msg($"Detected medium: {Medium}");


        // Handle different medium types
        switch (Medium)
        {
            case MediumType.IsoFile:
                Logger.Msg("Extracting ISO file");
                ExtractISO(ArgumentPath!);
                break;
            case MediumType.EsdFile:
                Helper.ExceptionFactory<ArgumentException>("ESD files are not supported.");
                break;
            case MediumType.Unknown:
                Helper.ExceptionFactory<ArgumentException>("Unknown medium type.");
                break;
        }
    }

    private static void DismChoicePhase()
    {
        Logger.Msg("Loading WIM");
        SetWinVersionAndArchitecture(Directories.MediumPath!);
    }

    private static void UpdatePhase()
    {
        if (SkipUpdate)
        {
            Logger.Warn("Update system is disabled. Skipping update phase.");
            return;
        }

        if (IntegratableImages is null)
        {
            Helper.ExceptionFactory<NullReferenceException>("IntegratableImages is null.");
            return;
        }

        Logger.Msg("Initializing update downloader");
        updateSystem = new();

        do
        {
            // Current Image
            var currentImage = IntegratableImages.First();

            // Set Windows Version and Architecture
            Logger.Msg("Setting Windows architecture and version");
            SetWindowsVersion(currentImage.ProductVersion.Build, currentImage.ProductType);
            SetArchitecture(currentImage.Architecture);

            // Confirm information
            ConsoleWriter.WriteLine(
                $"""
                --------------------
                Current Image:
                    WIM Index: {currentImage.ImageIndex}
                    Windows Image Version: {WindowsVersion}
                    Architecture: {SystemArchitecture}
                --------------------
                """
            , ConsoleColor.White);

            // Mount WIM
            if (!SkipDism)
            {
                Logger.Msg("Mounting WIM");
                MountWIM();
            }

            // Start Update System
            if (!SkipUpdate)
            {
                // Get updates for this version
                Logger.Msg("Finding latest updates");
                updateSystem.Start(WindowsVersion, SystemArchitecture);

                // Integrate updates if specified
                if (updateSystem.ReadyToIntegrate)
                {
                    Console.Clear();
                    Logger.Msg("Integrating updates. This may take a while.");
                    IntegrateUpdateFiles();
                }
            }

            if (!SkipDism)
            {
                // Unmount WIM
                Logger.Msg("Applying WIM changes");
                CommitWIM();
            }

            // Remove from the list the image we just handled
            IntegratableImages.Remove(currentImage);
        } while (IntegratableImages.Count > 0);
    }

    private static void CleanupPhase()
    {
        // Cleanup
        switch (Medium)
        {
            case MediumType.IsoFile:
                Logger.Msg(
                    $"""
                    Operations have completed. ISO rebuilding will be added in a future release.

                    You can find the extracted files in:
                        {Directories.MediumExtractPath}
                    """);
                break;
            default:
                Logger.Msg("Operations have completed. Your WIM file has been updated.");
                break;
        }

        Cleanup();
    }

    // PATH OPERATIONS
    private static void CreatePaths()
    {
        foreach (var path in Directories.All)
        {
            if (path == Path.GetTempPath()) continue;
            if (path is null) continue;

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }

    private static void DeleteTempPath()
    {
        string[] SafetyDirs =
        {
            Environment.SpecialFolder.MyDocuments.ToString(),
            Environment.SpecialFolder.MyMusic.ToString(),
            Environment.SpecialFolder.MyPictures.ToString(),
            Environment.SpecialFolder.MyVideos.ToString(),
            Environment.SpecialFolder.DesktopDirectory.ToString(),
            Environment.SpecialFolder.Windows.ToString(),
            Path.Combine(Environment.SpecialFolder.UserProfile.ToString(), "Downloads"),
        };

        var allDirs = Directories.All.Where(x => x != Path.GetTempPath());
        
        // emergency check to not delete system directories
        allDirs = allDirs.Where(x => !SafetyDirs.Contains(x));

        foreach (var dir in allDirs)
        {
            if (Directory.Exists(dir))
            {
                Logger.Log($"Cleaning up temporary directory: {dir}");
                Directory.Delete(dir, recursive: true);
            }
        }
    }

    private static void TestPath(string path)
    {
        if (Directory.Exists(path))
        {
            Helper.ExceptionFactory<ArgumentException>("Path specified is a directory. Directories are not supported.");
        }
        if (!File.Exists(path))
        {
            Helper.ExceptionFactory<FileNotFoundException>("Path specified does not exist.");
        }

        // Check if this is a relative path, and if it is, convert to fullpath
        if (!Path.IsPathRooted(path))
        {
            ArgumentPath = Path.GetFullPath(path);
        }
    }

    private static void TestSpace(string path)
    {
        var DriveRoot = Path.GetPathRoot(Path.GetTempPath())!.TrimEnd('\\');

        var FileInfo = new FileInfo(path);
        var DriveInfo = new DriveInfo(DriveRoot);

        if (DriveInfo.AvailableFreeSpace < FileInfo.Length)
        {
            Helper.ExceptionFactory<IOException>("Not enough space on the drive to extract the file.");
        }
    }

    private static void IdentifyMedium(string path)
    {
        Medium = Path.GetExtension(path).ToUpper() switch
        {
            ".WIM" => MediumType.WimFile,
            ".ISO" => MediumType.IsoFile,
            ".ESD" => MediumType.EsdFile,
            _ => MediumType.Unknown,
        };

        if (Medium == MediumType.IsoFile)
        {
            Directories.MediumPath = Path.Combine(Path.Combine(Directories.MediumExtractPath, "sources"), "install.wim");
            return;
        }

        if (Medium == MediumType.WimFile)
        {
            Directories.MediumPath = path;

            var parentPath = Path.GetDirectoryName(path);
            var di = new DirectoryInfo(parentPath!);
            if (di.Attributes.HasFlag(FileAttributes.ReadOnly))
            {
                Helper.ExceptionFactory<UnauthorizedAccessException>("Directory of the WIM is read only.");
            }
            return;
        }

        Helper.ExceptionFactory<ArgumentException>("The given path is not a valid WIM or ISO file.");
    }

    private static void ExtractISO(string path)
    {
        using var udfStream = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var reader = new UdfReader(udfStream);
        var root = reader.Root;
        var allFiles = root.GetFiles("*", searchOption: SearchOption.AllDirectories);

        // Create destination folder structure
        var buffer = new byte[81920];
        var totalBytes = allFiles.Sum(x => x.Length);
        var totalBytesRead = 0;
        var bytesRead = 0;

        foreach (var file in allFiles)
        {
            using var sourceFileStream = file.OpenRead();

            // Create directory if it doesn't exist
            var destinationPath = Path.Combine(Directories.MediumExtractPath, file.FullName);
            Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);

            using var destinationFileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write);

            while ((bytesRead = sourceFileStream.Read(buffer)) > 0)
            {
                destinationFileStream.Write(buffer, 0, bytesRead);
                totalBytesRead += bytesRead;
                var progress = (int)((double)totalBytesRead / totalBytes * 100);
                ConsoleWriter.WriteProgress(100, progress, 100, "Extracting", ConsoleColor.White);
            }
        }
    }

    //

    // DISM OPERATIONS
    private static void SetWinVersionAndArchitecture(string path)
    {
        try
        {
            DismImageInfoCollection ImageInfo = DismApi.GetImageInfo(path);
            IntegratableImages = [];
            var bar = new string('-', 20);

            ConsoleWriter.WriteLine(bar, ConsoleColor.Blue);
            foreach (var index in ImageInfo)
            {
                ConsoleWriter.WriteLine(
                    $"""
                    [{index.ImageIndex}]
                        {index.ImageName}
                        {index.Architecture}
                        {index.ProductVersion.Major}.{index.ProductVersion.Minor}.{index.ProductVersion.Build}.{index.ProductVersion.Revision}
                    """
                , ConsoleColor.White);
                ConsoleWriter.WriteLine(bar, ConsoleColor.Blue);
            }

            do
            {
                WimIndex = ConsoleWriter.PromptInt(
                    "Choose a WIM index. Enter 0 for all. "
                , ConsoleColor.Green);

                if (WimIndex == 0)
                {
                    break;
                }

            } while (!ImageInfo.Any(x => x.ImageIndex == WimIndex));

            if (WimIndex == 0)
            {
                IntegratableImages.AddRange(ImageInfo);
                return;
            }
            else
            {
                var SelectedImage = ImageInfo.First(x => x.ImageIndex == WimIndex);
                IntegratableImages.Add(SelectedImage);
            }
        }
        catch (Exception ex)
        {
            Helper.ExceptionFactory<Exception>($"Failed to get WIM index selection: {ex}");
        }
    }

    private static void MountWIM()
    {
        var existingImages = DismApi.GetMountedImages();
        foreach (var image in existingImages)
        {
            Logger.Warn($"Existing mounted image detected ({image.MountPath}). Unmounting image.");
            DismApi.UnmountImage(image.MountPath, false, progressCallback: DismCallback);
        }

        try
        {
            DismApi.MountImage(Directories.MediumPath!, Directories.DismMountPath, WimIndex, false, progressCallback: DismCallback);
            Logger.Msg($"WIM mounted successfully to [{Directories.DismMountPath}].");
        }
        catch (DismException ex)
        {
            Helper.ExceptionFactory<DismException>($"Failed to mount WIM: {ex}");
        }
    }

    private static void SetArchitecture(DismProcessorArchitecture architecture)
    {
        SystemArchitecture = architecture switch
        {
            DismProcessorArchitecture.Intel => Architecture.x86,
            DismProcessorArchitecture.AMD64 => Architecture.x64,
            DismProcessorArchitecture.ARM64 => Architecture.ARM64,
            _ => Architecture.Unknown,
        };
        if (SystemArchitecture == Architecture.Unknown)
        {
            Helper.ExceptionFactory<ArgumentException>($"{architecture} is not supported.");
        }
    }

    private static void SetWindowsVersion(int build, string productType)
    {
        if ((build >= 19041 && build <= 19045) || (build >= 22621 && build <= 22631))
        {
            Logger.Log("Enablement package build detected. Forcing Windows 10 22H2.");
            ConsoleWriter.WriteLine("""
                WARNING:
                    1904X and 226X1 builds are forced to their latest version.
                    This works, but they will not update if they are end of life.
                """, ConsoleColor.Yellow);
        }

        // Server ProductType is ServerNT, EditionID starts with Server
        // Client ProductType is WinNT, EditionID is random
        if (productType == "ServerNT")
        {
            switch (build)
            {
                case 6001:
                    WindowsVersion = WindowsVersion.WindowsServer2008;
                    break;
                case 7601:
                    WindowsVersion = WindowsVersion.WindowsServer2008R2;
                    break;
                case 9200:
                    WindowsVersion = WindowsVersion.WindowsServer2012;
                    break;
                case 9600:
                    WindowsVersion = WindowsVersion.WindowsServer2012R2;
                    break;
                case 14393:
                    WindowsVersion = WindowsVersion.WindowsServer2016;
                    break;
                case 17763:
                    WindowsVersion = WindowsVersion.WindowsServer2019;
                    break;
                case 20348:
                    WindowsVersion = WindowsVersion.WindowsServer2022;
                    break;
                case 26100:
                    WindowsVersion = WindowsVersion.WindowsServer2025;
                    break;
            }
        }

        if (productType == "WinNT")
        {
            WindowsVersion = build switch
            {
                7601 => WindowsVersion.Windows7,
                9600 => WindowsVersion.Windows81,
                10240 => WindowsVersion.Windows10RTM,
                10586 => WindowsVersion.Windows10TH2,
                14393 => WindowsVersion.Windows10RS1,
                15063 => WindowsVersion.Windows10RS2,
                16299 => WindowsVersion.Windows10RS3,
                17134 => WindowsVersion.Windows10RS4,
                17763 => WindowsVersion.Windows10RS5,
                18362 => WindowsVersion.Windows1019H1,
                18363 => WindowsVersion.Windows1019H2,
                19041 => WindowsVersion.Windows1022H2,
                19042 => WindowsVersion.Windows1022H2,
                19043 => WindowsVersion.Windows1022H2,
                19044 => WindowsVersion.Windows1022H2,
                19045 => WindowsVersion.Windows1022H2,
                22000 => WindowsVersion.Windows1121H2,
                22621 => WindowsVersion.Windows1123H2,
                22631 => WindowsVersion.Windows1123H2,
                26100 => WindowsVersion.Windows1124H2,
                _ => WindowsVersion.Unknown,
            };
        }

        if (WindowsVersion == WindowsVersion.Unknown)
        {
            Helper.ExceptionFactory<ArgumentException>("Windows version is unsupported.");
        }
    }

    private static void IntegrateUpdateFiles()
    {
        // Use DISM update API to integrate all files within the update folder
        var session = DismApi.OpenOfflineSession(Directories.DismMountPath);

        var updateFiles = Directory.GetFiles(Directories.DlUpdatesPath);
        var count = updateFiles.Length;
        var current = 1;
        foreach (var updateFile in updateFiles)
        {
            Logger.Msg($"[{current}/{count}] Integrating Update...");
            try
            {
                DismApi.AddPackage(session, updateFile, false, false, progressCallback: DismCallback);
                Helper.DeleteFile(updateFile);
            }
            catch (DismException ex)
            {
                Logger.Error($"Failed to integrate update: {ex.Message}");
            }
            current++;
        }

        DismApi.CloseSession(session);
    }

    private static void DismCallback(DismProgress dismProgress)
    {
        ConsoleWriter.WriteProgress(100, dismProgress.Current, dismProgress.Total, "DISM", ConsoleColor.Yellow);
    }

    private static void CommitWIM()
    {
        try
        {
            DismApi.UnmountImage(Directories.DismMountPath, true, progressCallback: DismCallback);
        }
        catch (DismException ex)
        {
            Helper.ExceptionFactory<DismException>($"Failed to commit WIM: {ex}");
        }
    }

    //

    // CLEANUP OPERATIONS
    internal static void Cleanup()
    {
        try
        {
            DismApi.CleanupMountpoints();
            DismApi.Shutdown();
        }
        catch
        {
            Logger.Error("Unable to close DISM mountpoints.");
        }

        // Delete Temp Path
        DeleteTempPath();
    }

    //


    // THINGS TO DO LATER
    private static void CreateBootableImage(string sourceFiles, string destinationPath)
    {
        if (!Directory.Exists(sourceFiles))
        {
            Helper.ExceptionFactory<DirectoryNotFoundException>("Source files directory does not exist.");
        }

        if (File.Exists(destinationPath))
        {
            Helper.ExceptionFactory<IOException>("Destination path is a file.");
        }

        if (Directory.Exists(destinationPath))
        {
            Helper.ExceptionFactory<IOException>("Destination path is a directory.");
        }


    }
}
