using DiscUtils.Udf;
using Microsoft.CST.RecursiveExtractor;
using Microsoft.Dism;

namespace WUIntegrate;

public class WUIntegrateRoot
{
    // Flags
    static readonly bool SkipDism = false;
    static readonly bool SkipUpdate = false;

    private static int WimIndex; // Selected WIM index
    private static MediumType Medium; // Type of medium we are working with
    private static WindowsVersion WindowsVersion; // Medium OS version
    private static Architecture SystemArchitecture; // Medium architecture
    private static string? ArgumentPath = null;
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

        if (!Helper.IsCurrentUserAdmin())
            ConsoleWriter.WriteLine(Constants.Notices.Admin, ConsoleColor.Yellow);

        ArgumentPath = args.FirstOrDefault();
        ArgumentPath = @"C:\Users\Ryze\Downloads\en-us_windows_10_enterprise_ltsc_2021_x64_dvd_d289cf96\sources\install.wim";

        ConsoleWriter.WriteLine(Constants.Notices.Startup, ConsoleColor.Yellow);

        // Test Arguments
        if (ArgumentPath == null)
        {
            ConsoleWriter.WriteLine(Constants.Notices.Usage, ConsoleColor.Yellow);
            return;
        }

        // Phase 1: Initialization
        Initialization();

        // Phase 2: DISM Choice
        if (SkipDism)
        {
            ConsoleWriter.WriteLine("[i] DISM is disabled. Continuing is impossible.", ConsoleColor.Red);
            CleanupPhase();
            return;
        }
        DismChoicePhase();

        // Phase 3: Update Phase
        UpdatePhase();

        // Phase 4: Cleanup
        CleanupPhase();

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
            ConsoleWriter.WriteLine("[i] Initializing DISM", ConsoleColor.Cyan);
            DismApi.Initialize(DismLogLevel.LogErrorsWarningsInfo);
        }
        else
        {
            ConsoleWriter.WriteLine("[i] DISM disabled!", ConsoleColor.Red);
        }

        // Create Paths
        ConsoleWriter.WriteLine("[i] Creating temporary directories", ConsoleColor.Cyan);
        CreatePaths();

        // Identify Medium
        ConsoleWriter.WriteLine("[i] Identifying installation medium type", ConsoleColor.Cyan);
        IdentifyMedium(ArgumentPath!);
        ConsoleWriter.WriteLine($"[i] Detected medium: {Medium}", ConsoleColor.Cyan);


        // Handle different medium types
        switch (Medium)
        {
            case MediumType.IsoFile:
                ConsoleWriter.WriteLine("[i] Extracting ISO file", ConsoleColor.Cyan);
                Task.Run(async () =>
                {
                    await ExtractISO(ArgumentPath!);
                }).Wait();
                break;
            case MediumType.WimFile:
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
        ConsoleWriter.WriteLine("[i] Loading WIM", ConsoleColor.Cyan);
        SetWinVersionAndArchitecture(Directories.MediumPath!);
    }

    private static void UpdatePhase()
    {
        // Start Update System
        if (!SkipUpdate)
        {
            ConsoleWriter.WriteLine("[i] Initializing update downloader", ConsoleColor.Cyan);
            updateSystem = new();
        }

        do
        {
            // Current Image
            var CurrentImage = IntegratableImages!.First();

            // Set Windows Version and Architecture
            ConsoleWriter.WriteLine("[i] Setting Windows architecture and version", ConsoleColor.Cyan);
            SetWindowsVersion(CurrentImage.ProductVersion.Build, CurrentImage.ProductType);
            SetArchitecture(CurrentImage.Architecture);

            // Confirm information
            ConsoleWriter.WriteLine(
                $"""
            --------------------
            Current Image:
                WIM Index: {CurrentImage.ImageIndex}
                Windows Image Version: {WindowsVersion}
                Architecture: {SystemArchitecture}
            --------------------
            """
            , ConsoleColor.White);

            // Mount WIM
            if (!SkipDism)
            {
                ConsoleWriter.WriteLine("[i] Mounting WIM", ConsoleColor.Cyan);
                MountWIM();
            }

            // Start Update System
            if (!SkipUpdate)
            {
                // Get updates for this version
                ConsoleWriter.WriteLine("[i] Finding latest updates", ConsoleColor.Cyan);
                updateSystem!.Start(WindowsVersion, SystemArchitecture);

                // Integrate updates if specified
                if (updateSystem.ReadyToIntegrate)
                {
                    Console.Clear();
                    ConsoleWriter.WriteLine("[i] Integrating updates. This may take a while.", ConsoleColor.Cyan);
                    IntegrateUpdateFiles();
                }
            }

            if (!SkipDism)
            {
                // Unmount WIM
                ConsoleWriter.WriteLine("[i] Applying WIM changes", ConsoleColor.Cyan);
                CommitWIM();
            }

            // Remove from the list the image we just handled
            IntegratableImages!.Remove(CurrentImage);
        } while (IntegratableImages!.Count > 0);
    }

    private static void CleanupPhase()
    {
        // Cleanup
        ConsoleWriter.WriteLine("[i] Cleaning up", ConsoleColor.Cyan);
        Cleanup();
    }

    // PATH OPERATIONS
    private static void CreatePaths()
    {
        foreach (var path in Directories.All)
        {
            if (path == Path.GetTempPath()) continue;
            if (path == null) continue;

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }

    private static void DeleteTempPath()
    {
        if (Directories.WuRoot != null)
        {
            if (!Directory.Exists(Directories.WuRoot))
                return;
            Directory.Delete(Directories.WuRoot, recursive: true);
        }
    }

    private static void TestPath(string path)
    {
        if (Directory.Exists(path))
            Helper.ExceptionFactory<ArgumentException>(
                "Path specified is a directory. Directories are not supported."
            );
        if (!File.Exists(path))
            Helper.ExceptionFactory<FileNotFoundException>("Path specified does not exist.");
    }

    private static void TestSpace(string path)
    {
        var DriveRoot = Path.GetPathRoot(Path.GetTempPath())!.TrimEnd('\\');

        var FileInfo = new FileInfo(path);
        var DriveInfo = new DriveInfo(DriveRoot);

        var DriveFreeSpace = DriveInfo.AvailableFreeSpace;
        var FileSize = FileInfo.Length;

        if (DriveFreeSpace < FileSize)
            Helper.ExceptionFactory<IOException>(
                "Not enough space on the drive to extract the file."
            );
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

        if (Medium == MediumType.WimFile || Medium == MediumType.EsdFile)
        {
            Directories.MediumPath = path;

            var parentPath = Path.GetDirectoryName(path);
            var di = new DirectoryInfo(parentPath!);
            if (di.Attributes.HasFlag(FileAttributes.ReadOnly))
            {
                Helper.ExceptionFactory<UnauthorizedAccessException>(
                    "Directory of the WIM is read only. This is not supported. Please move the WIM file to a writable directory or drive."
                );
            }
            return;
        }

        Helper.ExceptionFactory<ArgumentException>("The given path is not a valid WIM or ISO file.");
    }

    private static async Task ExtractISO(string path)
    {
        using var udfStream = new FileStream(path, FileMode.Open, FileAccess.Read);
        using var reader = new UdfReader(udfStream);
        var extractor = new Extractor();
        var root = reader.Root;
        var allFiles = root.GetFiles("*", searchOption: SearchOption.AllDirectories);

        // Create destination folder structure
        byte[] buffer = new byte[81920];
        long totalBytes = allFiles.Sum(x => x.Length);
        long totalBytesRead = 0;
        int bytesRead;

        foreach (var file in allFiles)
        {
            using var sourceFileStream = file.OpenRead();

            // Create directory if it doesn't exist
            var destinationPath = Path.Combine(Directories.MediumExtractPath, file.FullName);
            Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);

            using var destinationFileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write);

            while ((bytesRead = sourceFileStream.Read(buffer)) > 0)
            {
                await destinationFileStream.WriteAsync(buffer.AsMemory(0, bytesRead));
                totalBytesRead += bytesRead;
                int progress = (int)((double)totalBytesRead / totalBytes * 100);
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
            string bar = new('-', 20);

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

                if (WimIndex == 0) break;

            } while (!ImageInfo.Any(x => x.ImageIndex == WimIndex));

            if (WimIndex == 0)
            {
                foreach (var image in ImageInfo)
                {
                    IntegratableImages.Add(image);
                }
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
            ConsoleWriter.WriteLine($"[!] Existing mounted image detected: {image.MountPath}", ConsoleColor.Red);
            DismApi.UnmountImage(image.MountPath, false, progressCallback: DismCallback);
        }

        try
        {
            DismApi.MountImage(Directories.MediumPath!, Directories.DismMountPath, WimIndex, false, progressCallback: DismCallback);
            ConsoleWriter.WriteLine($"[i] WIM is now mounted to: {Directories.DismMountPath}", ConsoleColor.Yellow);
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
            Console.WriteLine("""
                WARNING:
                    1904X and 226X1 builds are forced to their latest version.
                    This works, but they will not update if they are end of life.
                """);
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
        int count = updateFiles.Length;
        int current = 1;
        foreach (var updateFile in updateFiles)
        {
            ConsoleWriter.WriteLine($"[{current}/{count}] Integrating Update...", ConsoleColor.Cyan);
            try
            {
                DismApi.AddPackage(session, updateFile, false, false, progressCallback: DismCallback);
                Helper.DeleteFile(updateFile);
            }
            catch (DismException ex)
            {
                ConsoleWriter.WriteLine($"[!] Failed to integrate update: {ex}", ConsoleColor.Magenta);
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
            // Ignore
        }

        // Delete Temp Path
        DeleteTempPath();
    }

    //


    // THINGS TO DO LATER
    private static void CreateBootableImage() { }
}
