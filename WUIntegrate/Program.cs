using DiscUtils.Udf;
using Microsoft.Dism;
using System.Collections.ObjectModel;

namespace WUIntegrate;

public class WUIntegrateRoot
{
    private static readonly bool SkipDism = false;
    private static readonly bool SkipUpdate = false;

    private static bool IsAdmin => Helper.IsCurrentUserAdmin();

    public static readonly Locations Directories = new()
    {
        WuRoot = Path.Combine(Path.GetTempPath(), "WUIntegrate"),
        ScanCabExtPath = Path.Combine(Path.GetTempPath(), "WUIntegrate", "scancab"),
        MediumPath = string.Empty,
        DismMountPath = Path.Combine(Path.GetTempPath(), "WUIntegrate", "mount"),
        MediumExtractPath = Path.Combine(Path.GetTempPath(), "WUIntegrate", "extract"),
        DlUpdatesPath = Path.Combine(Path.GetTempPath(), "WUIntegrate", "updates")
    };

    public static void Main(string[] args)
    {
        Console.Title = "WUIntegrate";
        CreatePaths();
        StartLogging();

        var inputPath = GetInputPath(args);
        if (inputPath == string.Empty)
        {
            Logger.Error(Constants.Notices.Usage);
            return;
        }

        if (!IsAdmin)
        {
            Logger.Error(Constants.Notices.Admin);
            return;
        }

        // Test Path
        TestPathValidity(inputPath);
        inputPath = Path.GetFullPath(inputPath);

        // Test Space
        TestDriveSpace(inputPath);

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

        // Identify Medium
        Logger.Msg("Identifying installation medium type");
        var mediumType = GetMediumType(inputPath);
        Logger.Msg($"Detected medium: {mediumType}");

        // Handle different medium types
        switch (mediumType)
        {
            case MediumType.IsoFile:
                Logger.Msg("Extracting ISO file");
                ExtractISO(inputPath);
                break;
            case MediumType.EsdFile:
                Helper.ExceptionFactory<ArgumentException>("ESD files are not supported.");
                break;
            case MediumType.Unknown:
                Helper.ExceptionFactory<ArgumentException>("Unknown medium type.");
                break;
        }

        if (!SkipDism && !SkipUpdate)
        {
            // Phase 3: Updating
            var windowsImages = GetDismImages(inputPath);
            UpdatePhase(ref windowsImages);
        } else
        {
            Logger.Warn("Skipping update phase because DISM is not available, or the Update system has been disabled.");
        }

        // Phase 4: Cleanup
        switch (mediumType)
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

        EndLogging();
        return;
    }

    private static string GetInputPath(string[] args)
    {
        return args.Length > 0 ? args[0] : string.Empty;
    }

    private static void StartLogging()
    {
        Logger.EnableLogging = true;
        Logger.LogFile = Path.Combine(Directories.WuRoot, "wu.log");
        var bar = new string('-', 20);
        Logger.Log(bar + " STARTING LOG " + bar);
    }

    private static void EndLogging()
    {
        var bar = new string('-', 20);
        Logger.Log(bar + " ENDING LOG " + bar);
    }

    private static IEnumerable<DismImageInfo> GetDismImages(string imagePath)
    {
        try
        {
            DismImageInfoCollection images = DismApi.GetImageInfo(imagePath);
            var bar = new string('-', 20);

            ConsoleWriter.WriteLine(bar, ConsoleColor.Blue);
            foreach (var image in images)
            {
                ConsoleWriter.WriteLine(
                    $"""
                    [{image.ImageIndex}]
                        {image.ImageName}
                        {image.Architecture}
                        {image.ProductVersion.Major}.{image.ProductVersion.Minor}.{image.ProductVersion.Build}.{image.ProductVersion.Revision}
                    """, ConsoleColor.White);
                ConsoleWriter.WriteLine(bar, ConsoleColor.Blue);
            }

            int index = 0;
            do
            {
                index = ConsoleWriter.PromptInt("Choose a WIM index, or enter 0 for all", ConsoleColor.Green);

                if (index == 0)
                {
                    break;
                }

            } while (!images.Any(x => x.ImageIndex == index));

            if (index == 0)
            {
                return images;
            }
            else
            {
                return images.Where(x => x.ImageIndex == index);
            }
        }
        catch (Exception ex)
        {
            Helper.ExceptionFactory<Exception>($"Failed to get WIM index selection: {ex}");
            return [];
        }
    }

    private static void UpdatePhase(ref IEnumerable<DismImageInfo> images)
    {
        if (!images.Any())
        {
            Helper.ExceptionFactory<NullReferenceException>("The WIM file has no images.");
            return;
        }

        Logger.Msg("Initializing update downloader");
        var updateSystem = new UpdateSystem();

        foreach (var image in images)
        {
            Logger.Msg($"Setting Windows architecture and version for this image");
            var version = GetWindowsVersion(image.ProductVersion.Build, image.ProductType);
            var architecture = GetArchitecture(image.Architecture);

            ConsoleWriter.WriteLine(
                $"""
                --------------------
                Current Image:
                    WIM Index: {image.ImageIndex}
                    Windows Image Version: {version}
                    Architecture: {architecture}
                --------------------
                """, ConsoleColor.White);

            Logger.Msg("Mounting WIM");
            MountWIM(image.ImageIndex);

            Logger.Msg("Finding latest updates");
            updateSystem.Start(version, architecture);

            if (updateSystem.ReadyToIntegrate)
            {
                Logger.Msg("Integrating updates. This may take a while.");
                IntegrateUpdateFiles();
            }

            CommitWIM();
        }
    }

    // PATH OPERATIONS
    private static void CreatePaths()
    {
        foreach (var path in Directories.All)
        {
            if (path == string.Empty)
            {
                continue;
            }

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }

    private static void DeleteTempPath()
    {
        string[] SafetyDirs =
        [
            Environment.SpecialFolder.MyDocuments.ToString(),
            Environment.SpecialFolder.MyMusic.ToString(),
            Environment.SpecialFolder.MyPictures.ToString(),
            Environment.SpecialFolder.MyVideos.ToString(),
            Environment.SpecialFolder.DesktopDirectory.ToString(),
            Environment.SpecialFolder.Windows.ToString(),
            Path.Combine(Environment.SpecialFolder.UserProfile.ToString(), "Downloads"),
        ];

        var allDirs = Directories.All
            .Where(x => x != Path.GetTempPath())
            .Where(x => !SafetyDirs.Contains(x))
            .Where(x => x != Directories.WuRoot);
;
        foreach (var dir in allDirs)
        {
            if (Directory.Exists(dir))
            {
                Logger.Log($"Cleaning up temporary directory: {dir}");
                Directory.Delete(dir, recursive: true);
            }
        }
    }

    private static void TestPathValidity(string path)
    {
        if (Directory.Exists(path))
        {
            Helper.ExceptionFactory<ArgumentException>("Path specified is a directory. Directories are not supported.");
        }
        if (!File.Exists(path))
        {
            Helper.ExceptionFactory<FileNotFoundException>("Path specified does not exist.");
        }
    }

    private static void TestDriveSpace(string path)
    {
        var DriveRoot = Path.GetPathRoot(Path.GetTempPath())!.TrimEnd('\\');

        var FileInfo = new FileInfo(path);
        var DriveInfo = new DriveInfo(DriveRoot);

        if (DriveInfo.AvailableFreeSpace < FileInfo.Length)
        {
            Helper.ExceptionFactory<IOException>("Not enough space on the drive to extract the file.");
        }
    }

    private static MediumType GetMediumType(string path)
    {
        var medium = Path.GetExtension(path).ToUpper() switch
        {
            ".WIM" => MediumType.WimFile,
            ".ISO" => MediumType.IsoFile,
            ".ESD" => MediumType.EsdFile,
            _ => MediumType.Unknown,
        };

        if (medium == MediumType.IsoFile)
        {
            Directories.MediumPath = Path.Combine(Path.Combine(Directories.MediumExtractPath, "sources"), "install.wim");
            return MediumType.IsoFile;
        }

        if (medium == MediumType.WimFile)
        {
            Directories.MediumPath = path;

            var parentPath = Path.GetDirectoryName(path);
            var di = new DirectoryInfo(parentPath!);
            if (di.Attributes.HasFlag(FileAttributes.ReadOnly))
            {
                Helper.ExceptionFactory<UnauthorizedAccessException>("Directory of the WIM is read only.");
            }
            return MediumType.WimFile;
        }

        return MediumType.Unknown; 
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
    private static void MountWIM(int index)
    {
        var existingImages = DismApi.GetMountedImages();
        foreach (var image in existingImages)
        {
            Logger.Warn($"Existing mounted image detected ({image.MountPath}). Unmounting image.");
            DismApi.UnmountImage(image.MountPath, false, progressCallback: DismCallback);
        }

        try
        {
            DismApi.MountImage(Directories.MediumPath!, Directories.DismMountPath, index, false, progressCallback: DismCallback);
            Logger.Msg($"WIM mounted successfully to [{Directories.DismMountPath}].");
        }
        catch (DismException ex)
        {
            Helper.ExceptionFactory<DismException>($"Failed to mount WIM: {ex}");
        }
    }

    private static Architecture GetArchitecture(DismProcessorArchitecture architecture)
    {
        return architecture switch
        {
            DismProcessorArchitecture.Intel => Architecture.x86,
            DismProcessorArchitecture.AMD64 => Architecture.x64,
            DismProcessorArchitecture.ARM64 => Architecture.ARM64,
            _ => Architecture.Unknown,
        };
    }

    private static WindowsVersion GetWindowsVersion(int build, string productType)
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
            return build switch
            {
                6001 => WindowsVersion.WindowsServer2008,
                7601 => WindowsVersion.WindowsServer2008R2,
                9200 => WindowsVersion.WindowsServer2012,
                9600 => WindowsVersion.WindowsServer2012R2,
                14393 => WindowsVersion.WindowsServer2016,
                17763 => WindowsVersion.WindowsServer2019,
                20348 => WindowsVersion.WindowsServer2022,
                26100 => WindowsVersion.WindowsServer2025,
                _ => WindowsVersion.Unknown,
            };
        }

        if (productType == "WinNT")
        {
            return build switch
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

        return WindowsVersion.Unknown;
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
            if (DismApi.GetMountedImages()
                .Select(x => x.MountPath)
                .Contains(Directories.DismMountPath))
            {
                Logger.Warn("Unmounting WIM image.");
                DismApi.UnmountImage(Directories.DismMountPath, true, progressCallback: DismCallback);
            }
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
}
