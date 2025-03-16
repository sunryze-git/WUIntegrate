using System.Diagnostics;
using System.Net;
using System.Reflection;
using System.Security.Principal;
using Microsoft.Dism;

namespace WUIntegrate;

public class Helper
{
    public static int PromptForInt(string promptMessage)
    {
        // Optimized to minimize overhead in printing and reading input
        Console.Write(promptMessage + ": ");
        string? input = Console.ReadLine();
        return int.Parse(input ?? "0");
    }

    public static bool PromptYesNo(string PromptMessage)
    {
        Console.Write(PromptMessage + " (Y/N) : ");
        char response = Char.ToUpper(Console.ReadKey().KeyChar);
        return response == 'Y';
    }

    public static void DeleteFile(string fileName)
    {
        File.Delete(fileName);
    }

    public static void DeleteFolder(string folderName)
    {
        if (!Directory.Exists(folderName))
            return;
        Directory.Delete(folderName, recursive: true);
    }

    public static void ExtractResource(string ResourceName, string DestinationPath)
    {
        Console.WriteLine($"Extracting resource {ResourceName} to {DestinationPath}");

        using var resource = Assembly
            .GetExecutingAssembly()
            .GetManifestResourceStream(ResourceName);
        using var file = new FileStream(DestinationPath, FileMode.Create, FileAccess.Write);
        if (resource == null)
            return;
        resource.CopyTo(file);
    }

    public static void ExtractFile(
        string SourcePath,
        string DestinationPath,
        string? SpecificArguments = null
    )
    {
        Console.WriteLine($"Extracting file {SourcePath} to {DestinationPath}");

        // Use the 7zr.exe to extract the file
        string arguments = $"x \"{SourcePath}\" -o\"{DestinationPath}\" {SpecificArguments} -y";

        if (!Path.Exists(WUIntegrateRoot.SevenZipExe))
        {
            ExtractResource(
                "WUIntegrate.Utils.7za.dll",
                Path.Combine(WUIntegrateRoot.UtilsPath!, "7za.dll")
            );
            ExtractResource(
                "WUIntegrate.Utils.7za.exe",
                Path.Combine(WUIntegrateRoot.UtilsPath!, "7za.exe")
            );
            ExtractResource(
                "WUIntegrate.Utils.7zxa.dll",
                Path.Combine(WUIntegrateRoot.UtilsPath!, "7zxa.dll")
            );
        }

        ProcessStartInfo startInfo = new()
        {
            FileName = WUIntegrateRoot.SevenZipExe,
            Arguments = arguments,
            CreateNoWindow = false,
            UseShellExecute = false,
        };

        try
        {
            using (Process? process = Process.Start(startInfo))
            {
                if (process == null)
                    Helper.ExceptionFactory<Exception>("Failed to start 7zr.exe process.");
                process!.WaitForExit();
            }

            Console.WriteLine($"Extraction has completed for {SourcePath} to {DestinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Extraction failed: {ex}");
        }
    }

    public static void DownloadFileUri(string Uri, string DestinationPath)
    {
        if (Uri == null)
            ExceptionFactory<ArgumentNullException>("URL to be downloaded was null.");
        if (Path.Exists(DestinationPath))
            return;

        Task.Run(async () =>
            {
                using HttpClient client = new();
                using HttpResponseMessage response = await client.GetAsync(Uri);

                response.EnsureSuccessStatusCode();

                using Stream contentStream = await response.Content.ReadAsStreamAsync(),
                    stream = new FileStream(
                        DestinationPath,
                        FileMode.Create,
                        FileAccess.Write,
                        FileShare.None
                    );
                await contentStream.CopyToAsync(stream);
            })
            .Wait();
    }

    public static bool IsCurrentUserAdmin()
    {
        using WindowsIdentity identity = WindowsIdentity.GetCurrent();
        WindowsPrincipal principal = new(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    public static void Error(string errorMessage)
    {
        UpdateSystem.Cleanup();
        WUIntegrateRoot.Cleanup();
        throw new Exception(errorMessage);
    }

    public static void ExceptionFactory<T>(string errorMessage)
        where T : Exception
    {
        UpdateSystem.Cleanup();
        WUIntegrateRoot.Cleanup();
        var ex =
            Activator.CreateInstance(typeof(T), errorMessage) as T
            ?? throw new InvalidOperationException("Failed to create exception instance.");
        throw ex;
    }
}

public class WUIntegrateRoot
{
    enum MediumType
    {
        WimFile,
        EsdFile,
        IsoFile,
        Unknown,
    }

    const string StartupNotice = "WUIntegrate - Made by sunryze";
    const string UsageNotice = """
        Usage:
        WUIntegrate.exe [MediumPath]
        """;

    const string AdminNotice = """
        WUIntegrate requires administrator permissions in order to use DISM commands. Please run as administrator.

        WUIntegrate will never perform any operations that are against your fundamental privacy rights. WUIntegrate source code is available at the GitHub page.
        """;

    // Runtime settings
    internal static string? TemporaryPath; // WUIntegrate
    internal static string? UtilsPath; // WUIntegrate\utils
    internal static string? SevenZipExe; // WUIntegrate\utils\7za.exe
    internal static string? offlinescancab; // WUIntegrate\offlinescancab
    internal static string? MediumPath; // Input WIM file.
    internal static string? DismMountPath; // WUIntegrate\mount
    internal static string? MediumExtractPath; // WUIntegrate\extract
    internal static string? DownloadedUpdates; // WUIntegrate\updates

    internal static int WimIndex; // Selected WIM index

    static MediumType Medium; // Type of medium we are working with

    internal static UpdateSystem.WindowsVersion WindowsVersion; // Medium OS version
    internal static UpdateSystem.Architecture SystemArchitecture; // Medium architecture

    internal static string? ArgumentPath = null;

    public static void Main(string[] args)
    {
        if (!Helper.IsCurrentUserAdmin())
            Console.WriteLine(AdminNotice);

        ArgumentPath = args.FirstOrDefault();

        Console.WriteLine(StartupNotice);
        ArgumentPath = @"C:\Users\Ryze\Documents\en-us_windows_10_22h2_x64\sources\install.esd";

        // Test Arguments
        if (ArgumentPath == null)
        {
            Console.WriteLine(UsageNotice);
            return;
        }

        // Test Path
        TestPath(ArgumentPath);

        // Test Space
        TestSpace(ArgumentPath);

        // Initializing DISM API
        Console.WriteLine("Initializing DISM API...");
        DismApi.Initialize(DismLogLevel.LogErrorsWarningsInfo);

        // Create Paths
        Console.WriteLine("Creating temporary directories...");
        CreatePaths();

        // Identify Medium
        Console.WriteLine("Identifying medium...");
        IdentifyMedium(ArgumentPath);

        // Extract 7z
        Helper.ExtractResource("WUIntegrate.Utils.7za.dll", Path.Combine(UtilsPath!, "7za.dll"));
        Helper.ExtractResource("WUIntegrate.Utils.7za.exe", Path.Combine(UtilsPath!, "7za.exe"));
        Helper.ExtractResource("WUIntegrate.Utils.7zxa.dll", Path.Combine(UtilsPath!, "7zxa.dll"));

        // Handle different medium types
        switch (Medium)
        {
            case MediumType.IsoFile:
                Console.WriteLine("Extracting ISO file...");
                ExtractISO();
                break;
            case MediumType.WimFile:
                break;
            case MediumType.EsdFile:
                Helper.ExceptionFactory<ArgumentException>("ESD files are not supported.");
                break;
            case MediumType.Unknown:
                Helper.Error("Unknown medium type.");
                break;
        }

        // Get WIM Index
        Console.WriteLine("Getting WIM index selection...");
        SetWinVersionAndArchitecture(MediumPath!);

        // Confirm information
        Console.WriteLine(
            $"""
            WIM Index: {WimIndex}
            Windows Image Version: {WindowsVersion}
            Architecture: {SystemArchitecture}

            WIM Path: {MediumPath}
            DISM Mount Path: {DismMountPath}
            """
        );

        // Mount WIM
        Console.WriteLine("Mounting WIM...");
        MountWIM();

        // Start Update System
        Console.WriteLine("Starting Update System...");
        UpdateSystem updateSystem = new(WindowsVersion, SystemArchitecture);

        // Integrate updates if specified
        if (updateSystem.ReadyToIntegrate)
        {
            IntegrateUpdateFiles();
        }

        // Commit WIM
        Console.WriteLine("Committing WIM...");
        CommitWIM();

        // Cleanup
        Console.WriteLine("Cleaning up...");
        Cleanup();
    }

    // PATH OPERATIONS
    private static void CreatePaths()
    {
        // Make Temporary Path Root
        var temp = Path.GetTempPath();
        var tempDirNew = Path.Combine(temp, "WUIntegrate");
        var tempSevenZipPath = Path.Combine(tempDirNew, "utils");
        var tempDismMountPath = Path.Combine(tempDirNew, "mount");
        var tempMediumExtractPath = Path.Combine(tempDirNew, "extract");
        var tempUpdateDlPath = Path.Combine(tempDirNew, "updates");

        if (!Directory.Exists(tempDirNew))
        {
            Directory.CreateDirectory(tempDirNew);
            Directory.CreateDirectory(tempSevenZipPath);
            Directory.CreateDirectory(tempDismMountPath);
            Directory.CreateDirectory(tempMediumExtractPath);
            Directory.CreateDirectory(tempUpdateDlPath);
        }

        TemporaryPath = tempDirNew;
        DismMountPath = tempDismMountPath;
        MediumExtractPath = tempMediumExtractPath;
        UtilsPath = tempSevenZipPath;
        DownloadedUpdates = tempUpdateDlPath;
        SevenZipExe = Path.Combine(UtilsPath, "7za.exe");
    }

    private static void DeleteTempPath()
    {
        if (TemporaryPath != null)
        {
            Directory.Delete(TemporaryPath, recursive: true);
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
        Console.WriteLine($"Medium identified as {Medium}");
        if (Medium == MediumType.IsoFile)
        {
            MediumPath = Path.Combine(Path.Combine(MediumExtractPath!, "sources"), "install.wim");
            return;
        }

        if (Medium == MediumType.WimFile || Medium == MediumType.EsdFile)
        {
            MediumPath = path;

            var parentPath = Path.GetDirectoryName(path);
            var di = new DirectoryInfo(parentPath!);
            if (di.Attributes.HasFlag(FileAttributes.ReadOnly))
            {
                Helper.ExceptionFactory<UnauthorizedAccessException>(
                    "Directory of the WIM/ESD is read only. This is not supported. Please move the WIM file to a writable directory or drive."
                );
            }
            return;
        }

        Helper.ExceptionFactory<ArgumentException>("The given path was not a WIM/ESD or ISO file.");
    }

    private static void ExtractISO()
    {
        if (File.Exists(SevenZipExe))
        {
            string arguments = $"x \"{MediumPath}\" -o\"{MediumExtractPath}\" -y";

            ProcessStartInfo startInfo = new()
            {
                FileName = SevenZipExe,
                Arguments = arguments,
                CreateNoWindow = true,
                UseShellExecute = false,
            };

            try
            {
                using (Process? process = Process.Start(startInfo))
                {
                    if (process == null)
                        Helper.ExceptionFactory<Exception>("Failed to start 7zr.exe process.");
                    process!.WaitForExit();
                }

                Console.WriteLine("ISO extraction has completed.");
            }
            catch (Exception ex)
            {
                Helper.ExceptionFactory<Exception>($"ISO extraction failed: {ex}");
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

            foreach (var index in ImageInfo)
            {
                Console.WriteLine(
                    $"""
                    [{index.ImageIndex}]
                        {index.ImageName}
                        {index.Architecture}
                    """
                );
            }

            do
            {
                WimIndex = Helper.PromptForInt(
                    "Please select the number for the WIM index you would like to integrate to: "
                );
            } while (!ImageInfo.Any(x => x.ImageIndex == WimIndex));

            var SelectedImage = ImageInfo.First(x => x.ImageIndex == WimIndex);

            SetWindowsVersion(SelectedImage.ProductVersion.Build, SelectedImage.ProductType);
            SetArchitecture(SelectedImage.Architecture);
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
            Console.WriteLine($"Existing mounted image detected: {image.MountPath}");
            DismApi.UnmountImage(image.MountPath, false);
        }

        try
        {
            DismApi.MountImage(MediumPath!, DismMountPath!, WimIndex, false);
            Console.WriteLine($"WIM is now mounted to: {DismMountPath}");
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
            DismProcessorArchitecture.Intel => UpdateSystem.Architecture.x86,
            DismProcessorArchitecture.AMD64 => UpdateSystem.Architecture.x64,
            DismProcessorArchitecture.ARM64 => UpdateSystem.Architecture.ARM64,
            _ => UpdateSystem.Architecture.Unknown,
        };
        if (SystemArchitecture == UpdateSystem.Architecture.Unknown)
        {
            Helper.ExceptionFactory<ArgumentException>($"{architecture} is not supported.");
        }
    }

    private static void SetWindowsVersion(int build, string productType)
    {
        // Server ProductType is ServerNT, EditionID starts with Server
        // Client ProductType is WinNT, EditionID is random
        if (productType == "ServerNT")
        {
            switch (build)
            {
                case 6001:
                    WindowsVersion = UpdateSystem.WindowsVersion.WindowsServer2008;
                    break;
                case 7601:
                    WindowsVersion = UpdateSystem.WindowsVersion.WindowsServer2008R2;
                    break;
                case 9200:
                    WindowsVersion = UpdateSystem.WindowsVersion.WindowsServer2012;
                    break;
                case 9600:
                    WindowsVersion = UpdateSystem.WindowsVersion.WindowsServer2012R2;
                    break;
                case 14393:
                    WindowsVersion = UpdateSystem.WindowsVersion.WindowsServer2016;
                    break;
                case 17763:
                    WindowsVersion = UpdateSystem.WindowsVersion.WindowsServer2019;
                    break;
                case 20348:
                    WindowsVersion = UpdateSystem.WindowsVersion.WindowsServer2022;
                    break;
                case 26100:
                    WindowsVersion = UpdateSystem.WindowsVersion.WindowsServer2025;
                    break;
            }
        }

        if (productType == "WinNT")
        {
            WindowsVersion = build switch
            {
                7601 => UpdateSystem.WindowsVersion.Windows7,
                9600 => UpdateSystem.WindowsVersion.Windows81,
                10240 => UpdateSystem.WindowsVersion.Windows10RTM,
                10586 => UpdateSystem.WindowsVersion.Windows10TH2,
                14393 => UpdateSystem.WindowsVersion.Windows10RS1,
                15063 => UpdateSystem.WindowsVersion.Windows10RS2,
                16299 => UpdateSystem.WindowsVersion.Windows10RS3,
                17134 => UpdateSystem.WindowsVersion.Windows10RS4,
                17763 => UpdateSystem.WindowsVersion.Windows10RS5,
                18362 => UpdateSystem.WindowsVersion.Windows1019H1,
                18363 => UpdateSystem.WindowsVersion.Windows1019H2,
                19041 => UpdateSystem.WindowsVersion.Windows1020H1,
                19042 => UpdateSystem.WindowsVersion.Windows1020H2,
                19043 => UpdateSystem.WindowsVersion.Windows1021H1,
                19044 => UpdateSystem.WindowsVersion.Windows1021H2,
                19045 => UpdateSystem.WindowsVersion.Windows1022H2,
                22000 => UpdateSystem.WindowsVersion.Windows1121H2,
                22621 => UpdateSystem.WindowsVersion.Windows1122H2,
                22631 => UpdateSystem.WindowsVersion.Windows1123H2,
                26100 => UpdateSystem.WindowsVersion.Windows1124H2,
                _ => UpdateSystem.WindowsVersion.Unknown,
            };
        }

        if (WindowsVersion == UpdateSystem.WindowsVersion.Unknown)
        {
            Helper.ExceptionFactory<ArgumentException>("Windows version is unsupported.");
        }
    }

    private static void IntegrateUpdateFiles()
    {
        // Use DISM update API to integrate all files within the update folder

        var session = DismApi.OpenOfflineSession(DismMountPath!);

        var updateFiles = Directory.GetFiles(DownloadedUpdates!);
        foreach (var updateFile in updateFiles)
        {
            try
            {
                DismApi.AddPackage(session, updateFile, false, false);
            }
            catch (DismException ex)
            {
                Console.WriteLine($"Failed to integrate update: {ex}");
            }
        }

        DismApi.CloseSession(session);
    }

    private static void CommitWIM()
    {
        try
        {
            DismApi.UnmountImage(DismMountPath!, true);
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

    private static void DisplayRunningOSInformation() { }
}
