using System.Runtime.InteropServices;
using System.Security.Principal;
using SevenZipExtractor;
using Vanara;

namespace WUIntegrate
{
    public class Helper
    {

        public static void DeleteFile(string fileName)
        {
            if (!File.Exists(fileName))
            {
                return;
            }
            File.Delete(fileName);
        }

        public static void DeleteFolder(string folderName)
        {
            if (!Directory.Exists(folderName))
            {
                return;
            }
            Directory.Delete(folderName, recursive: true);
        }

        public static void ExtractFile(string SourcePath, string DestinationPath, string? SpecificFolderName = null)
        {
            if (SpecificFolderName is null) // Normal Extraction
            {
                Logger.Msg($"Extracting file {SourcePath} to {DestinationPath}");
                using ArchiveFile archiveFile = new(SourcePath);
                archiveFile.Extract(DestinationPath);
            }
            else
            {
                // Determine subfolders as well for the specific folder
                using ArchiveFile archiveFile = new(SourcePath);

                archiveFile.Extract(entry =>
                {
                    if (entry.FileName.StartsWith(SpecificFolderName, StringComparison.OrdinalIgnoreCase))
                    {
                        string targetFullPath = Path.Combine(DestinationPath, entry.FileName);
                        return targetFullPath;
                    }
                    return null;
                }, DestinationPath);
            }
        }

        public static void DownloadFile(string Url, string DestinationPath)
        {
            int attempts = 0;
            int maxRetries = 5;
            while (attempts < maxRetries)
            {
                try
                {
                    if (Url is null)
                    {
                        ExceptionFactory<ArgumentNullException>("URL to be downloaded was null.");
                    }

                    Task.Run(async () =>
                    {
                        Logger.Log($"Downloading file from {Url} to {DestinationPath}");
                        using HttpClient client = new();
                        using HttpResponseMessage response = await client.GetAsync(Url);
                        using Stream responseStream = await response.Content.ReadAsStreamAsync();
                        using Stream destinationStream = new FileStream(
                            DestinationPath,
                            FileMode.Create,
                            FileAccess.Write,
                            FileShare.None,
                            bufferSize: 81920,
                            useAsync: true
                        );
                        response.EnsureSuccessStatusCode();

                        await responseStream.CopyToAsync(destinationStream);

                        Logger.Log($"Downloaded file from {Url} to {DestinationPath}");
                        responseStream.Close();

                        destinationStream.Flush();
                        destinationStream.Close();
                    })
                        .Wait();
                    attempts = maxRetries;
                }
                catch (HttpRequestException ex)
                {
                    attempts++;
                    Logger.Warn($"Failed to download file: {ex.Message}");
                    if (attempts >= maxRetries)
                    {
                        ExceptionFactory<HttpRequestException>($"Failed to download file after {maxRetries} attempts: {ex.Message}");
                    }
                }
            }
        }

        public static bool IsCurrentUserAdmin()
        {
            using WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        public static void ExceptionFactory<T>(string errorMessage)
            where T : Exception
        {
            UpdateSystem.Cleanup();
            WUIntegrateRoot.Cleanup();
            var ex =
                Activator.CreateInstance(typeof(T), errorMessage) as T
                ?? throw new InvalidOperationException("Failed to create exception instance.");

            Logger.Crash(ex);

            throw ex;
        }
    }
}
