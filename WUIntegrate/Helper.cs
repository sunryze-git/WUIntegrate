using System.Security.Principal;
using SevenZipExtractor;

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
                var specificFiles = archiveFile.Entries
                    .Where(x => x.FileName.StartsWith(SpecificFolderName, StringComparison.OrdinalIgnoreCase))
                    .ToHashSet();

                archiveFile.Extract(entry =>
                {
                    if (specificFiles.Contains(entry))
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
                    if (Path.Exists(DestinationPath))
                    {
                        Logger.Warn($"File {DestinationPath} already exists.");
                        return;
                    }

                    Task.Run(async () =>
                    {
                        using HttpClient client = new();
                        using HttpResponseMessage response = await client.GetAsync(Url);

                        response.EnsureSuccessStatusCode();
                        using Stream responseStream = await response.Content.ReadAsStreamAsync();
                        using Stream destinationStream = new FileStream(
                            DestinationPath,
                            FileMode.Create,
                            FileAccess.Write,
                            FileShare.None
                        );

                        var totalBytes = response.Content.Headers.ContentLength!;

                        var BufferSize = 81920;
                        var buffer = new byte[BufferSize];

                        var totalBytesRead = 0;
                        var bytesRead = 0;
                        while ((bytesRead = await responseStream.ReadAsync(buffer)) > 0)
                        {
                            await destinationStream.WriteAsync(buffer.AsMemory(0, bytesRead));
                            totalBytesRead += bytesRead;

                            var progress = (int)((double)totalBytesRead / totalBytes * 100);

                            ConsoleWriter.WriteProgress(100, progress, 100, "Downloading", ConsoleColor.Magenta);
                        }
                    })
                        .Wait();
                }
                catch (HttpRequestException ex)
                {
                    attempts++;
                    if (attempts >= maxRetries)
                    {
                        ExceptionFactory<HttpRequestException>($"Failed to download file after {maxRetries} attempts: {ex.Message}");
                    }
                }
            }
            
            GC.Collect();
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
