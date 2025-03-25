using System.Xml.Linq;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace WUIntegrate
{

    struct Update
    {
        public DateTime? CreationDate { get; set; }
        public string? UpdateId { get; set; }
        public int? RevisionId { get; set; }
        public string? KbNumber { get; set; }
        public WindowsVersion? OsVersion { get; set; }
        public Architecture? Architecture { get; set; }
        public IEnumerable<string>? SupersededBy { get; set; }
        public IEnumerable<string>? Prerequisites { get; set; }
    };

    struct DownloadLink
    {
        public string DownloadId { get; set; }
        public string Url { get; set; }
        public string FileName { get; set; }
    }

    class LookupTable
    {
        public SortedList<int, string> PackageLookupTable;

        public LookupTable()
        {
            PackageLookupTable = [];
        }

        public void Add(int startValue, string packageName)
        {
            PackageLookupTable[startValue] = packageName;
        }
    }

    partial class UpdateSystem
    {
        private const string OFFLINE_CAB = @"https://wsusscn2.cab";

        private string? indexXmlPath;
        private string? packageXmlPath;
        private string? localizationPath;

        private readonly string cabinetDownloadPath = Path.Combine(WUIntegrateRoot.Directories.WuRoot, "wsusscn2.cab");

        private readonly ConcurrentDictionary<int, Update> Updates = [];
        private readonly LookupTable lookupTable = new();

        private WindowsVersion isoVersion;
        private Architecture isoArchitecture;

        private static readonly JsonSerializerOptions jsonSerializerOptions = new() { WriteIndented = false };

        public bool ReadyToIntegrate { get; set; }

        [GeneratedRegex(@"(http[s]?\:\/\/(?:dl\.delivery\.mp\.microsoft\.com|(?:catalog\.s\.)?download\.windowsupdate\.com)\/[^\'""]*)")]
        private static partial Regex DownloadUrlRegex();

        [GeneratedRegex(@"KB\d{7}")]
        private static partial Regex KbNumberRegex();

        public UpdateSystem()
        {

            Logger.Msg("Downloading offline update scan cabinet file.");
            DownloadOfflineCab();

            Logger.Msg("Extracting offline update scan cabinet file.");
            ExtractScanCabinet();

            Logger.Msg("Extracting update package cabinet.");
            ExtractPackageCabinet();

            //Logger.Msg("Loading package lookup table.");
            //if (indexXmlPath is null)
            //{
            //    Helper.ExceptionFactory<FileNotFoundException>("Index.xml was not detected.");
            //    return;
            //}
            //LoadLookupTable(indexXmlPath!);

            Logger.Msg("Extracting package localization files.");
            ExtractLocalization();

            Logger.Msg("Loading package XML.");
            if (packageXmlPath is null)
            {
                Helper.ExceptionFactory<FileNotFoundException>("Package.xml was not detected.");
                return;
            }
            LoadPackageXml(packageXmlPath!);

            Logger.Msg("Loading localization files.");
            LoadPackageLocalization();

            Logger.Msg("Cleaning up downloaded files.");
            Cleanup();

            Logger.Msg("Removing bad updates.");
            RemoveUpdatesWithoutKB();
        }

        public void Start(WindowsVersion windowsVersion, Architecture architecture)
        {
            isoVersion = windowsVersion;
            isoArchitecture = architecture;

            Logger.Msg("Finding the latest updates...");
            var latestUpdates = GetLatestUpdates(Updates.Values
                .Where(u => u.OsVersion == windowsVersion && u.Architecture == architecture));

            PrintInformation(latestUpdates);

            if (ConsoleWriter.ChoiceYesNo($"Are you sure you want to integrate updates? There is NO undoing this action.", ConsoleColor.Green))
            {
                StartUpdateDownload(latestUpdates);
                ReadyToIntegrate = true;
            }
            else
            {
                Logger.Warn("Updates will not be integrated.");
                ReadyToIntegrate = false;
            }

            Logger.Msg("Update system has finished operations.");
        }

        private void PrintInformation(IEnumerable<Update> customUpdates)
        {
            Logger.Log($"Integrating updates for: {isoVersion}, Count: {customUpdates.Count()}.");
            var bar = new string('-', 20);
            ConsoleWriter.WriteLine($"""
                You are integrating updates for: {isoVersion}.

                Architecture: {isoArchitecture}.
                Updates To Integrate: {customUpdates.Count()}.
                """, ConsoleColor.White);
            ConsoleWriter.WriteLine(bar, ConsoleColor.Blue);
            foreach (var update in customUpdates)
            {
                ConsoleWriter.WriteLine($"""
                    KB Number: {update.KbNumber}.
                    Release Date: {update.CreationDate}.
                    OS Version: {update.OsVersion}.
                    Architecture: {update.Architecture}.
                    """, ConsoleColor.White);
                ConsoleWriter.WriteLine(bar, ConsoleColor.Blue);
            }
        }

        private static void StartUpdateDownload(IEnumerable<Update> updatesList) {
            // Start the download of the updates
            Task.Run(async () =>
            {
                await DownloadUpdates(updatesList);
            }).Wait();
        }

        private static async Task DownloadUpdates(IEnumerable<Update> updatesList)
        {
            List<string> updatesToGet = [.. updatesList
                .Where(x => x.UpdateId is not null)
                .Select(x => x.UpdateId)];

            if (updatesToGet.Count == 0)
            {
                Logger.Msg("No updates to download.");
                return;
            }

            var downloadDir = Path.Combine(WUIntegrateRoot.Directories.DlUpdatesPath);

            foreach (var updateId in updatesToGet)
            {
                var downloadLinks = await SearchUpdateCatalog(updateId);
                foreach (var link in downloadLinks)
                {
                    var downloadPath = Path.Combine(downloadDir, link.FileName);
                    Helper.DownloadFile(link.Url, downloadPath);
                }
            }
        }

        private static async Task<List<DownloadLink>> SearchUpdateCatalog(string updateId)
        {
            // Create JSON data
            var updateData = new
            {
                size = 0,
                updateID = updateId,
                uidInfo = updateId
            };

            // Convert to JSON format
            var jsonPost = JsonSerializer.Serialize(updateData, jsonSerializerOptions);
            var requestContent = $"updateIDs=[{jsonPost}]";

            // Create HTTP client
            using var httpClient = new HttpClient();
            var content = new StringContent(
                requestContent,
                Encoding.UTF8,
                "application/x-www-form-urlencoded"
            );

            // Make POST request
            var response = await httpClient.PostAsync(
                "https://www.catalog.update.microsoft.com/DownloadDialog.aspx",
                content
            );

            // Get response content
            var responseText = await response.Content.ReadAsStringAsync();

            responseText = responseText.Replace("www.download.windowsupdate", "download.windowsupdate");

            // Use regex to extract the download links (same pattern as PowerShell)
            var matches = DownloadUrlRegex().Matches(responseText);

            // Create the result list
            var downloadLinks = new List<DownloadLink>();

            foreach (Match match in matches)
            {
                if (match.Success)
                {
                    var url = match.Groups[1].Value;

                    // Extract filename from URL
                    var fileName = Path.GetFileName(new Uri(url).LocalPath);
                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = $"update_{Guid.NewGuid()}.cab";
                    }

                    downloadLinks.Add(new DownloadLink
                    {
                        DownloadId = updateId, // Using the update GUID as the download ID
                        Url = url,
                        FileName = fileName
                    });
                }
            }

            return downloadLinks;
        }

        public static void Cleanup()
        {
            Helper.DeleteFolder(WUIntegrateRoot.Directories.ScanCabExtPath);
        }

        private void DownloadOfflineCab()
        {
            Helper.DownloadFile(OFFLINE_CAB, cabinetDownloadPath);
        }

        private void ExtractScanCabinet()
        {
            Helper.ExtractFile(cabinetDownloadPath, WUIntegrateRoot.Directories.ScanCabExtPath);
            Helper.DeleteFile(cabinetDownloadPath);

            indexXmlPath = Path.Combine(WUIntegrateRoot.Directories.ScanCabExtPath, "index.xml");
        }

        private void ExtractPackageCabinet()
        {
            var targetCab = Path.Combine(WUIntegrateRoot.Directories.ScanCabExtPath, "package.cab");
            if (!File.Exists(targetCab))
            {
                Helper.ExceptionFactory<FileNotFoundException>("Package.CAB was not detected.");
            }

            Helper.ExtractFile(targetCab, Path.Combine(WUIntegrateRoot.Directories.ScanCabExtPath, "package"));

            packageXmlPath = Path.Combine(WUIntegrateRoot.Directories.ScanCabExtPath, "package", "package.xml");
        }

        private void ExtractLocalization()
        {
            var destination = Path.Combine(WUIntegrateRoot.Directories.ScanCabExtPath, "localizations");

            foreach (var file in Directory.GetFiles(WUIntegrateRoot.Directories.ScanCabExtPath))
            {
                var fileObject = new FileInfo(file);
                // if file is not a directory, is a .cab file, and is not package.cab
                if (fileObject.Extension == ".cab" && fileObject.Name != "package.cab")
                {
                    Helper.ExtractFile(file, destination, @"l\en");
                }
            }

            localizationPath = Path.Combine(destination, "l", "en");
        }

        private void RemoveUpdatesWithoutKB()
        {
            var updatesToRemove = Updates
                .Where(x => x.Value.KbNumber is null || x.Value.OsVersion is null);

            foreach (var key in updatesToRemove)
            {
                if (!Updates.TryRemove(key))
                {
                    Logger.Warn($"Unable to remove update {key}");
                }
            }
        }

        private void LoadLookupTable(string path)
        {
            using var reader = XmlReader.Create(path);
            var lookupXml = XDocument.Load(reader);
            if (lookupXml.Root is null)
            {
                return;
            }

            // Get the first child element children (cablistElements)
            var firstElement = lookupXml.Root.Elements().First();
            if (firstElement is null)
            {
                return;
            }

            foreach (var element in firstElement.Elements())
            {
                var rangeStartAttr = element.Attribute("RANGESTART");
                var nameAttr = element.Attribute("NAME");

                if (rangeStartAttr is null || nameAttr is null)
                {
                    continue;
                }

                if (int.TryParse(rangeStartAttr.Value, out int rangeStart))
                {
                    lookupTable.Add(rangeStart, nameAttr.Value);
                }
            }
            Logger.Msg("Loaded update lookup table.");
        }

        private void LoadPackageLocalization()
        {
            Dictionary<string, WindowsVersion> OsVersionPatterns = new()
            {
                ["Windows Server 2008"] = WindowsVersion.WindowsServer2008,
                ["Windows Server 2008 R2"] = WindowsVersion.WindowsServer2008R2,
                ["Windows Server 2012"] = WindowsVersion.WindowsServer2012,
                ["Windows Server 2012 R2"] = WindowsVersion.WindowsServer2012R2,
                ["Windows Server 2016"] = WindowsVersion.WindowsServer2016,
                ["Windows Server 2019"] = WindowsVersion.WindowsServer2019,
                ["Windows Server 2022"] = WindowsVersion.WindowsServer2022,
                ["Windows Server 2025"] = WindowsVersion.WindowsServer2025,
                ["Windows 7"] = WindowsVersion.Windows7,
                ["Windows 8.1"] = WindowsVersion.Windows81,
                ["Windows 10 Version 1507"] = WindowsVersion.Windows10RTM,
                ["Windows 10 Version 1511"] = WindowsVersion.Windows10TH1,
                ["Windows 10 Version 1607"] = WindowsVersion.Windows10TH2,
                ["Windows 10 Version 1703"] = WindowsVersion.Windows10RS2,
                ["Windows 10 Version 1709"] = WindowsVersion.Windows10RS3,
                ["Windows 10 Version 1803"] = WindowsVersion.Windows10RS4,
                ["Windows 10 Version 1809"] = WindowsVersion.Windows10RS5,
                ["Windows 10 Version 1903"] = WindowsVersion.Windows1019H1,
                ["Windows 10 Version 1909"] = WindowsVersion.Windows1019H2,
                ["Windows 10 Version 2004"] = WindowsVersion.Windows1020H1,
                ["Windows 10 Version 20H2"] = WindowsVersion.Windows1020H2,
                ["Windows 10 Version 21H1"] = WindowsVersion.Windows1021H1,
                ["Windows 10 Version 21H2"] = WindowsVersion.Windows1021H2,
                ["Windows 10 Version 22H2"] = WindowsVersion.Windows1022H2,
                ["Windows 11 Version 21H2"] = WindowsVersion.Windows1121H2,
                ["Windows 11 Version 22H2"] = WindowsVersion.Windows1122H2,
                ["Windows 11 Version 23H2"] = WindowsVersion.Windows1123H2,
                ["Windows 11 Version 24H2"] = WindowsVersion.Windows1124H2

            };

            Dictionary<string, Architecture> ArchitecturePatterns = new()
            {
                ["x64"] = Architecture.x64,
                ["x86"] = Architecture.x86,
                ["ARM64"] = Architecture.ARM64
            };

            if (localizationPath is null)
            {
                Helper.ExceptionFactory<DirectoryNotFoundException>("Localization path was not detected.");
            }
            var files = Directory.GetFiles(localizationPath!);

            // Pre-compile the regex patterns for performance
            Dictionary<WindowsVersion, Regex> compiledOsVersionPatterns = OsVersionPatterns
                .ToDictionary(p => p.Value, p => new Regex(p.Key, RegexOptions.Compiled));

            Dictionary<Architecture, Regex> compiledArchitecturePatterns = ArchitecturePatterns
                .ToDictionary(p => p.Value, p => new Regex(p.Key, RegexOptions.Compiled));

            foreach (string file in files)
            {
                if (!int.TryParse(Path.GetFileName(file), out int filename) || !Updates.ContainsKey(filename))
                {
                    continue;
                }

                using XmlReader reader = XmlReader.Create(file);
                XDocument xmlDoc = XDocument.Load(reader);

                if (xmlDoc.Root is null)
                {
                    continue;
                }

                var titleElement = xmlDoc.Root.Elements("Title").FirstOrDefault();
                if (titleElement is null)
                {
                    continue;
                }

                var title = titleElement.Value;

                if (title == "Driver" || title.Contains("Office") || title.Contains("SQL"))
                {
                    continue;
                }

                if (!xmlDoc.Root.Elements("Description").Any())
                {
                    continue;
                }

                // Get the KB number
                var match = KbNumberRegex().Match(title);
                if (!match.Success)
                {
                    continue;
                }
                var kbNumber = match.Value;

                // Get OS Version and Architecture
                WindowsVersion osVersion = WindowsVersion.Unknown;
                Architecture architecture = Architecture.x86;

                osVersion = compiledOsVersionPatterns
                    .FirstOrDefault(p => p.Value.IsMatch(title))
                    .Key;

                architecture = compiledArchitecturePatterns
                    .FirstOrDefault(p => p.Value.IsMatch(title))
                    .Key;

                if (Updates.TryGetValue(filename, out var update))
                {
                    update.KbNumber = kbNumber;
                    update.OsVersion = osVersion;
                    update.Architecture = architecture;

                    Updates[filename] = update;
                }
            }

            GC.Collect();
        }

        private void LoadPackageXml(string path)
        {
            using var reader = XmlReader.Create(path);
            var packageXML = XDocument.Load(reader);
            if (packageXML.Root is null) return;

            var root = packageXML.Root;
            var ns = "http://schemas.microsoft.com/msus/2004/02/OfflineSync";

            var updatesElement = root.Elements().First().Elements();

            Parallel.ForEach(updatesElement, update =>
            {
                if (!DateTime.TryParse(update.Attribute("CreationDate")?.Value, out var parsedDate))
                {
                    return;
                }

                var revisionIdAttribute = update.Attribute("RevisionId");
                if (revisionIdAttribute is null)
                {
                    return;
                }

                if (!Int32.TryParse(revisionIdAttribute.Value, out var revisionId))
                {
                    return;
                }

                // Pre-size the list if count is available
                var supersededBy = update.Elements(XName.Get("SupersededBy", ns)).Elements();
                var supersededUpdateIds = supersededBy // 
                    .Where(x => x.Attribute("Id") is not null)
                    .Select(x => x.Attribute("Id")!.Value);

                // Do the same for if there is prerequsites
                var prerequisites = update.Elements(XName.Get("Prerequisites", ns)).Elements();
                var prerequisiteIds = prerequisites
                    .Where(x => x.Attribute("Id") is not null)
                    .Select(x => x.Attribute("Id")!.Value);

                // Ensure we don't collide
                if (Updates.ContainsKey(revisionId))
                {
                    return;
                }

                Updates[revisionId] = new Update
                {
                    UpdateId = update.Attribute("UpdateId")?.Value,
                    RevisionId = revisionId,
                    CreationDate = parsedDate,
                    SupersededBy = supersededUpdateIds,
                    Prerequisites = prerequisiteIds
                };
            });

            GC.Collect();
            Logger.Msg("Finished loading package.XML.");
        }

        private static int? TryParseInt(string value)
        {
            return Int32.TryParse(value, out int result) ? (int?)result : null;
        }

        private static IEnumerable<Update> GetLatestUpdates(IEnumerable<Update> customUpdates)
        {
            var updatesToRemove = new HashSet<Update>();

            foreach (var update in customUpdates)
            {
                if (update.SupersededBy is not null)
                {
                    var supersededBy = update.SupersededBy
                        .Select(i => TryParseInt(i))
                        .Where(id => id.HasValue)
                        .Select(id => id!.Value);

                    foreach (var otherUpdate in customUpdates)
                    {
                        if (otherUpdate.RevisionId is null) continue;
                        if (supersededBy.Contains(otherUpdate.RevisionId.Value))
                        {
                            updatesToRemove.Add(update);
                            break;
                        }
                    }
                }
            }

            return customUpdates.Where(update => !updatesToRemove.Contains(update));
        }

    }

}

