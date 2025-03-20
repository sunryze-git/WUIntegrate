using System.Xml.Linq;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;
using System.Text.Json;

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
        public List<string>? SupersededBy { get; set; }
        public List<string>? Prerequisites { get; set; }
    };

    struct CatalogSearchResult
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string KbNumber { get; set; }
    }

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

        public string? GetFilename(int value)
        {
            var keys = PackageLookupTable.Keys;

            foreach (var key in keys)
            {
                if (key >= value)
                {
                    return PackageLookupTable[key];
                }
            }

            return null;
        }
    }

    partial class UpdateSystem
    {
       
        const string OFFLINE_CAB = @"https://wsusscn2.cab";

        string? indexXmlPath;
        string? packageXmlPath;
        string? localizationPath;

        private readonly string cabinetDownloadPath = Path.Combine(WUIntegrateRoot.Directories!.WuRoot, "wsusscn2.cab");

        private readonly Dictionary<int, Update> Updates = [];
        private readonly LookupTable lookupTable = new();

        private WindowsVersion isoVersion;
        private Architecture isoArchitecture;

        private static readonly JsonSerializerOptions jsonSerializerOptions = new() { WriteIndented = false };

        public bool ReadyToIntegrate;


        [GeneratedRegex(@"(http[s]?\:\/\/(?:dl\.delivery\.mp\.microsoft\.com|(?:catalog\.s\.)?download\.windowsupdate\.com)\/[^\'""]*)")]
        private static partial Regex DownloadUrlRegex();

        [GeneratedRegex(@"KB\d{7}")]
        private static partial Regex KbNumberRegex();

        public UpdateSystem()
        {

            ConsoleWriter.WriteLine("[i] Downloading offline update scan cabinet...", ConsoleColor.Yellow);
            DownloadOfflineCab();

            ConsoleWriter.WriteLine("[i] Extracting offline update scan cabinet...", ConsoleColor.Yellow);
            ExtractScanCabinet();

            ConsoleWriter.WriteLine("[i] Extracting package cabinet...", ConsoleColor.Yellow);
            ExtractPackageCabinet();

            ConsoleWriter.WriteLine("[i] Loading lookup table...", ConsoleColor.Yellow);
            if (indexXmlPath == null) Helper.ExceptionFactory<FileNotFoundException>("Index.xml was not detected.");
            LoadLookupTable(indexXmlPath!);

            ConsoleWriter.WriteLine("[i] Extracting localization files...", ConsoleColor.Yellow);
            ExtractLocalization();

            ConsoleWriter.WriteLine("[i] Loading package.xml...", ConsoleColor.Yellow);
            if (packageXmlPath == null) Helper.ExceptionFactory<FileNotFoundException>("Package.xml was not detected.");
            LoadPackageXml(packageXmlPath!);

            ConsoleWriter.WriteLine("[i] Extracting localization files (Defender will make this VERY slow). Please Wait...", ConsoleColor.Yellow);
            LoadPackageLocalization();

            ConsoleWriter.WriteLine("[i] Performing cleanup on download files...", ConsoleColor.Yellow);
            Cleanup();

            ConsoleWriter.WriteLine("[i] Removing updates without KB number or OS version...", ConsoleColor.Yellow);
            RemoveUpdatesWithoutKB();
        }

        public void Start(WindowsVersion windowsVersion, Architecture architecture)
        {
            isoVersion = windowsVersion;
            isoArchitecture = architecture;

            // Get updates for this version
            IEnumerable<Update> SpecificVersionUpdates = Updates.Values.Where(u => u.OsVersion == windowsVersion).Intersect(Updates.Values.Where(u => u.Architecture == architecture));

            ConsoleWriter.WriteLine("[i] Finding the latest updates...", ConsoleColor.Yellow);
            var latestUpdates = GetLatestUpdates([.. SpecificVersionUpdates]);

            PrintInformation(latestUpdates);

            if (ConsoleWriter.ChoiceYesNo($"Are you sure you want to integrate updates? There is NO undoing this action.", ConsoleColor.Green))
            {
                StartUpdateDownload(latestUpdates);
                ReadyToIntegrate = true;
            }
            else
            {
                ConsoleWriter.WriteLine("[!] Updates will not be integrated.", ConsoleColor.Red);
                ReadyToIntegrate = false;
            }

            ConsoleWriter.WriteLine("[i] Update System has finished operations.", ConsoleColor.Yellow);
        }

        private void PrintInformation(IEnumerable<Update> customUpdates)
        {
            string bar = new string('-', 20);
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

        private void StartUpdateDownload(IEnumerable<Update> updatesList) {
            // Start the download of the updates
            Task.Run(async () =>
            {
                await DownloadUpdates(updatesList);
            }).Wait();
        }

        private async Task DownloadUpdates(IEnumerable<Update> updatesList)
        {
            List<string> updatesToGet = new(Updates.Count);
            foreach (var update in updatesList)
            {
                if (update.UpdateId != null)
                {
                    updatesToGet.Add(update.UpdateId);
                }
            }

            if (updatesToGet.Count == 0)
            {
                ConsoleWriter.WriteLine("[i] No updates to download.", ConsoleColor.Yellow);
                return;
            }

            string downloadDir = Path.Combine(WUIntegrateRoot.Directories!.DlUpdatesPath);

            foreach (var updateId in updatesToGet)
            {
                var downloadLinks = await SearchUpdateCatalog(updateId);
                foreach (var link in downloadLinks)
                {
                    string downloadPath = Path.Combine(downloadDir, link.FileName);
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
            string jsonPost = JsonSerializer.Serialize(updateData, jsonSerializerOptions);
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
                    string fileName = System.IO.Path.GetFileName(new Uri(url).LocalPath);
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

                    Console.WriteLine($"Found download link: {url}");
                }
            }

            return downloadLinks;
        }

        public static void Cleanup()
        {
            Helper.DeleteFolder(WUIntegrateRoot.Directories!.ScanCabExtPath);
        }

        private void DownloadOfflineCab()
        {
            Helper.DownloadFile(OFFLINE_CAB, cabinetDownloadPath);
        }

        private void ExtractScanCabinet()
        {
            Helper.ExtractFile(cabinetDownloadPath, WUIntegrateRoot.Directories!.ScanCabExtPath);
            Helper.DeleteFile(cabinetDownloadPath);

            indexXmlPath = Path.Combine(WUIntegrateRoot.Directories.ScanCabExtPath, "index.xml");
        }

        private void ExtractPackageCabinet()
        {
            var targetCab = Path.Combine(WUIntegrateRoot.Directories!.ScanCabExtPath, "package.cab");
            if (!File.Exists(targetCab)) Helper.ExceptionFactory<FileNotFoundException>("Package.CAB was not detected.");

            Helper.ExtractFile(targetCab, Path.Combine(WUIntegrateRoot.Directories.ScanCabExtPath, "package"));

            packageXmlPath = Path.Combine(WUIntegrateRoot.Directories.ScanCabExtPath, "package", "package.xml");
        }

        private void ExtractLocalization()
        {
            var destination = Path.Combine(WUIntegrateRoot.Directories!.ScanCabExtPath, "localizations");

            if (lookupTable.PackageLookupTable.Count == 0) Helper.ExceptionFactory<Exception>("Lookup table is empty. Unable to continue.");

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
            var keysToRemove = new List<int>();

            foreach (var update in Updates)
            {
                if (update.Value.KbNumber == null || update.Value.OsVersion == null)
                {
                    keysToRemove.Add(update.Key);
                }
            }

            foreach (var key in keysToRemove)
            {
                Updates.Remove(key);
            }
        }

        private void LoadLookupTable(string path)
        {
            using var reader = XmlReader.Create(path);
            var lookupXml = XDocument.Load(reader);
            if (lookupXml.Root == null) return;

            // Get the first child element children (cablistElements)
            var firstElement = lookupXml.Root.Elements().First();
            if (firstElement == null) return;

            foreach (var element in firstElement.Elements())
            {
                var rangeStartAttr = element.Attribute("RANGESTART");
                var nameAttr = element.Attribute("NAME");

                if (rangeStartAttr == null || nameAttr == null) continue;

                if (int.TryParse(rangeStartAttr.Value, out int rangeStart))
                {
                    lookupTable.Add(rangeStart, nameAttr.Value);
                }
            }
            ConsoleWriter.WriteLine("[i] Loaded update lookup table.", ConsoleColor.Yellow);
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

            if (localizationPath == null) Helper.ExceptionFactory<DirectoryNotFoundException>("Localization path was not detected.");
            string[] files = Directory.GetFiles(localizationPath!);

            // Pre-compile the regex patterns for performance
            Dictionary<WindowsVersion, Regex> compiledOsVersionPatterns = [];
            foreach (var p in OsVersionPatterns)
            {
                compiledOsVersionPatterns.Add(p.Value, new Regex(p.Key, RegexOptions.Compiled));
            }

            Dictionary<Architecture, Regex> compiledArchitecturePatterns = [];
            foreach (var p in ArchitecturePatterns)
            {
                compiledArchitecturePatterns.Add(p.Value, new Regex(p.Key, RegexOptions.Compiled));
            }

            foreach (string file in files)
            {
                if (!int.TryParse(Path.GetFileName(file), out int filename)) continue;
                if (!Updates.ContainsKey(filename)) continue;

                using XmlReader reader = XmlReader.Create(file);
                XDocument xmlDoc = XDocument.Load(reader);

                if (xmlDoc.Root == null) continue;

                var titleElement = xmlDoc.Root.Elements("Title").FirstOrDefault();
                if (titleElement == null) continue;

                string title = titleElement.Value;

                if (title == "Driver" || title.Contains("Office") || title.Contains("SQL")) continue;

                if (!xmlDoc.Root.Elements("Description").Any()) continue;

                // Get the KB number
                Match match = KbNumberRegex().Match(title);
                if (!match.Success) continue;
                string kbNumber = match.Value;

                // Get OS Version and Architecture
                WindowsVersion? osVersion = null;
                Architecture? architecture = null;

                foreach (var p in compiledOsVersionPatterns)
                {
                    match = p.Value.Match(title);
                    if (match.Success)
                    {
                        osVersion = p.Key;
                        break;
                    }
                }

                foreach (var p in compiledArchitecturePatterns)
                {
                    match = p.Value.Match(title);
                    if (match.Success)
                    {
                        architecture = p.Key;
                        break;
                    }
                    else // if no architecture is specified it is probably x86, this is done in the .NET updates
                    {
                        architecture = Architecture.x86;
                    }
                }

                // Update the dictionary entry
                var updateObject = Updates[filename];
                updateObject.KbNumber = kbNumber;
                updateObject.OsVersion = osVersion;
                updateObject.Architecture = architecture;
                Updates[filename] = updateObject;
            }
        }

        private void LoadPackageXml(string path)
        {
            using var reader = XmlReader.Create(path);
            var packageXML = XDocument.Load(reader);
            if (packageXML.Root == null) return;

            XElement root = packageXML.Root;
            XNamespace ns = "http://schemas.microsoft.com/msus/2004/02/OfflineSync";

            var updatesElement = root.Elements().First().Elements();

            // For each Update in the Updates element
            foreach (var update in updatesElement)
            {
                DateTime? parsedDate = null;
                if (DateTime.TryParse(update.Attribute("CreationDate")?.Value, out var date))
                {
                    parsedDate = date;
                }

                var revisionIdAttribute = update.Attribute("RevisionId");
                if (revisionIdAttribute == null) continue;

                int revisionId = Int32.Parse(revisionIdAttribute.Value);

                // Pre-size the list if count is available
                var supersededBy = update.Elements(ns + "SupersededBy").Elements().ToList();
                var supersededUpdateIds = new List<string>(supersededBy.Count);

                foreach (var element in supersededBy)
                {
                    var idAttribute = element.Attribute("Id");
                    if (idAttribute != null)
                    {
                        supersededUpdateIds.Add(idAttribute.Value);
                    }
                }

                // Do the same for if there is prerequsites
                var prerequisites = update.Elements(ns + "Prerequisites").Elements().ToList();
                var prerequisiteIds = new List<string>(prerequisites.Count);

                foreach (var element in prerequisites)
                {
                    var pIdAttribute = element.Attribute("Id");
                    if (pIdAttribute != null)
                    {
                        prerequisiteIds.Add(pIdAttribute.Value);
                    }
                }

                Updates[revisionId] = new Update
                {
                    UpdateId = update.Attribute("UpdateId")?.Value,
                    RevisionId = revisionId,
                    CreationDate = parsedDate,
                    SupersededBy = supersededUpdateIds,
                    Prerequisites = prerequisiteIds
                };
            }
            ConsoleWriter.WriteLine("[i] Finished loading package.xml.", ConsoleColor.Yellow);
        }

        private static int? TryParseInt(string value)
        {
            return int.TryParse(value, out int result) ? (int?)result : null;
        }

        private static IEnumerable<Update> GetLatestUpdates(IEnumerable<Update> customUpdates)
        {
            var latestUpdates = customUpdates.ToList();

            var updatesToRemove = new HashSet<Update>();

            foreach (var update in latestUpdates)
            {
                if (update.SupersededBy != null)
                {
                    var supersededBy = update.SupersededBy
                        .Select(i => TryParseInt(i))
                        .Where(id => id.HasValue)
                        .Select(id => id!.Value)
                        .ToList();

                    foreach (var otherUpdate in latestUpdates)
                    {
                        if (otherUpdate.RevisionId == null) continue;
                        if (supersededBy.Contains(otherUpdate.RevisionId.Value))
                        {
                            updatesToRemove.Add(update);
                            break;
                        }
                    }
                }
            }

            return latestUpdates.Where(update => !updatesToRemove.Contains(update));
        }

    }

}

