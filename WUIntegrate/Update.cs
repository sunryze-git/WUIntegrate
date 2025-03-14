using System.Xml.Linq;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;
using System.Text.Json;

namespace WUIntegrate
{

    // What the fuck is the update folder??

    // package?? i dont know yet, but within them:

    // C is the core folder. dunno wtf it is.
    // E is Eula.
    // L is Localized Properties speciifcally the title + KB, and the Description.
    // X is Cabinet schema related shit.


    // OsVersion Meanings:
    // 2008: Windows Server 2008
    // 2008R2: Windows Server 2008 R2
    // 2012: Windows Server 2012
    // 2012R2: Windows Server 2012 R2
    // 2016: Windows Server 2016
    // 2019: Windows Server 2019
    // 2022: Windows Server 2022
    // 2025: Windows Server 2025
    // 7601: Windows 7
    // 9600: Windows 8.0
    // 9601: Windows 8.1
    // 1507: Windows 10 1507
    // 1511: Windows 10 1511
    // 1607: Windows 10 1607
    // 1703: Windows 10 1703
    // 1709: Windows 10 1709
    // 1803: Windows 10 1803
    // 1809: Windows 10 1809
    // 1903: Windows 10 1903
    // 1909: Windows 10 1909
    // 2004: Windows 10 20H1
    // 20H2: Windows 10 20H2
    // 21H1: Windows 10 21H1
    // 10-21H2: Windows 10 21H2
    // 10-22H2: Windows 10 22H2
    // 11-21H2: Windows 11 21H2
    // 11-22H2: Windows 11 22H2
    // 23H2: Windows 11 23H2
    // 24H2: Windows 11 24H2

    struct Update
    {
        public DateTime? CreationDate { get; set; }
        public string? UpdateId { get; set; }
        public int? RevisionId { get; set; }
        public string? KbNumber { get; set; }
        public UpdateSystem.WindowsVersion? OsVersion { get; set; }
        public UpdateSystem.Architecture? Architecture { get; set; }
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

    class UpdateSystem
    {
        internal enum WindowsVersion
        {
            Unknown,
            Windows7,
            Windows81,
            Windows10RTM,
            Windows10TH1,
            Windows10TH2,
            Windows10RS1,
            Windows10RS2,
            Windows10RS3,
            Windows10RS4,
            Windows10RS5,
            Windows1019H1,
            Windows1019H2,
            Windows1020H1,
            Windows1020H2,
            Windows1021H1,
            Windows1021H2,
            Windows1022H2,
            Windows1121H2,
            Windows1122H2,
            Windows1123H2,
            Windows1124H2,
            WindowsServer2008,
            WindowsServer2008R2,
            WindowsServer2012,
            WindowsServer2012R2,
            WindowsServer2016,
            WindowsServer2019,
            WindowsServer2022,
            WindowsServer2025
        }

        internal enum Architecture
        {
            x64,
            x86,
            ARM64,
            Unknown
        }

        const string OFFLINE_CAB = @"https://wsusscn2.cab";

        string? indexXmlPath;
        string? packageXmlPath;
        string? localizationPath;

        private readonly string cabinetDownloadPath = Path.Combine(WUIntegrateRoot.TemporaryPath!, "wsusscn2.cab");

        private readonly Dictionary<int, Update> Updates = [];
        private readonly LookupTable lookupTable = new();

        private readonly WindowsVersion isoVersion;
        private readonly Architecture isoArchitecture;

        public bool ReadyToIntegrate;
        public UpdateSystem(WindowsVersion windowsVersion, Architecture architecture)
        {
            isoVersion = windowsVersion;
            isoArchitecture = architecture;

            Console.WriteLine("Downloading offline update scan cabinet...");
            DownloadOfflineCab();

            Console.WriteLine("Extracting offline update scan cabinet...");
            ExtractScanCabinet();

            Console.WriteLine("Extracting package cabinet...");
            ExtractPackageCabinet();

            Console.WriteLine("Loading lookup table...");
            if (indexXmlPath == null) Helper.Error("Index.xml was not detected.");
            LoadLookupTable(indexXmlPath!);

            Console.WriteLine("Extracting localization files...");
            ExtractLocalization();

            Console.WriteLine("Loading package.xml...");
            if (packageXmlPath == null) Helper.Error("Package.xml was not detected.");
            LoadPackageXml(packageXmlPath!);

            Console.WriteLine("Extracting localization files (Defender will make this VERY slow). Please Wait...");
            LoadPackageLocalization();

            Console.WriteLine("Performing cleanup on download files...");
            Cleanup();

            Console.WriteLine("Removing updates without KB number or OS version...");
            RemoveUpdatesWithoutKB();

            Console.WriteLine("Removing updates for other architectures...");
            RemoveOtherArchitectures(isoArchitecture);

            Console.WriteLine("Removing updates from other versions...");
            RemoveUpdatesFromOtherVersions();

            Console.WriteLine("Finding the latest updates...");
            GetLatestUpdates();

            PrintInformation();

            if (Helper.PromptYesNo($"Are you sure you want to integrate updates? There is NO undoing this action."))
            {
                StartUpdateDownload();
                ReadyToIntegrate = true;
            } else
            {
                Console.WriteLine("Updates will not be integrated.");
                ReadyToIntegrate = false;
            }

            Console.WriteLine("Update System has finished integration tasks.");
        }

        private void PrintInformation()
        {
            Console.WriteLine($"""
                You are integrating updates for: {isoVersion}.

                Architecture: {isoArchitecture}.
                Updates To Integrate: {Updates.Count}.
                """);
        }

        private void StartUpdateDownload() {
            // Start the download of the updates
            Task.Run(async () =>
            {
                await DownloadUpdates();
            }).Wait();
        }

        private async Task DownloadUpdates()
        {
            List<string> updatesToGet = new(Updates.Count);
            foreach (var update in Updates)
            {
                if (update.Value.UpdateId != null)
                {
                    updatesToGet.Add(update.Value.UpdateId);
                }
            }

            if (updatesToGet.Count == 0)
            {
                Console.WriteLine("No updates to download.");
                return;
            }

            string downloadDir = Path.Combine(WUIntegrateRoot.DownloadedUpdates!);

            foreach (var updateId in updatesToGet)
            {
                var downloadLinks = await SearchUpdateCatalog(updateId);
                foreach (var link in downloadLinks)
                {
                    string downloadPath = Path.Combine(downloadDir, link.FileName);
                    Helper.DownloadFileUri(link.Url, downloadPath);
                }
            }
        }

        private static async Task<List<DownloadLink>> SearchUpdateCatalog(string updateId)
        {
            Console.WriteLine($"Getting download links for: {updateId}...");

            // Create JSON data
            var updateData = new
            {
                size = 0,
                updateID = updateId,
                uidInfo = updateId
            };

            // Convert to JSON format
            string jsonPost = JsonSerializer.Serialize(updateData, new JsonSerializerOptions { WriteIndented = false });
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

            Console.WriteLine($"Response status code: {response.StatusCode}");
            var responseText = await response.Content.ReadAsStringAsync();

            responseText = responseText.Replace("www.download.windowsupdate", "download.windowsupdate");

            // Use regex to extract the download links (same pattern as PowerShell)
            var linkRegex = new Regex(
                @"(http[s]?\:\/\/(?:dl\.delivery\.mp\.microsoft\.com|(?:catalog\.s\.)?download\.windowsupdate\.com)\/[^\'""]*)");
            var matches = linkRegex.Matches(responseText);

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
            Helper.DeleteFolder(WUIntegrateRoot.offlinescancab!);
        }

        private void DownloadOfflineCab()
        {
            Helper.DownloadFileUri(OFFLINE_CAB, cabinetDownloadPath);
        }

        private void ExtractScanCabinet()
        {
            WUIntegrateRoot.offlinescancab = Path.Combine(WUIntegrateRoot.TemporaryPath!, "offlinescancab");
            Helper.ExtractFile(cabinetDownloadPath, Path.Combine(WUIntegrateRoot.TemporaryPath!, "offlinescancab"));
            Helper.DeleteFile(cabinetDownloadPath);

            indexXmlPath = Path.Combine(WUIntegrateRoot.offlinescancab, "index.xml");
        }

        private void ExtractPackageCabinet()
        {
            var targetCab = Path.Combine(WUIntegrateRoot.offlinescancab!, "package.cab");
            if (!File.Exists(targetCab)) Helper.Error("Package.CAB was not detected.");

            Helper.ExtractFile(targetCab, Path.Combine(WUIntegrateRoot.offlinescancab!, "package"));

            packageXmlPath = Path.Combine(WUIntegrateRoot.offlinescancab!, "package", "package.xml");
        }

        private void ExtractLocalization()
        {
            var destination = Path.Combine(WUIntegrateRoot.offlinescancab!, "localizations");

            if (lookupTable.PackageLookupTable.Count == 0) Helper.Error("Lookup table is empty. Unable to continue.");

            Helper.ExtractFile(Path.Combine(WUIntegrateRoot.offlinescancab!, "package*.cab"), Path.Combine(WUIntegrateRoot.offlinescancab!, destination), @"l\en\*");

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
            Console.WriteLine("Loaded lookup table.");
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

            if (localizationPath == null) Helper.Error("Localization path was not detected.");
            string[] files = Directory.GetFiles(localizationPath!);
            Regex kbRegex = new(@"KB\d{7}", RegexOptions.Compiled);

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
                Match match = kbRegex.Match(title);
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
            Console.WriteLine("Finished loading package.xml.");
        }

        private void RemoveUpdatesFromOtherVersions()
        {
            var keysToRemove = new List<int>();
            foreach (var update in Updates)
            {
                if (update.Value.OsVersion != null && update.Value.OsVersion != isoVersion)
                {
                    keysToRemove.Add(update.Key);
                }
            }
            foreach (var key in keysToRemove)
            {
                Updates.Remove(key);
            }
        }

        private void GetLatestUpdates()
        {
            bool updatesRemoved;

            do
            {
                updatesRemoved = false;
                var updatesToRemove = new List<int>();

                foreach (var update in Updates)
                {
                    if (update.Value.SupersededBy != null)
                    {

                        // Convert the list of strings to a list of integers.
                        List<int?> supersededBy = [];
                        foreach (var i in update.Value.SupersededBy)
                        {
                            if (int.TryParse(i, out int id))
                            {
                                supersededBy.Add(id);
                            }
                        }

                        // Search other updates for the update that supersedes this one.
                        foreach (var otherUpdate in Updates)
                        {
                            if (otherUpdate.Value.RevisionId == null) continue;

                            if (supersededBy.Contains(otherUpdate.Value.RevisionId))
                            {
                                updatesToRemove.Add(update.Key);
                                updatesRemoved = true;
                                break;
                            }
                        }
                    }
                }

                // Remove the updates that have been superseded by another update.
                foreach (var update in updatesToRemove)
                {
                    Updates.Remove(update);
                }

            } while (updatesRemoved);
        }

        private void RemoveOtherArchitectures(Architecture currentArchitecture)
        {
            var updatesToRemove = new List<int>();

            foreach (var update in Updates)
            {
                if (update.Value.Architecture != null && update.Value.Architecture != currentArchitecture)
                {
                    updatesToRemove.Add(update.Key);
                }
            }

            foreach (var update in updatesToRemove)
            {
                Updates.Remove(update);
            }
        }
    }
}