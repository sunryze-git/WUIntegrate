namespace WUIntegrate
{
    class Locations(string SysTemp, string WuRoot, string UtilPath, string ScanCabExtPath, string DismMountPath, string MediumExtractPath, string DlUpdatesPath)
    {
        public string SysTemp { get; } = SysTemp;
        public string WuRoot { get; } = WuRoot;
        public string UtilPath { get; } = UtilPath;
        public string? SevenZipExe { get; set; }
        public string ScanCabExtPath { get; } = ScanCabExtPath;
        public string? MediumPath { get; set; }
        public string DismMountPath { get; } = DismMountPath;
        public string MediumExtractPath { get; } = MediumExtractPath;
        public string DlUpdatesPath { get; } = DlUpdatesPath;

        public IEnumerable<string> All
        {
            get
            {
                yield return SysTemp;
                yield return WuRoot;
                yield return UtilPath;
                yield return SevenZipExe!;
                yield return ScanCabExtPath;
                yield return MediumPath!;
                yield return DismMountPath;
                yield return MediumExtractPath;
                yield return DlUpdatesPath;
            }
        }
    }
}
