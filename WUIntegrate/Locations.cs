namespace WUIntegrate
{
    public class Locations()
    {
        public required string SysTemp { get; set; }
        public required string WuRoot { get; set; }
        public required string ScanCabExtPath { get; set; }
        public required string? MediumPath { get; set; }
        public required string DismMountPath { get; set; }
        public required string MediumExtractPath { get; set; }
        public required string DlUpdatesPath { get; set; }

        public IEnumerable<string> All
        {
            get
            {
                yield return SysTemp;
                yield return WuRoot;
                yield return ScanCabExtPath;
                yield return MediumPath!;
                yield return DismMountPath;
                yield return MediumExtractPath;
                yield return DlUpdatesPath;
            }
        }
    }
}
