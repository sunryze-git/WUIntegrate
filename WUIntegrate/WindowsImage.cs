using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Dism;

namespace WUIntegrate
{
    public class WindowsImage
    {
        public IReadOnlyCollection<DismImageInfo> Indexes { get
            {
                return DismApi.GetImageInfo(ImagePath);
            }
        }


        private readonly string ImagePath;
        private readonly string MountPath;

        private DismSession? DismSession;

        WindowsImage(string wimPath, string mountPath)
        {
            ImagePath = wimPath;
            MountPath = mountPath;
            DismApi.Initialize(DismLogLevel.LogErrorsWarningsInfo);
        }

        public void IntegrateUpdates(string path)
        {
            var pathIsDirectory = Directory.Exists(path);
            var pathIsFile = File.Exists(path);
            var pathContainsFiles = Directory.EnumerateFiles(path).Any();

            if (pathIsDirectory && !pathIsFile && pathContainsFiles)
            {
                try
                {
                    if (DismSession is null)
                    {
                        DismSession = DismApi.OpenOfflineSession(ImagePath);
                    }

                    foreach (var file in Directory.EnumerateFiles(path))
                    {
                        DismApi.AddPackage(DismSession, file, false, true);
                    }

                    DismApi.CloseSession(DismSession);
                }
                catch (DismException ex)
                {
                    Logger.Error($"Failed to integrate updates: {ex.Message}");
                }

                return;
            }

            Logger.Error("Invalid path to updates directory.");
        }

        public void Mount(int imageIndex)
        {
            try
            {
                DismApi.MountImage(ImagePath, MountPath, imageIndex);
            }
            catch (DismException ex)
            {
                Logger.Error($"Failed to mount image: {ex.Message}");   
            }
        }

        public void Unmount(bool commitChanges)
        {
            try
            {
                DismApi.UnmountImage(MountPath, commitChanges);
            }
            catch (DismException ex)
            {
                Logger.Error($"Failed to unmount image: {ex.Message}");
            }
        }
    }
}
