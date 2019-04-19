using GenerateMasterDataForExcel.IOFiles.IOExcel;
using GenerateMasterDataForExcel.IOFilesProvider;
using GenerateMasterDataForExcel.ReadFiles;

namespace GenerateMasterDataForExcel.ServiceLocators
{
    public sealed class ServiceLocatorProvider
    {
        private static ServiceLocatorProvider _singleton = new ServiceLocatorProvider();
        public ServiceLocator ReadExcelCurrent { get; private set; }
        public ServiceLocator IoFileNameCurrent { get; private set; }
        public ServiceLocator ioExcelFileCurrent { get; private set; }

        public static ServiceLocatorProvider GetInstance
        {
            get { return _singleton; }
        }

        private ServiceLocatorProvider()
        {
            // IOFileNameの依存関係を登録
            IoFileNameCurrent = new ServiceLocator();
            IoFileNameCurrent.Register<IIOFileNamesProvider>(new IOFileName());

            // ReadExcelFilesの依存関係を登録
            ReadExcelCurrent = new ServiceLocator();
            ReadExcelCurrent.Register<IReadExcelFiles>(new ReadExcelFiles());

            // IOExcelFilesの依存関係を登録
            ioExcelFileCurrent = new ServiceLocator();
            ioExcelFileCurrent.Register<IIOExcelFilesProvider>(new IOExcelFiles());
        }
    }
}