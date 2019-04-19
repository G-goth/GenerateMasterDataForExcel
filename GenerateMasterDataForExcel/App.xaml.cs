using GenerateMasterDataForExcel.IOFilesProvider;
using GenerateMasterDataForExcel.ReadFiles;
using GenerateMasterDataForExcel.ServiceLocators;
using System.Windows;

namespace GenerateMasterDataForExcel
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        private static IIOExcelFilesProvider ioExcelFilesProvider;
        private static IReadExcelFiles readExcelFilesProvider;

        public void GetInstance()
        {
            ioExcelFilesProvider = ServiceLocatorProvider.GetInstance.ioExcelFileCurrent.Resolve<IIOExcelFilesProvider>();
            readExcelFilesProvider = ServiceLocatorProvider.GetInstance.ReadExcelCurrent.Resolve<IReadExcelFiles>();
        }
    }
}
