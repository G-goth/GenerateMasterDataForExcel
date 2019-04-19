using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;

namespace GenerateMasterDataForExcel.IOFilesProvider
{
    interface IIOFileNamesProvider
    {
        // IOFileName
        List<string> GetSpecifiedExtensionFileFullPath(string fileExtentions);
        string GetSpecifiedExtensionFileName(string fileFullPath);
        List<string> GetSpecifiedExtensionFileNameToList(List<string> fileFullPathList);
    }

    interface IIOExcelFilesProvider
    {
        // IOExcelFiles
        int GetExcelSheetNumberMax(string fileName);
        List<int> GetExcelSheetNumberList(string fileName);
        XLWorkbook GetExcelObject(string fileName);
        string[,] ExtractionExcelData(int sheetNumber, XLWorkbook workbook);
        string[][] ExtractionExcelDataJagged(int sheetNumber, XLWorkbook workBook);
    }

    interface IIOCSVProvider
    {
        void ReadCSV(string fileName);
        List<string[]> ExtractionCSV(StreamReader sr);
    }
}
