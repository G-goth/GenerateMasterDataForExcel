using ClosedXML.Excel;

namespace GenerateMasterDataForExcel.ReadFiles
{
    public interface IReadExcelFiles
    {
        void GenerateStageAndWave(XLWorkbook workbook, string fileName);
    }
}
