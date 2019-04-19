using ClosedXML.Excel;
using GenerateMasterDataForExcel.IOFilesProvider;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GenerateMasterDataForExcel.IOFiles.IOExcel
{
    class IOExcelFiles : IIOExcelFilesProvider
    {
        /// <summary>
        /// 取得したExcelファイルの最大シート数
        /// </summary>
        /// <param name="fileName">Excelファイル名</param>
        /// <returns>最大シート数</returns>
        public int GetExcelSheetNumberMax(string fileName)
        {
            try
            {
                using (XLWorkbook workBook = new XLWorkbook(@"" + fileName))
                {
                    return workBook.Worksheets.Count;
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e);
                return 0;
            }
        }

        /// <summary>
        /// 取得したExcelファイルのシート数を連番でListに登録して返す
        /// </summary>
        /// <param name="fileName">Excelファイル名</param>
        /// <returns>連番の入ったList</returns>
        public List<int> GetExcelSheetNumberList(string fileName)
        {
            List<int> serialNumber = new List<int>();
            using (XLWorkbook workBook = new XLWorkbook(@"" + fileName))
            {
                try
                {
                    serialNumber.AddRange(Enumerable.Repeat(1, workBook.Worksheets.Count).ToList());
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e);
                    return serialNumber;
                }
            }
            return serialNumber;
        }

        /// <summary>
        /// ClosedXMLでExcelワークブックを取得する
        /// </summary>
        /// <param name="fileName">ファイル名(拡張子まで)</param>
        /// <returns>XLWorkbook型変数</returns>
        public XLWorkbook GetExcelObject(string fileName)
        {
            using (XLWorkbook workBook = new XLWorkbook(@"" + fileName))
            {
                try
                {
                    return workBook;
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e);
                    return workBook;
                }
            }
        }

        /// <summary>
        /// Excelのデータをセルごとに読み込んで2次元配列に代入する
        /// </summary>
        /// <param name="sheetNumber">シート番号(1から)</param>
        /// <param name="workBook">XLWorkbook変数</param>
        /// <returns>使用しているExcelのセルのデータをstringの2次元配列で返す</returns>
        public string[,] ExtractionExcelData(int sheetNumber, XLWorkbook workBook)
        {
            IXLWorksheet workSheet = workBook.Worksheet(sheetNumber);
            (int column, int row) xlCellAddress;
            xlCellAddress.column = workSheet.LastColumnUsed().ColumnNumber();
            xlCellAddress.row = workSheet.LastRowUsed().RowNumber();
            string[,] xlStrDataArray = new string[xlCellAddress.column, xlCellAddress.row];
            return xlStrDataArray;
        }

        /// <summary>
        /// Excelのデータをセルごとに読み込んでジャグ配列に代入する
        /// </summary>
        /// <param name="sheetNumber">シート番号(1から)</param>
        /// <param name="workBook">XLWorkbook変数</param>
        /// <returns>使用しているExcelのセルのデータをstringのジャグ配列で返す</returns>
        public string[][] ExtractionExcelDataJagged(int sheetNumber, XLWorkbook workBook)
        {
            IXLWorksheet workSheet = workBook.Worksheet(sheetNumber);
            (int column, int row) xlCellAddress;
            xlCellAddress.column = workSheet.LastColumnUsed().ColumnNumber();
            xlCellAddress.row = workSheet.LastRowUsed().RowNumber();

            // ジャグ配列にExcelのセルのデータを入れる
            string[][] xlStrDataJaggedArray = new string[xlCellAddress.row][];
            xlStrDataJaggedArray = Enumerable.Range(0, xlCellAddress.column)
                .Select(row => (new string[xlCellAddress.column]).Select(str => workSheet.Cell(1, 1).Value.ToString()).ToArray()).ToArray();

            string[][] xlStrDataArray2 = new string[xlCellAddress.row][];
            return xlStrDataJaggedArray;
        }
    }
}
