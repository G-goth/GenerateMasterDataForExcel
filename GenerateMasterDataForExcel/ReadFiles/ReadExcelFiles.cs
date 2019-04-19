using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GenerateMasterDataForExcel.ReadFiles
{
    public class ReadExcelFiles : IReadExcelFiles
    {
        public void GenerateStageAndWave(XLWorkbook workbook, string fileName)
        {
            IXLWorksheet worksheet1 = workbook.Worksheet(1);
            IXLWorksheet worksheet2 = workbook.Worksheet(2);
            int row1 = worksheet1.LastRowUsed().RowNumber() - 1;
            List<int> stageIDs = new List<int>();
            List<int> stageWave = new List<int>();
            List<int> totalEnemies = new List<int>();

            // A列とB列のデータを取得
            stageIDs = Enumerable.Range(2, row1).Select(index => worksheet1.Cell(index, 1).GetValue<int>()).ToList();
            stageWave = Enumerable.Range(2, row1).Select(index => worksheet1.Cell(index, 2).GetValue<int>()).ToList();
            totalEnemies = Enumerable.Range(2, row1).Select(index => worksheet1.Cell(index, 3).GetValue<int>()).ToList();

            // ステージIDとウェーブ数の組み合わせをList<(int, int)>に入れ込む
            List<(int id, int wave)> stageIDsTuple = new List<(int, int)>();
            for (int ids = 0; ids < stageIDs.Count; ++ids)
            {
                double allocationAns = Math.Floor(totalEnemies[ids] / (double)stageWave[ids]);
                for (int wave = 0; wave < stageWave[ids]; ++wave)
                {
                    for (int allocation = 0; allocation <= allocationAns; ++allocation)
                    {
                        if (stageIDsTuple.Count(tuples => tuples.id == stageIDs[ids]) < totalEnemies[ids])
                        {
                            stageIDsTuple.Add((stageIDs[ids], (wave + 1)));
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }

            // ExcelファイルにステージIDとウェーブ数を書き出す
            for (int i = 0; i < stageIDsTuple.Count; ++i)
            {
                worksheet2.Cell((i + 1), 1).Value = stageIDsTuple[i].id;
                worksheet2.Cell((i + 1), 2).Value = stageIDsTuple[i].wave;
            }
            workbook.SaveAs(fileName);
            Console.WriteLine("Done!!");
        }
    }
}
