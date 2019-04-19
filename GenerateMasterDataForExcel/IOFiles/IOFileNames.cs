using GenerateMasterDataForExcel.IOFilesProvider;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GenerateMasterDataForExcel
{
    class IOFileName : IIOFileNamesProvider
    {
        /// <summary>
        /// 指定された拡張子のファイルのフルパスを検索して取得する。
        /// </summary>
        /// <param name="fileExtentions">拡張子名(xlsx、docxなど)</param>
        /// <returns>返り値はstringのList なかった場合は空のListが返り値となる</returns>
        public List<string> GetSpecifiedExtensionFileFullPath(string fileExtentions)
        {
            string dir = System.Environment.CurrentDirectory;
            System.Environment.CurrentDirectory = dir;
            string[] files = System.IO.Directory.GetFiles(@"" + dir, "*", System.IO.SearchOption.TopDirectoryOnly);
            List<string> fileFullPath = new List<string>();

            int fileExtentionsCount;
            string trimStr;
            foreach (string item in files)
            {
                // 3文字拡張子ファイルのフルパスのみを入れる
                fileExtentionsCount = item.Count() - 3;
                trimStr = item.Substring(fileExtentionsCount);
                if (trimStr == fileExtentions)
                {
                    fileFullPath.Add(item);
                }
                // 4文字拡張子ファイルのフルパスのみを入れる
                fileExtentionsCount = item.Count() - 4;
                trimStr = item.Substring(fileExtentionsCount);
                if (trimStr == fileExtentions)
                {
                    fileFullPath.Add(item);
                }
            }
            return fileFullPath;
        }

        /// <summary>
        /// ファイルのフルパスから拡張子付きファイル名を取得
        /// </summary>
        /// <param name="fileFullPath">ファイルのフルパス</param>
        /// <returns>拡張子付きファイル名</returns>
        public string GetSpecifiedExtensionFileName(string fileFullPath)
        {
            string dir = System.Environment.CurrentDirectory;
            Uri u1 = new Uri(dir);
            Uri u2 = new Uri(fileFullPath);
            Uri relativeUri = u1.MakeRelativeUri(u2);
            string relativePath = relativeUri.ToString();
            int separateFilename = relativePath.IndexOf("/") + 1;

            return relativePath.Substring(separateFilename);
        }

        /// <summary>
        /// ファイルのフルパスから拡張子付きファイル名をListで取得
        /// </summary>
        /// <param name="fileFullPathList">ファイル名のList</param>
        /// <returns>ファイル名の相対パス</returns>
        public List<string> GetSpecifiedExtensionFileNameToList(List<string> fileFullPathList)
        {
            // 得られた絶対パスをUriコンストラクタに値を入れてList初期化
            Uri u1 = new Uri(System.Environment.CurrentDirectory);
            List<Uri> u2List = fileFullPathList.Select(filePath => new Uri(filePath)).ToList();

            // 得られた絶対パスのリストごとに相対パスを作ってリストに入れる
            List<string> relativePathList = new List<string>();
            return relativePathList = u2List.Select(u2 => u1.MakeRelativeUri(u2).ToString().Substring(u1.MakeRelativeUri(u2).ToString().IndexOf("/") + 1)).ToList();
        }
    }
}
