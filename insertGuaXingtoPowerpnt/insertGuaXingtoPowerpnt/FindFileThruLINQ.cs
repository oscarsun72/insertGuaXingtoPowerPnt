using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace insertGuaXingtoPowerpnt
{
    public class FindFileThruLINQ
    {//如何查詢具有指定屬性或名稱的檔案(c # ) | Microsoft Docs：How to query for files with a specified attribute or name:https://docs.microsoft.com/zh-tw/dotnet/csharp/programming-guide/concepts/linq/how-to-query-for-files-with-a-specified-attribute-or-name
        #region 參考
        /*
        //class FindFileByExtension//LINQ and file directories : https://docs.microsoft.com/zh-tw/dotnet/csharp/programming-guide/concepts/linq/linq-and-file-directories
        // This query will produce the full path for all .txt files  
        // under the specified folder including subfolders.  (包含子資料夾！）
        // It orders the list according to the file name.  
        static void main()
        {
            //string startFolder = @"c:\program files\Microsoft Visual Studio 9.0\";
            string startFolder = @"W:\@@@華語文工具及資料@@@\Macros\古文字\";

            // Take a snapshot of the file system.  
            System.IO.DirectoryInfo dir =
                new System.IO.DirectoryInfo(startFolder);

            // This method assumes that the application has discovery permissions  
            // for all folders under the specified path.  
            IEnumerable<System.IO.FileInfo> fileList =
                dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories);

            //Create the query  
            IEnumerable<System.IO.FileInfo> fileQuery =
                from file in fileList
                where file.Extension == ".png"//".txt"
                orderby file.Name
                select file;

            //Execute the query. This might write out a lot of files!  
            foreach (System.IO.FileInfo fi in fileQuery)
            {
                Console.WriteLine(fi.FullName);
            }

            // Create and execute a new query by using the previous
            // query as a starting point. fileQuery is not
            // executed again until the call to Last()  
            var newestFile =//取得最新的（最後的）檔案
                (from file in fileQuery
                 orderby file.CreationTime
                 select new { file.FullName, file.CreationTime })
                .Last();

            Console.WriteLine("\r\nThe newest .txt file is {0}. " +
                "Creation time: {1}",
                newestFile.FullName, newestFile.CreationTime);

            // Keep the console window open in debug mode.  
            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        } 
        */
        #endregion
        static DirectoryInfo dir; static string StartFolder;
        static IEnumerable<FileInfo> fileList;

        internal static List<string> getfilesfullnameIn古文字(string qText
            , string folderfullNameBackslash, string subFolderName = "")
        {
            string startFolder = folderfullNameBackslash + subFolderName;
            // Take a snapshot of the file system.
            if (StartFolder!=startFolder|| dir==null)
            {
                dir = new DirectoryInfo(startFolder);
                // This method assumes that the application has discovery permissions  
                // for all folders under the specified path. 
                fileList = dir.GetFiles("*.png",SearchOption.AllDirectories);
                StartFolder = startFolder;
            }
            //Create the query  
            IEnumerable<FileInfo> fileQuery =
                from file in fileList
                where file.Name.IndexOf(qText) > -1
                select file;
            List<string> filefullnameList = new List<string>();
            foreach (FileInfo fi in fileQuery)
            {
                filefullnameList.Add(fi.FullName);
            }
            return filefullnameList;
        }

        internal static string getfilefullnameIn古文字(string qText
                                    , string folderfullNameBackslash,
                            string subFolderName = "")
        {
            string startFolder = folderfullNameBackslash + subFolderName;
            if (StartFolder!=startFolder||dir==null)
            {
                dir = new DirectoryInfo(startFolder);
                fileList = dir.GetFiles("*.png",SearchOption.AllDirectories);
                StartFolder = startFolder;
            }
            IEnumerable<FileInfo> fi = (from file in fileList
                                        where file.Name.IndexOf(qText) > -1
                                        select file);
            if (fi.Count()>0)
            {
                 return fi.First().FullName;
            }
            return "";
        }

    }
}
