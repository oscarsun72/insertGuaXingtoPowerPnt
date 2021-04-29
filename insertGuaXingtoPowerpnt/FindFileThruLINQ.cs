using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Windows.Forms;

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
        DirectoryInfo dir; internal string StartFolder;
        IEnumerable<FileInfo> fileList;
        public FindFileThruLINQ() { }
        public FindFileThruLINQ(string startFolder)
        {
            StartFolder = startFolder;
            dir = new DirectoryInfo(startFolder);
            fileList = dir.GetFiles("*.*", SearchOption.AllDirectories);
        }

        internal string TopFolder { get => StartFolder; }
        internal IEnumerable<FileInfo> FileList { get => fileList; }

        internal IEnumerable<FileInfo> findFiles(string fileNameInclude,
            string extensionNameNodot="")
        {
            if (dir == null) return null;
            //https://bit.ly/3xtAAzX
            //https://bit.ly/3vwBfOX
            DirectorySecurity directorySecurity = dir.GetAccessControl();
            if (directorySecurity.AreAccessRulesProtected)
            {
                MessageBox.Show("此資料夾是無法讀取的！","",
                    MessageBoxButtons.OK,MessageBoxIcon.Error);return null;
            }
            fileList = dir.GetFiles("."+ extensionNameNodot, SearchOption.AllDirectories);
            IEnumerable<FileInfo> filequeryResult =
                from file in fileList
                where file.Name.IndexOf(fileNameInclude) > -1
                select file;
            return filequeryResult;
        }
        internal List<string> getfilesfullnameIn古文字(string qText
            , string folderfullNameBackslash, string subFolderName = "")
        {
            string startFolder = folderfullNameBackslash + subFolderName;
            // Take a snapshot of the file system.
            dir = new DirectoryInfo(startFolder);
            // This method assumes that the application has discovery permissions  
            // for all folders under the specified path. 
            string ext = "png";
            if (folderfullNameBackslash.IndexOf("行書") > -1)
            {
                ext = "jpg";
            }
            fileList = dir.GetFiles("*." + ext, SearchOption.AllDirectories);
            StartFolder = startFolder;
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

        internal string getfilefullnameIn古文字(string qText
                                    , string folderfullNameBackslash,
                            string subFolderName = "")
        {
            string startFolder = folderfullNameBackslash + subFolderName;
            dir = new DirectoryInfo(startFolder);
            string ext = "png";
            if (startFolder.IndexOf("行書") > -1) ext = "jpg";
            fileList = dir.GetFiles("*."+ext, SearchOption.AllDirectories);
            StartFolder = startFolder;
            IEnumerable<FileInfo> fi = (from file in fileList
                                        where file.Name.IndexOf(qText) > -1
                                        select file);
            if (fi.Count() > 0)
            {
                return fi.First().FullName;
            }
            return "";
        }

    }
}
