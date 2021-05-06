using insertGuaXingtoPowerpnt;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.AccessControl;
using System.Windows.Forms;
using powerPnt = Microsoft.Office.Interop.PowerPoint;

namespace CharacterConverttoCharacterPics
{
    public class DirFiles
    {//以後目錄、路徑均要取得最後的反斜線

        DirectoryInfo dir; internal string TopFolder;
        IEnumerable<FileInfo> fileList;
        public DirFiles() { }
        public DirFiles(string topFolder)
        {
            //路徑及存取權的有效性當由呼叫端來檢查！20210506
            TopFolder = topFolder;
            dir = new DirectoryInfo(topFolder);
            fileList = dir.GetFiles("*.*", SearchOption.AllDirectories);
        }

        internal IEnumerable<FileInfo> getAllFiles
        {
            get => fileList ?? null;
            //上式為此式之簡化：get => fileList == null ? null : fileList;
        }
        internal IEnumerable<FileInfo> getPNGs
        {
            get
            {
                return
                    from f in fileList
                    where f.Extension.Equals(".png",//原來 Extension 屬性值包括前綴「.」號20210506
                    System.StringComparison.OrdinalIgnoreCase)
                    select f;
            }
        }

        internal static string getDirRoot
        {//https://www.google.com/search?q=c%23+%E5%8F%96%E5%BE%97%E5%B0%88%E6%A1%88%E8%B7%AF%E5%BE%91&rlz=1C1JRYI_enTW948TW948&oq=%E5%8F%96%E5%BE%97%E5%B0%88%E6%A1%88%E8%B7%AF%E5%BE%91&aqs=chrome.1.69i57j0i5i30l2.7266j0j7&sourceid=chrome&ie=UTF-8
            get =>
                new DirectoryInfo(
                System.AppDomain.CurrentDomain.BaseDirectory)
                .Parent.Parent.Parent.FullName + "\\";
        }

        internal static FileInfo getCjk_basic_IDS_UCS_Basic_txt()
        {
            DirectoryInfo dirRoot = new DirectoryInfo(getDirRoot);
            IEnumerable<FileInfo> fileList = dirRoot.GetFiles
                ("*.txt", SearchOption.AllDirectories);
            IEnumerable<FileInfo> fileQuery =
                from file in fileList
                where file.Name.IndexOf("cjk-basic-IDS-UCS-Basic.txt") > -1
                select file;
            if (fileQuery.Count() > 0)
                return fileQuery.First();
            else
                return null;
        }

        internal static FileInfo getFontOkList_txt()
        {
            //先求方便了，否則一下要兼顧太多檔案20210426
            const string f = @"G:\我的雲端硬碟\programming程式設計開發\fontOkList.txt";
            if (File.Exists(f))
                return new FileInfo(@"G:\我的雲端硬碟\programming程式設計開發\fontOkList.txt");
            else
            {//判斷成功了
                DirectoryInfo dirRoot = new DirectoryInfo(getDirRoot);
                DirectorySecurity directorySecurity = dirRoot.GetAccessControl();
                if (directorySecurity.AreAccessRulesProtected)
                {
                    if (File.Exists(f))
                        return new FileInfo(f);
                    else
                    {
                        MessageBox.Show("此資料夾是無法讀取的！\n\r" +
                            "而「" + f + "」檔案又不存在，故無法執行！", "",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.OpenForms[0].Close();//return null;
                    }
                }
                IEnumerable<FileInfo> fileList = dirRoot.GetFiles
                    ("*.txt", SearchOption.AllDirectories);
                IEnumerable<FileInfo> fileQuery =
                    from file in fileList
                    where file.Name.IndexOf("fontOkList.txt") > -1
                    select file;
                if (fileQuery.Count() > 0)
                    return fileQuery.First();
                else
                    return null;
            }
        }
        internal static string getDir各字型檔相關()
        {
            return getCjk_basic_IDS_UCS_Basic_txt().DirectoryName;
        }

        internal powerPnt.Presentation get字圖母片pptm()
        {
            powerPnt.Application pptApp = App.AppPpt;
            foreach (powerPnt.Presentation ppt in pptApp.Presentations)
            {
                if (ppt.Name == "字圖母片.pptm")
                {
                    return ppt;
                }
            }
            return pptApp.Presentations.Open(
                getDirRoot + "字圖母片.pptm");
        }
        internal static void getPicFolder(string picFolderPath)
        {
            if (Directory.Exists(picFolderPath) == false)
            {
                Directory.CreateDirectory(picFolderPath);
            }
        }

        internal static string PicsRootFolder
        { get => Path華語文工具及資料 + "Macros\\古文字\\"; }

        internal static string Path華語文工具及資料
        {
            get
            {
                string dir;
                DriveInfo[] dis = DriveInfo.GetDrives();
                //List<string> dirs = new List<string> { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };
                foreach (DriveInfo di in dis)//(string item in dirs)
                {
                    //dir = item + ":\\@@@華語文工具及資料@@@" + subFolder;
                    dir = di.Name + "@@@華語文工具及資料@@@\\";
                    if (Directory.Exists(dir))
                    {
                        return dir;
                    }
                }
                return "";
            }
        }
        //without \
        internal static string getPicDir(picEnum pE)//without \
        {
            string subFolder;
            switch (pE)
            {
                case picEnum.卦圖64:
                    subFolder = "Macros\\64卦圖";
                    break;
                case picEnum.卦形8:
                    subFolder = "Macros\\64卦圖\\8卦圖";
                    break;
                case picEnum.小篆:
                    subFolder = "Macros\\古文字\\台大說文小篆字圖";
                    break;
                default://路徑特殊的就析出寫在上面20210410
                    subFolder = "Macros\\古文字\\" +
                       Application.OpenForms[0].Controls["listBox1"].Text;
                    break;
            }
            return Path華語文工具及資料 + subFolder;
        }
        internal static string getFullNameNTUswxz(string dir, string x)
        {//為免ADO存取資料庫失敗而增此
            return new FindFileThruLINQ().getfilefullnameIn古文字(x, dir);
        }

        //將指定資料夾包成同名壓縮檔zip
        internal static void zipFolderFiles(string dir)
        {
            if (Directory.Exists(dir) == false) return;
            DirectoryInfo di = new DirectoryInfo(dir);
            string fZip = di.Parent.FullName + "\\" + di.Name + ".zip";
            if (File.Exists(fZip)) File.Delete(fZip);
            ZipFile.CreateFromDirectory(dir, fZip,
                CompressionLevel.NoCompression, true
                );
        }
    }
}
