using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using powerPnt = Microsoft.Office.Interop.PowerPoint;

namespace CharacterConverttoCharacterPics
{
    public class DirFiles
    {//以後目錄、路徑均要取得最後的反斜線
        internal static string getDirRoot
        {//https://www.google.com/search?q=c%23+%E5%8F%96%E5%BE%97%E5%B0%88%E6%A1%88%E8%B7%AF%E5%BE%91&rlz=1C1JRYI_enTW948TW948&oq=%E5%8F%96%E5%BE%97%E5%B0%88%E6%A1%88%E8%B7%AF%E5%BE%91&aqs=chrome.1.69i57j0i5i30l2.7266j0j7&sourceid=chrome&ie=UTF-8
            get =>
                new DirectoryInfo(
                System.AppDomain.CurrentDomain.BaseDirectory)
                .Parent.Parent.Parent.Parent.FullName+"\\";
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
            DirectoryInfo dirRoot = new DirectoryInfo(getDirRoot);
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
        internal static string getDir各字型檔相關()
        {
            return getCjk_basic_IDS_UCS_Basic_txt().DirectoryName;
        }

        internal powerPnt.Presentation get字圖母片pptm()
        {
            powerPnt.Application pptApp = App.AppPpt;
            foreach (powerPnt.Presentation ppt in pptApp.Presentations)
            {
                if (ppt.Name== "字圖母片.pptm")
                {
                    return ppt;
                }
            }
            return pptApp.Presentations.Open(
                getDirRoot + "字圖母片.pptm");
        }

        internal static void getPicFolder(string picFolderPath)
        {
            if (Directory.Exists(picFolderPath)==false)
            {
                Directory.CreateDirectory(picFolderPath);
            }
        }
    }
}
