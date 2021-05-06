using CharacterConverttoCharacterPics;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace insertGuaXingtoPowerpnt
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            #region 測試根目錄：
            //Console.WriteLine(
            //CharacterConverttoCharacterPics.DirFiles.getDirRoot);
            #endregion
            #region  測試調整圖片大小功能:
            //PicsOps.resize字圖ArchiveOrigionIn560x690Folder(30,"11");
            //warnings.playSound();
            #endregion
        }
    }
}
