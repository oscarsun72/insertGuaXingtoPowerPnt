using CharacterConverttoCharacterPics;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace insertGuaXingtoPowerpnt
{
    class PicsOps
    {
        Image img;
        public PicsOps() { }//預設建構器（default constructor）仍須寫出來，只要有別的建構子
        public PicsOps(Image imgFromFullName)
        {
            img = imgFromFullName;
        }

        internal void resizePNGsSaveAsZip(string dirSource, string dirDest,
            int percent)
        {
            resizePNGsSaveAs(dirSource, dirDest, percent);
            DirFiles.zipFolderFiles(dirDest);
        }
        internal void resizePNGsSaveAs(string dirSource, string dirDest,
            int percent)
        {
            if (!Directory.Exists(dirDest)) Directory.CreateDirectory(dirDest);
            IEnumerable<FileInfo> fileListPNG =
                new DirFiles(dirSource).getPNGs;
            foreach (FileInfo item in fileListPNG)
            {
                img = Image.FromFile(item.FullName);
                saveAsNewPNG(resizeImage(new Size(img.Width * percent / 100,
                    img.Height * percent / 100)),
                     item.FullName.Replace(dirSource, dirDest));
            }
            Process ps = new Process();
            ps.StartInfo.FileName = dirDest;
            ps.Start();
        }

        internal string getPicFullname(string whatsCharacter)
        {
            return new FindFileThruLINQ().getfilefullnameIn古文字(whatsCharacter,
                DirFiles.getPicDir(Form1.PicE) + "\\");
        }

        void saveAsNewPNG(Image img, string destFullName)
        {
            img.Save(destFullName, ImageFormat.Png);
        }

        internal Image resizeImage(
            Size size)
        {//https://www.cnblogs.com/Yesi/p/5952783.html
            if (img == null) return null;
            //获取图片宽度
            int sourceWidth = img.Width;
            //获取图片高度
            int sourceHeight = img.Height;

            float nPercent;
            //计算宽度的缩放比例
            float nPercentW = ((float)size.Width / (float)sourceWidth);
            //计算高度的缩放比例
            float nPercentH = ((float)size.Height / (float)sourceHeight);

            if (nPercentH < nPercentW)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;
            //期望的宽度
            int destWidth = (int)(sourceWidth * nPercent);
            //期望的高度
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((System.Drawing.Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            //绘制图像
            g.DrawImage(img, 0, 0, destWidth, destHeight);
            g.Dispose();
            return (System.Drawing.Image)b;
        }
        
        //將第一次生成的原件字圖調整大小，以冀提昇在Word中插入字圖的性能，並把原件放在560x690的子資料夾中
        internal static void resize字圖ArchiveOrigionIn560x690Folder(int percent,
            string dirDfolder = "文鼎鋼筆行楷B")
        {//包括配置存放的資料夾
            PicsOps po = new PicsOps();            
            //string dirD = @"D:\@@@華語文工具及資料@@@\Macros\古文字\" + dirDfolder;
            string dirD = DirFiles.PicsRootFolder + dirDfolder+"\\";
            string dirS = dirD + "560x690\\";
            if (!Directory.Exists(dirS)) Directory.CreateDirectory(dirS);
            DirectoryInfo dirSdi = new DirectoryInfo(dirS);
            DirectoryInfo dirDdi = new DirectoryInfo(dirD);
            if (dirSdi.EnumerateFiles().Count() == 0 &&//「Count()」是擴充方法要「using System.Linq;」才能用 20210506
                dirDdi.EnumerateFiles().Count() != 0)
            {
                foreach (FileInfo item in dirDdi.GetFiles())
                {
                    item.MoveTo(item.FullName.Replace(dirD,
                        dirS));
                }
            }
            po.resizePNGsSaveAsZip(dirS,dirD, percent);

        }
    }

    #region 圖片
    //特殊的才需要，其他的不需要了。（只要路徑有規則、圖片皆為png，就不必列出了
    //此只是作為判斷時參考爾。卦圖、小篆是路徑；行書是 jpg，故須列出作判斷
    //餘均由listBox1來控制判斷項即可）
    enum picEnum : byte
    {//the zero-based index as listbox 20210411
        卦圖64, 卦形8, 行書, 小篆, Default
    }
    /* , 甲骨文, 金文, 隸書, 文鼎隸書B, 文鼎隸書DB, 文鼎隸書HKM, 文鼎隸書M,
華康行書體, 文鼎行楷L, DFGGyoSho_W7, DFPGyoSho_W7,文鼎魏碑B, 文鼎行楷碑體B, 文鼎鋼筆行楷M, DFPOYoJun_W5,DFPPenJi_W4,

FangSong, Adobe_仿宋_Std_R, 文鼎仿宋B, 文鼎仿宋L,

教育部標準楷書, Adobe_楷体_StdR, KaiTi, 文鼎標準楷體ProM,
文鼎顏楷H, 文鼎顏楷U, 文鼎毛楷B, 文鼎毛楷EB, 文鼎毛楷H,
DFMinchoP_W5,
DFGothicP_W5,
DFGKanTeiRyu_W11, 文鼎古印體B,
文鼎雕刻體B, DFKinBun_W3,
DFGFuun_W7
} */
    #endregion
}
