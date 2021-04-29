using CharacterConverttoCharacterPics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace insertGuaXingtoPowerpnt
{
    class PicsOps
    {
        internal string getPicFullname(string whatsCharacter)
    {
        return new FindFileThruLINQ().getfilefullnameIn古文字(whatsCharacter,
            DirFiles.getPicDir(Form1.PicE)+"\\");
    }
    }

    #region 圖片
    //特殊的才需要，其他的不需要了。（只要路徑有規則、圖片皆為png，就不必列出了
    //此只是作為判斷時參考爾。卦圖、小篆是路徑；行書是 jpg，故須列出作判斷
    //餘均由listBox1來控制判斷項即可）
    enum picEnum : byte
    {//the zero-based index as listbox 20210411
        卦圖64, 卦形8,行書, 小篆, Default
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
