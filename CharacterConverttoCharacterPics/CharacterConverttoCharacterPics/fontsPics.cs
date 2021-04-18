using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using powerPnt = Microsoft.Office.Interop.PowerPoint;
using winWord = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Drawing;

namespace CharacterConverttoCharacterPics
{
    public class fontsPics//負責「字型轉字圖」業務/工作
    {
        static winWord.Application appDoc = App.AppDoc;
        static powerPnt.Application appPpt = App.AppPpt;
        internal static winWord.Document getFontCharacterset(string fontName)
        {//準備好各字型檔(不含缺字)相關者
            //https://www.google.com/search?q=c%23+%E8%AE%80%E5%8F%96txt&rlz=1C1JRYI_enTW948TW948&sxsrf=ALeKk00EZy0V-LIAiQBz6f5tr6PPx2AI4w%3A1618768409405&ei=GXJ8YKmVGIu9mAW1io6QDw&oq=c%23+%E8%AE%80%E5%8F%96&gs_lcp=Cgdnd3Mtd2l6EAMYADICCAAyAggAMgIIADICCAAyAggAMgIIADICCAAyAggAMgIIADICCAA6BQgAELADOgQIIxAnOgQIABBDOgcIABCxAxBDOgQIABAeOgYIABAIEB46CAgAEAgQChAeULWyUlih0FNg-uhTaAtwAHgBgAGPBogB2gmSAQU3LjYtMZgBAKABAaoBB2d3cy13aXrIAQHAAQE&sclient=gws-wiz            
            winWord.Document d=appDoc.Documents.Add("");
            d.ActiveWindow.Visible = true;
            d.Range().Text = new StreamReader(DirFiles.getCjk_basic_IDS_UCS_Basic_txt().FullName).ReadToEnd(); 
            d.Range().Font.NameFarEast = fontName;
            string docName = DirFiles.getDir各字型檔相關() + "\\" + fontName + "(不含缺字).docm";
            if (File.Exists(docName))
            {
                MessageBox.Show("檔名重複！請檢查，再繼續");
                Application.OpenForms[0].BackColor = Color.White;
                return null;
            }
            d.SaveAs2(docName);
            FontsOpsDoc.removeNoFont(d, fontName);
            d.Save();
            return d;
        }


        
    }
}
