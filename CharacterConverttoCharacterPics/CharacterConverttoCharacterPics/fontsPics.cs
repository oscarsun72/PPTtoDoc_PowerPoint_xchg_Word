using System.Drawing;
using System.IO;
using System.Windows.Forms;
using powerPnt = Microsoft.Office.Interop.PowerPoint;
using winWord = Microsoft.Office.Interop.Word;

namespace CharacterConverttoCharacterPics
{
    public class fontsPics//負責「字型轉字圖」業務/工作
    {
        static winWord.Application appDoc = App.AppDoc;
        //static powerPnt.Application appPpt = App.AppPpt;
        internal static winWord.Document getFontCharacterset(string fontName)
        {//準備好各字型檔(不含缺字)相關者
            //https://www.google.com/search?q=c%23+%E8%AE%80%E5%8F%96txt&rlz=1C1JRYI_enTW948TW948&sxsrf=ALeKk00EZy0V-LIAiQBz6f5tr6PPx2AI4w%3A1618768409405&ei=GXJ8YKmVGIu9mAW1io6QDw&oq=c%23+%E8%AE%80%E5%8F%96&gs_lcp=Cgdnd3Mtd2l6EAMYADICCAAyAggAMgIIADICCAAyAggAMgIIADICCAAyAggAMgIIADICCAA6BQgAELADOgQIIxAnOgQIABBDOgcIABCxAxBDOgQIABAeOgYIABAIEB46CAgAEAgQChAeULWyUlih0FNg-uhTaAtwAHgBgAGPBogB2gmSAQU3LjYtMZgBAKABAaoBB2d3cy13aXrIAQHAAQE&sclient=gws-wiz            
            winWord.Document d = appDoc.Documents.Add("");
            //d.ActiveWindow.Visible = true;
            d.Range().Text = new StreamReader(DirFiles.getCjk_basic_IDS_UCS_Basic_txt().FullName).ReadToEnd();
            d.Range().Font.NameFarEast = fontName;
            string docName = DirFiles.getDir各字型檔相關() + "\\" + fontName + "(不含缺字).docm";
            if (File.Exists(docName))
            {
                MessageBox.Show("檔名重複！請檢查，再繼續");
                if (App.DocAppOpenByCode) appDoc.Quit(winWord.
                    WdSaveOptions.wdDoNotSaveChanges);
                appDoc = null;
                Application.OpenForms[0].BackColor = Color.White;
                return null;
            }
            d.SaveAs2(docName);
            FontsOpsDoc.removeNoFont(d, fontName);
            d.Save();
            return d;
        }


        internal static powerPnt.Presentation prepareFontPPT(string fontName, float
            fontsize)
        {
            powerPnt.Presentation ppt = DirFiles.get字圖母片pptm();
            ppt.Slides[2].Shapes[1].TextFrame.TextRange.Font.NameFarEast =
                fontName;
            ppt.Slides[2].Shapes[1].TextFrame.TextRange.Font.Size = fontsize;
            ppt.SaveAs(DirFiles.getCjk_basic_IDS_UCS_Basic_txt().DirectoryName +"\\"+
                fontName + "(不含缺字).pptm");
            //ppt.Application.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;//沒有這行一樣會現出現來，與Word不同。
            return ppt;
        }


        static void addCharsSlides(winWord.Document fontCharacterset,
            powerPnt.Presentation ppt)
        {
            string X = ""; int i = 0; powerPnt.SlideRange sld;
            //if f = "" Then Exit Sub
            //If InStr(f, "?") Then
            //    MsgBox "路徑中有亂碼，請檢查，或修改程式"
            //    Exit Sub
            //End If
            //powerPnt.Presentation ppt = App.AppPpt.ActivePresentation;
            foreach (winWord.Range an in fontCharacterset.Range().Characters)
            {
                //if (InStr(Chr(13) & Chr(7) & Chr(9) & Chr(8) & Chr(10) _
                //        , a) = 0 Then
                //https://dotblogs.com.tw/mis2000lab/2013/11/06/126917
                //https://bbs.csdn.net/topics/90012123
                if ("\r\n\t".IndexOf(an.Text) == -1)
                {
                    if (X.IndexOf(an.Text) == -1)
                    {
                        X += an.Text;
                        sld = ppt.Slides[ppt.Slides.Count].Duplicate();
                        sld.Shapes[1].TextFrame.TextRange.Text = an.Text;
                        i++;
                        if (i % 500 == 0)
                        { //https://docs.microsoft.com/zh-tw/dotnet/csharp/language-reference/operators/arithmetic-operators
                            fontCharacterset.Application.StatusBar = "已處理" + i + "個字";
                        }
                    }
                }
            }
            ppt.Slides[2].Delete();//第2張是作樣本Duplicate()之依據，故用完即丟
            warnings.playBeep();
            ppt.Save();//以免當掉            
        }

        internal static void addCharsSlidesExportPng(winWord.Document fontCharacterset,
            powerPnt.Presentation ppt, string exportDir)
        {
            addCharsSlides(fontCharacterset, ppt);
            if (fontCharacterset.Range().Characters.Count ==
                            ppt.Slides.Count) //若不分段，則Word後有一個chr(13)與此母片前多一張，正好抵消
                exportPng(ppt, exportDir);//字數一致，就直接轉成字圖 20210408
            else
                MessageBox.Show("字數不同，請檢查");
            if (App.PptAppOpenByCode) ppt.Application.Quit();
            ppt.Close();//還是會自己關掉，要比較久而已 感恩感恩　南無阿彌陀佛 20210419
            ppt = null;
            App.AppPpt.Quit();
            App.AppPpt = null;
            fontCharacterset.Close();
            if(App.DocAppOpenByCode) appDoc.Quit();
            appDoc = null;
            warnings.playSound();// (ppt.Slides.Count);
        }

        static void exportPng(powerPnt.Presentation ppt, string picDir)
        {
            string w ;DirFiles.getPicFolder(picDir);
            if (picDir == "") return;
            if (picDir.Substring(picDir.Length - 1, 1) != "\\") picDir += "\\";
            if (!Directory.Exists(picDir)) return;
            foreach (powerPnt.Slide sld in ppt.Slides)
            {
                w = sld.Shapes[1].TextFrame.TextRange.Text.Trim();
                if (w != "")
                    sld.Export(picDir + w + ".png", "PNG");//https://docs.microsoft.com/zh-tw/office/vba/api/powerpoint.slide.export
            }
            //Process.Start(picDir);//Shell "explorer " & pth, vbMaximizedFocus;
            //https://happyduck1020.pixnet.net/blog/post/34382453-c%23-%E9%96%8B%E5%95%9F%E8%B3%87%E6%96%99%E5%A4%BE
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = picDir;
            prc.Start();
            Application.DoEvents();
            warnings.playBeep();
        }

    }

}

