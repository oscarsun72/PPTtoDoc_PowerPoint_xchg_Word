using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Media;
using System.Windows.Forms;
using powerPnt = Microsoft.Office.Interop.PowerPoint;
using winWord = Microsoft.Office.Interop.Word;

namespace CharacterConverttoCharacterPics
{
    public class fontsPics//負責「字型轉字圖」業務/工作
    {

        //static powerPnt.Application appPpt = App.AppPpt;
        internal winWord.Document
            getFontCharacterset(string fontName)
        {//準備好各字型檔(不含缺字)相關者
            //https://www.google.com/search?q=c%23+%E8%AE%80%E5%8F%96txt&rlz=1C1JRYI_enTW948TW948&sxsrf=ALeKk00EZy0V-LIAiQBz6f5tr6PPx2AI4w%3A1618768409405&ei=GXJ8YKmVGIu9mAW1io6QDw&oq=c%23+%E8%AE%80%E5%8F%96&gs_lcp=Cgdnd3Mtd2l6EAMYADICCAAyAggAMgIIADICCAAyAggAMgIIADICCAAyAggAMgIIADICCAA6BQgAELADOgQIIxAnOgQIABBDOgcIABCxAxBDOgQIABAeOgYIABAIEB46CAgAEAgQChAeULWyUlih0FNg-uhTaAtwAHgBgAGPBogB2gmSAQU3LjYtMZgBAKABAaoBB2d3cy13aXrIAQHAAQE&sclient=gws-wiz            
            App app = new App(); winWord.Application appDoc=app.AppDoc;
            winWord.Document d;
            try {d = appDoc.Documents.Add(""); } catch { app.AppDoc = null; appDoc = app.AppDoc; d = appDoc.Documents.Add(""); }
            //d.ActiveWindow.Visible = true;
            using (StreamReader sr = new StreamReader(DirFiles.getCjk_basic_IDS_UCS_Basic_txt().FullName))
                d.Range().Text = sr.ReadToEnd();//sr在出此行後即會調用Dispose()清除記憶體
            d.Range().Font.NameFarEast = fontName;
            string docName = DirFiles.getDir各字型檔相關() + "\\" +
                fontName + "(不含缺字).docx";
            if (File.Exists(docName))
            {
                DialogResult dr =
                    MessageBox.Show("檔名重複！請檢查，再繼續...是否沿用舊檔？", "",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (dr == DialogResult.Cancel)
                {
                    if (app.DocAppOpenByCode) appDoc.Quit(winWord.
                        WdSaveOptions.wdDoNotSaveChanges);
                    appDoc = null;
                    Application.OpenForms[0].BackColor = Color.White;
                    return null;
                }//沿用舊檔
                d.Close(winWord.WdSaveOptions.wdDoNotSaveChanges);
                return appDoc.Documents.Open(docName);
            }
            else//不存在舊檔
            {
                d.SaveAs2(docName);
                FontsOpsDoc.removeNoFont(d, fontName);
                d.Save();
                return d;
            }
        }


        internal powerPnt.Presentation prepareFontPPT(string fontName,
            float fontsize, string filenameSaveAs = "")
        {
            if (filenameSaveAs == "") filenameSaveAs = fontName + "(不含缺字).pptm";
            //DirFiles df = new DirFiles();
            //powerPnt.Presentation ppt = df.get字圖母片pptm();
            powerPnt.Presentation ppt = DirFiles.get字圖母片pptm();
            ppt.Slides[2].Shapes[1].TextFrame.TextRange.Font.NameFarEast =
                fontName;
            ppt.Slides[2].Shapes[1].TextFrame.TextRange.Font.Size = fontsize;
            ppt.SaveAs(DirFiles.getCjk_basic_IDS_UCS_Basic_txt().DirectoryName + "\\" +
                filenameSaveAs);
            //ppt.Application.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;//沒有這行一樣會現出現來，與Word不同。
            return ppt;
        }
        int charPicCounter = 0;
        public int CharPicCounter { get => charPicCounter; }

        List<powerPnt.Presentation> addCharsSlides(winWord.Document fontCharacterset,
            powerPnt.Presentation ppt, string pptFullname
            , int howManyCharsPPT = 5000)
        {
            //不用using，是怕離開using後 btn 被 dispose 而連帶的Form1上的button1也會消失。若真如此，可見下行乃指針或參考指令，非複製原物件為副本之指令也20210427
            //using (Button btn = (Button)Application.OpenForms[0].Controls["button1"])
            Button btn = (Button)Application.OpenForms[0].Controls["button1"];
            {//https://bit.ly/3xbEGfH
             //https://oscarsun72.blogspot.com/2021/04/reprintedusing-c.html
                btn.Parent.Refresh();
                int charPicCounterOK = 0;
                string Xpicsok = ""; powerPnt.SlideRange sld;
                List<powerPnt.Presentation> returnPPTs = new List<powerPnt.Presentation>();
                string fontname = ppt.Slides[2].Shapes[1].TextFrame.TextRange.Font.NameFarEast;
                float fontsize = ppt.Slides[2].Shapes[1].TextFrame.TextRange.Font.Size;

            //if f = "" Then Exit Sub
            //If InStr(f, "?") Then
            //    MsgBox "路徑中有亂碼，請檢查，或修改程式"
            //    Exit Sub
            //End If
            //powerPnt.Presentation ppt = App.AppPpt.ActivePresentation;
            //foreach (winWord.Range an in fontCharacterset.Range().Characters)
            crashRedo:
                winWord.Range rng = fontCharacterset.Range(fontCharacterset.Characters[charPicCounterOK + 1]
                    .Start, fontCharacterset.Range().End);
                foreach (winWord.Range an in rng.Characters)
                {
                    //if (InStr(Chr(13) & Chr(7) & Chr(9) & Chr(8) & Chr(10) _
                    //        , a) = 0 Then
                    //https://dotblogs.com.tw/mis2000lab/2013/11/06/126917
                    //https://bbs.csdn.net/topics/90012123
                    if ("\r\n\t".IndexOf(an.Text) == -1)
                    {
                        if (Xpicsok.IndexOf(an.Text) == -1)
                        {
                            Xpicsok += an.Text;
                            try
                            {
                                sld = ppt.Slides[ppt.Slides.Count].Duplicate();
                                sld.Shapes[1].TextFrame.TextRange.Text = an.Text;
                            }
                            catch (System.Exception)
                            {
                                Application.DoEvents();//讓系統處理完pptApp當掉的程序
                                App app = new App();
                                app.AppPpt = null;
                                ppt = app.AppPpt.Presentations.Open(pptFullname);
                                if (ppt.Slides.Count == 2)
                                {
                                    charPicCounter = 0; Xpicsok = "";
                                    goto crashRedo;
                                }
                                sld = ppt.Slides.Range(ppt.Slides.Count);
                                if (sld.Shapes[1].TextFrame.TextRange.Text != an.Text)
                                {
                                    sld = ppt.Slides[ppt.Slides.Count].Duplicate();
                                    sld.Shapes[1].TextFrame.TextRange.Text = an.Text;
                                }
                            }
                            charPicCounter++;
                            if (charPicCounter % 500 == 0)
                            { //https://docs.microsoft.com/zh-tw/dotnet/csharp/language-reference/operators/arithmetic-operators
                              //fontCharacterset.Application.StatusBar = "已處理" + i + "個字";
                                btn.Text = "已處理" + charPicCounter + "個字";
                                btn.Parent.Refresh();//光是按鈕refresh則按鈕會消失不見
                            }

                            if (charPicCounter % howManyCharsPPT == 0)
                            {//預設5000字一檔，以提升效率
                                ppt.Slides[2].Delete();//第2張是作樣本Duplicate()之依據，故用完即丟
                                warnings.playBeep();
                                ppt.Save();//準備分檔、準備下一檔
                                charPicCounterOK = charPicCounter;
                                returnPPTs.Add(ppt);
                                //ppt.Close(); //存在List中後續要用，不能關掉
                                //ppt.Windows[1].ViewType=好像也沒有隱藏功能，只好秀著
                                ppt = prepareFontPPT(fontname,
                                    fontsize, fontname + "(不含缺字)" +
                                    (charPicCounter + 1).ToString() + "~.pptm");
                            }
                        }
                    }
                }
                ppt.Slides[2].Delete();//第2張是作樣本Duplicate()之依據，故用完即丟
                //warnings.playBeep();
                SystemSounds.Asterisk.Play();
                ppt.Save();//以免當掉
                returnPPTs.Add(ppt);
                return returnPPTs;
            }
        }

        internal void addCharsSlidesExportPng(winWord.Document fontCharacterset,
            powerPnt.Presentation ppt, string exportDir, int howManyCharsPPT = 5000)
        {
            List<powerPnt.Presentation> ppts = addCharsSlides(
                fontCharacterset, ppt, ppt.FullName, howManyCharsPPT);
            int fontCharactersetCount = fontCharacterset.Range().Characters.Count;
            if (ppts.Count > 0)
            {
                int pptSlidesCtr = 0;
                foreach (powerPnt.Presentation item in ppts)
                {
                    pptSlidesCtr += (item.Slides.Count - 1);//每檔皆會有第一張空白投影片
                    exportPng(item, exportDir, ppts.Count);//直接轉成字圖 20210419
                }

                if (++pptSlidesCtr != fontCharactersetCount)//投影片總數加1則與Word檔文字總數會包括最後一個分段符號吻合
                {
                    warnings.playSound();
                    //btn.Parent.Refresh();
                    Button btn = (Button)Application.OpenForms[0].Controls["button1"];
                    btn.Text = "字數有所不同，請留意！";
                    btn.Parent.BackColor = Color.BurlyWood;//MessageBox.Show("字數有所不同，請留意！", "注意：",
                    //MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                DirFiles.openFolder(exportDir);
            }
            else//沒有分割（分別存）ppt檔的話
            {
                if (fontCharactersetCount != ppt.Slides.Count) //若不分段，則Word後有一個chr(13)與此母片前多一張，正好抵消
                {
                    warnings.playSound();
                    //btn.Parent.Refresh();
                    Button btn = (Button)Application.OpenForms[0].Controls["button1"];
                    btn.Text = "字數有所不同，請留意！";
                    btn.Parent.BackColor = Color.BurlyWood;
                    //MessageBox.Show("字數有所不同，請留意！", "注意：",
                    //      MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                exportPng(ppt, exportDir);//字數一致，就直接轉成字圖 20210408
                ppt.Close();//還是會自己關掉，要比較久而已 感恩感恩　南無阿彌陀佛 20210419
                //ppt = null;//close()了就不用此行設為null了
            }
            App app = new App();
            if (app.PptAppOpenByCode)//若是PowerPoint是由程式開啟則關閉。會經過一段時間，或本應用程式結束後一段時間，才會關閉20210419
            {
                ppt.Application.Quit(); 
            }//然剛才發現，只要本應用程式關閉，則會瞬間跟著關掉20210419 20:19
            fontCharacterset.Close();
            if (app.DocAppOpenByCode) fontCharacterset.Application.Quit();            
            //warnings.playSound();// (ppt.Slides.Count);
        }

        void exportPng(powerPnt.Presentation ppt, string picDir
            , int pptsCount = 0)
        {
            string w; DirFiles.getPicFolder(picDir);
            if (picDir == "") return;
            if (picDir.Substring(picDir.Length - 1, 1) != "\\") picDir += "\\";
            if (!Directory.Exists(picDir)) return;
            foreach (powerPnt.Slide sld in ppt.Slides)
            {
                w = sld.Shapes[1].TextFrame.TextRange.Text.Trim();
                if (w != "")//第一張投影片即是空白
                    sld.Export(picDir + w + ".png", "PNG");//https://docs.microsoft.com/zh-tw/office/vba/api/powerpoint.slide.export
            }
            if (pptsCount == 0)
            {
                DirFiles.openFolder(picDir);
            }
        }


    }

}

