using Microsoft.Office.Core;
using System;
using System.IO;
using System.Media;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WinWord = Microsoft.Office.Interop.Word;
namespace Doc_PPt
{
    class TableDocOps
    {
        WinWord.Application wdApp;
        public TableDocOps()
        {
            try
            {
                wdApp = (WinWord.Application)Marshal.GetActiveObject("Word.Application");
            }
            catch (Exception)
            {
                wdApp = new WinWord.Application();
                //throw;
            }


        }


        internal void splitTableByEachRowTitleed字源圖片()
        {
            WinWord.Document d = wdApp.ActiveDocument;
            if (d.Name != "＃字源圖片 （象形）.docx")
            {
                foreach (WinWord.Document item in wdApp.Documents)
                {
                    if (item.Name == "＃字源圖片 （象形）.docx")
                    {
                        item.Activate(); d = item; break;
                    }
                }
                string dFullname = "";
                if (wdApp.ActiveDocument.Name != "＃字源圖片 （象形）.docx")
                {
                    TextBox textBox1 = (TextBox)Application.OpenForms[0].Controls[0];
                    if (textBox1.Text.IndexOf("字源圖片") > 1)
                    {
                        dFullname = textBox1.Text.Replace(@"file:///", "").Replace("%20", " ");
                    }
                    else
                    {
                        MessageBox.Show("請在textBox1文字方塊輸入「字源圖片」的全檔名"); return;
                    }
                }

                if (dFullname == "" || !File.Exists(dFullname))
                {
                    MessageBox.Show("請在textBox1文字方塊輸入「字源圖片」的「正確的」全檔名"); return;
                }
                d = wdApp.Documents.Open(dFullname);
            }
            d.Tables[1].Cell(3, 1).Range.Characters[1].Select();
            WinWord.Selection Selection = d.ActiveWindow.Selection;
            Selection.Collapse(WinWord.WdCollapseDirection.wdCollapseStart);
            int r, s, s1; WinWord.Cell cel; WinWord.Range rng;
            WinWord.InlineShape inlsp; WinWord.Table tb;
            //List<WinWord.InlineShape> inlsps = new List<WinWord.InlineShape>();
            WinWord.Row rw;
            r = 1;
            rng = Selection.Range;
            wdApp.ScreenUpdating = false;
            d.Tables[1].Rows.Add(); d.Tables[1].Rows.Add();//最後會留下一個表格再予刪除
            while (Selection.Information[WinWord.WdInformation.wdWithInTable])
            {
                Selection.SplitTable();
                /* 表格置中都無效
                 * Selection.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                */
                rw = Selection.Document.Tables[1].Rows[1];
                rw.Range.Copy();
                Selection.Document.Tables[Selection.Document.Tables.Count].Range.Characters[1].Select();
                Selection.Collapse(WinWord.WdCollapseDirection.wdCollapseStart);
                Selection.Paste();
                Selection.Document.Tables[Selection.Document.Tables.Count].Range.Characters[1].Select();
                Selection.Collapse(WinWord.WdCollapseDirection.wdCollapseStart);
                Selection.MoveLeft();
                if (Selection.Document.Tables[r].Rows.Count == 1)
                    cel = Selection.Document.Tables[r].Cell(1, 8);
                else
                    cel = Selection.Document.Tables[r].Cell(2, 8);
                if (cel.Range.InlineShapes.Count > 0) {; }
                else
                {
                    if (Selection.Document.Tables[r].Rows.Count > 1)
                        cel = Selection.Document.Tables[r].Cell(2, 8);
                }
                s = Selection.Start;
                rng.SetRange(s, s);
                if (cel.Range.InlineShapes.Count > 0)
                {
                    cel.Range.InlineShapes[1].Select();
                    Selection.Cut();
                    s1 = Selection.Start;
                    if (s1 > s)
                    {
                        while (rng.Information[WinWord.WdInformation.wdWithInTable])
                        {
                            s1--;
                            rng.SetRange(s1, s1);
                        }
                    }
                    else if (s1 < s)
                    {
                        while (rng.Information[WinWord.WdInformation.wdWithInTable])
                        {
                            s1++;
                            rng.SetRange(s1, s1);
                        }

                    }
                    rng.Select();
                    Selection.Paste();//圖片貼到定位

                    //foreach (WinWord.InlineShape insp in Selection.Previous().InlineShapes)
                    //{
                    //    inlsps.Add(insp);
                    //    insp.Height += 181;
                    //    insp.Width += 181;
                    //}

                    if (Selection.Previous().InlineShapes.Count > 0)
                    {
                        inlsp = Selection.Previous().InlineShapes[1];
                        inlsp.LockAspectRatio = MsoTriState.msoTrue;
                        inlsp.Height = 200;
                    }//調整圖片大小
                    else
                    {
                        Selection.MoveRight(WinWord.WdUnits.wdCharacter, 1, WinWord.WdMovementType.wdExtend);
                        inlsp = Selection.InlineShapes[1];
                        inlsp.Height += 181;// = Selection.InlineShapes[1].Height + 181;
                        inlsp.Width += 181;//= Selection.InlineShapes[1].Height + 181;
                    }
                    //圖片置中
                    //Selection.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                    //插入表格，將圖片置入
                    tb = Selection.Tables.Add(Selection.Range, 1, 2);
                    tb.Borders.InsideLineStyle = WinWord.WdLineStyle.wdLineStyleSingle;
                    tb.Borders.OutsideLineStyle = WinWord.WdLineStyle.wdLineStyleDouble;
                    //tb.AutoFitBehavior(WinWord.WdAutoFitBehavior.wdAutoFitFixed);
                    //tb.Cell(1, 1).SetWidth(Selection.Document.Tables[r].PreferredWidth -
                    //  tb.Cell(1, 2).Width, WinWord.WdRulerStyle.wdAdjustNone);
                    //表格置中
                    //此無效：tb.Range.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                    //這才有效：
                    tb.Rows.Alignment=WinWord.WdRowAlignment.wdAlignRowCenter;
                    inlsp.Select(); Selection.Cut();//剪下圖片貼入表格
                    tb.Cell(1, 2).Range.Characters[1].Select();
                    Selection.Paste();
                    tb.PreferredWidthType = WinWord.WdPreferredWidthType.wdPreferredWidthPoints;
                    tb.PreferredWidth = (float)549.6378;//Selection.Document.Tables[r].PreferredWidth;
                    tb.Range.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                    Selection.MoveDown();
                    //Selection.Collapse(WinWord.WdCollapseDirection.wdCollapseEnd);
                    //與下一分割出來的表格空2行（段）
                    Selection.InsertParagraphAfter(); Selection.InsertParagraphAfter();
                }
                Selection.Document.Tables[r].Columns[8].Cells.Delete();
                r += 2;//r++; r++; //前面Tables.Add多插一表格，計數要再加1
                if (r > d.Tables.Count) break;
                if (Selection.Document.Tables[r].Rows.Count > 3)//結束時，尚須修改。目前可以權且加幾空白列在最後一列後
                    Selection.Document.Tables[r].Rows[3].Select();
                else
                    break;
            }
            d.Tables[d.Tables.Count].Delete();
            wdApp.ScreenUpdating = true;
            SystemSounds.Beep.Play();//Beep
            //https://blog.kkbruce.net/2019/03/csharpformusicplay.html#.YHiXtqzivsQ

        }


    }
}
