using Microsoft.Office.Core;
using System.Media;
using winWord = Microsoft.Office.Interop.Word;
namespace Doc_PPt
{
    class TableDocOps
    {
        readonly winWord.Application wdApp; winWord.Document d; const string docName = "＃字源圖片 （象形）.docx";
        public TableDocOps()
        {

            wdApp = docApp.getDocApp();

        }


        internal void splitTableByEachRowTitleed字源圖片()
        {
            if (wdApp.Documents.Count > 0)
            {
                d = wdApp.ActiveDocument;
                if (d.Name != docName)
                {
                    foreach (winWord.Document item in wdApp.Documents)
                    {
                        if (item.Name == docName)
                        {
                            item.Activate(); d = item; break;
                        }
                    }
                }

                if (d.Name != docName)
                    d = DocOps.openDoc(DocOps.getDocFullname());
            }
            else
            {
                d = DocOps.openDoc(DocOps.getDocFullname());
            }
            d.Tables[1].Cell(3, 1).Range.Characters[1].Select();
            winWord.Selection Selection = d.ActiveWindow.Selection;
            Selection.Collapse(winWord.WdCollapseDirection.wdCollapseStart);
            int r, s, s1; winWord.Cell cel; winWord.Range rng;
            winWord.InlineShape inlsp; winWord.Table tb;
            //List<WinWord.InlineShape> inlsps = new List<WinWord.InlineShape>();
            winWord.Row rw;
            r = 1;
            rng = Selection.Range;
            wdApp.ScreenUpdating = false;
            d.Tables[1].Rows.Add(); d.Tables[1].Rows.Add();//最後會留下一個表格再予刪除
            while (Selection.Information[winWord.WdInformation.wdWithInTable])
            {
                Selection.SplitTable();
                /* 表格置中都無效
                 * Selection.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                */
                rw = Selection.Document.Tables[1].Rows[1];
                rw.Range.Copy();
                Selection.Document.Tables[Selection.Document.Tables.Count].Range.Characters[1].Select();
                Selection.Collapse(winWord.WdCollapseDirection.wdCollapseStart);
                Selection.Paste();
                Selection.Document.Tables[Selection.Document.Tables.Count].Range.Characters[1].Select();
                Selection.Collapse(winWord.WdCollapseDirection.wdCollapseStart);
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
                        while (rng.Information[winWord.WdInformation.wdWithInTable])
                        {
                            s1--;
                            rng.SetRange(s1, s1);
                        }
                    }
                    else if (s1 < s)
                    {
                        while (rng.Information[winWord.WdInformation.wdWithInTable])
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
                        Selection.MoveRight(winWord.WdUnits.wdCharacter, 1, winWord.WdMovementType.wdExtend);
                        inlsp = Selection.InlineShapes[1];
                        inlsp.Height += 181;// = Selection.InlineShapes[1].Height + 181;
                        inlsp.Width += 181;//= Selection.InlineShapes[1].Height + 181;
                    }
                    //圖片置中
                    //Selection.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                    //插入表格，將圖片置入
                    tb = Selection.Tables.Add(Selection.Range, 1, 2);
                    tb.Borders.InsideLineStyle = winWord.WdLineStyle.wdLineStyleSingle;
                    tb.Borders.OutsideLineStyle = winWord.WdLineStyle.wdLineStyleDouble;
                    //表格置中
                    //此無效：tb.Range.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                    //這才有效：//http://www.wordbanter.com/showthread.php?t=110960
                    tb.Rows.Alignment = winWord.WdRowAlignment.wdAlignRowCenter;
                    inlsp.Select(); Selection.Cut();//剪下圖片貼入表格
                    tb.Cell(1, 2).Range.Characters[1].Select();
                    Selection.Paste();
                    tb.PreferredWidthType = winWord.WdPreferredWidthType.wdPreferredWidthPoints;//https://stackoverflow.com/questions/54159142/set-table-column-widths-in-word-macro-vba
                    tb.PreferredWidth = (float)549.6378;//Selection.Document.Tables[r].PreferredWidth;
                    tb.Range.ParagraphFormat.Alignment = winWord.WdParagraphAlignment.wdAlignParagraphCenter;
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
