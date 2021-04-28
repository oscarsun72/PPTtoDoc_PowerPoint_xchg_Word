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
        {//將字源圖片原來的總表，分割成一個字源一個表格。且都有標題。即每表格有二列
            //此表格下面再插入一個新的表格，以便置入字源論述及字源圖片和靜態筆順20210428
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
                {
                    DocOps doc = new DocOps();
                    d = doc.openDoc(DocOps.getDocFullname());
                    if (d==null)return;
                }
            }
            else
            {
                DocOps doc = new DocOps();
                d = doc.openDoc(DocOps.getDocFullname());
                if (d == null)return;
            }
            //放在指定位置以開始
            d.Tables[1].Cell(3, 1).Range.Characters[1].Select();
            winWord.Selection Selection = d.ActiveWindow.Selection;
            Selection.Collapse(winWord.WdCollapseDirection.wdCollapseStart);
            int r, s; winWord.Cell cel; winWord.Range rng;
            winWord.InlineShape inlsp; winWord.Table tb;
            //List<WinWord.InlineShape> inlsps = new List<WinWord.InlineShape>();
            winWord.Row rw; winWord.Range rngInlSp; 
            const float picCellWidth= 122.7F, picCellHeight= 120.2F;
             r = 1;
            rng = Selection.Range;
            wdApp.ScreenUpdating = false;
            d.Tables[1].Rows.Add(); d.Tables[1].Rows.Add();//最後會留下一個表格再予刪除
            int picsCount;
            //開始逐字分割為一表格：
            while (Selection.Information[winWord.WdInformation.
                wdWithInTable])
            {
                Selection.SplitTable();
                /* 表格置中都無效
                 * Selection.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                */
                rw = Selection.Document.Tables[1].Rows[1];
                rw.Range.Copy();//準備標題列
                Selection.Document.Tables[Selection.Document.
                    Tables.Count].Range.Characters[1].Select();
                Selection.Collapse(winWord.WdCollapseDirection.
                    wdCollapseStart);
                Selection.Paste();//貼上標題列
                Selection.Document.Tables[Selection.Document.Tables.Count]
                    .Range.Characters[1].Select();
                Selection.Collapse(winWord.WdCollapseDirection.
                    wdCollapseStart);
                Selection.MoveLeft();//分割完表格，就定位
                Selection.InsertParagraphAfter();
                Selection.Collapse(winWord.WdCollapseDirection.
                    wdCollapseEnd);

                //插入表格，準備將圖片置入
                tb = Selection.Tables.Add(Selection.Range, 2, 2);
                tb.Range.Cells.VerticalAlignment = winWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;//置中對齊
                tb.Range.ParagraphFormat.Alignment = winWord.WdParagraphAlignment.wdAlignParagraphLeft; //向左對齊
                tb.Cell(1, 1).VerticalAlignment = winWord.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                tb.Columns[1].Width = 359.15F;
                tb.Columns[2].Width = picCellWidth;//圖片儲存格的寬
                tb.Rows[2].Cells.Merge();//第二列合併儲存格
                tb.Borders.InsideLineStyle = //內框樣式
                    winWord.WdLineStyle.wdLineStyleSingle;
                tb.Borders.OutsideLineStyle =//外框樣式 
                    winWord.WdLineStyle.wdLineStyleDouble;
                tb.Rows[1].Height = picCellHeight;//圖片儲存格的高
                tb.Rows[2].Height = 56;                
                //表格置中
                //此無效：tb.Range.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                //這才有效：//http://www.wordbanter.com/showthread.php?t=110960
                tb.Rows.Alignment = winWord.WdRowAlignment.
                    wdAlignRowCenter;
                tb.PreferredWidthType = winWord.WdPreferredWidthType.
                    wdPreferredWidthPoints;//https://stackoverflow.com/questions/54159142/set-table-column-widths-in-word-macro-vba
                tb.PreferredWidth = (float)549.6378;//固定表格寬度 Selection.Document.Tables[r].PreferredWidth;

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
                rng.SetRange(s, s);//記下圖片要貼上的位置
                picsCount =cel.Range.InlineShapes.Count;
                if (picsCount>0)
                {
                    inlsp = cel.Range.InlineShapes[1];
                    inlsp.Select();
                    Selection.Cut();//剪下圖片，準備移動位置
                    #region 圖片若不貼新表格中，則如下：
                    /*                    s1 = Selection.Start;
                    if (s1 > s)//兩種inlineshape圖形作用不同，故須分別處置
                    {
                        while (rng.Information[winWord.WdInformation
                            .wdWithInTable])
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

                    //foreach (WinWord.InlineShape insp in
                    //              Selection.Previous().InlineShapes)
                    //{
                    //    inlsps.Add(insp);
                    //    insp.Height += 181;
                    //    insp.Width += 181;
                    //}

                    //docx檔中的圖形有二種，貼上後插入點位置也不同，故皆須分別處置
                    if (Selection.Previous().InlineShapes.Count > 0)
                    {
                        inlsp = Selection.Previous().InlineShapes[1];
                        inlsp.LockAspectRatio = MsoTriState.msoTrue;
                        inlsp.Height = 200;
                    }//調整圖片大小
                    else
                    {//這種選取後周邊呈虛線形的圖形，與一般的圖形剪下貼上後插入點的落點會不同(會在貼上的圖形前），且用Selection.Next()也取不到它，須先將其選取 20210418
                        Selection.MoveRight(winWord.WdUnits.wdCharacter, 1, winWord.WdMovementType.wdExtend);
                        inlsp = Selection.InlineShapes[1];
                        inlsp.Height += 181;// = Selection.InlineShapes[1].Height + 181;
                        inlsp.Width += 181;//= Selection.InlineShapes[1].Height + 181;
                        Selection.Collapse(winWord.WdCollapseDirection.wdCollapseEnd);
                    }
                    //圖片置中
                    //Selection.ParagraphFormat.Alignment = WinWord.WdParagraphAlignment.wdAlignParagraphCenter;
                    
                    inlsp.Select(); Selection.Cut();//剪下圖片貼入新插入的表格中
                    */
                    #endregion
                    tb.Cell(1, 2).Range.Characters[1].Select();
                    Selection.Paste();//貼上圖片，配合儲存格調整圖片大小                    
                    if (Selection.Previous().InlineShapes.Count > 0)
                        rngInlSp = Selection.Previous();
                    else //(Selection.Next().InlineShapes.Count > 0)
                        rngInlSp = Selection.Next();
                        rngInlSp.InlineShapes[1].LockAspectRatio = MsoTriState.msoTrue;
                        rngInlSp.InlineShapes[1].Height = picCellHeight;
                        if (rngInlSp.InlineShapes[1].Width>picCellWidth)
                        {
                            rngInlSp.InlineShapes[1].Width = picCellWidth;
                        }
                    //以上圖片貼到定位且處理好其大小了
                    //離開圖片
                    Selection.MoveDown(Count:2);
                    //Selection.Collapse(WinWord.WdCollapseDirection.wdCollapseEnd);
                    //與下一分割出來的表格空2行（段）--即與下一個漢字字源表分開來（距離拉開）
                    Selection.InsertParagraphAfter(); Selection.InsertParagraphAfter();
                }
                //以上有圖時的處理，以下缺圖者亦同然：
                Selection.Document.Tables[r].Columns[8].Cells.Delete();//原來放置圖片的那欄刪除
                r+=2; //前面Tables.Add多插一表格，計數要再加1
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
