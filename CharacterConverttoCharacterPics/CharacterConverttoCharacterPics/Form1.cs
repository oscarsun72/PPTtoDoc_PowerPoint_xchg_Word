﻿using System;
using System.Drawing;
using System.Windows.Forms;
using WinWord = Microsoft.Office.Interop.Word;



namespace CharacterConverttoCharacterPics
{
    public partial class Form1 : Form
    {
        //static PowerPnt.Application pptApp; 
        //static  WinWord.Application wdApp;
        //PowerPnt.Presentation ppt;
        //WinWord.Document doc; 
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //if (ppnt==null||wwrd==null)
            //{
            //    MessageBox.Show("請開啟 Word 與 PowerPoint 再繼續");                
            //}
            //else
            //{
            //    pptApp = (PowerPnt.Application)ppnt;
            //    wdApp = (WinWord.Application)wwrd;                
            //    if (wdApp.Documents.Count > 0)
            //    {
            //        doc = wdApp.ActiveDocument;
            //        textBox1.Text = doc.FullName;
            //    }
            //    ppt = pptApp.ActivePresentation;
            //}

        }

        private void button1_Click(object sender, EventArgs e)
        {
            WinWord.Document wd = fontsPics.getFontCharacterset(textBox2.Text);
            if (wd != null)
            {
                BackColor = Color.Red;
                string picFolder = textBox1.Text,fontname= textBox2.Text;
                if (picFolder.IndexOf(fontname)==-1)
                { picFolder += ("\\" + fontname + "\\"); }
                fontsPics.addCharsSlidesExportPng(wd,
                    fontsPics.prepareFontPPT(fontname, float.Parse(textBox3.Text))
                    , picFolder);
                BackColor = Color.Green;
            }
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = Clipboard.GetText();
        }
    }
}
