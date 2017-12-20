using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace DocXLab
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FileInfo fi = new FileInfo(@"D:\Test.docx");

            DocXLab.CreateDocX_Hello(fi);

            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Create a document.
            FileInfo fi = new FileInfo(@"D:\Test2.docx");

            DocXLab.CreateDocX_SomeStuff(fi);

            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Create a document.
            FileInfo fi = new FileInfo(@"D:\Test3.docx");

            // 爭議款結案通知書
            DocXLab.CreateDocX_FormulatedDocument(fi);

            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Create a document.
            FileInfo fi = new FileInfo(@"D:\Test4.docx");
            FileInfo fiTpl = new FileInfo(@"DocXTpl01.docx");

            // 爭議款結案通知書
            DocXLab.CreateDocX_WithTplDocument(fi, fiTpl);

            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Create a document.
            FileInfo fiTpl = new FileInfo(@"DocXTpl01.docx");

            // show the result.
            System.Diagnostics.Process.Start(fiTpl.FullName);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FileInfo fi = new FileInfo(@"D:\Test6.docx");

            DocXLab.CreateDocX_SimpleTable(fi);

            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            FileInfo fi = new FileInfo(@"D:\Test7.docx");

            DocXLab.CreateDocX_SimpleTable2(fi);
        
            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // Create a document.
            FileInfo fi = new FileInfo(@"D:\Test8.docx");
            FileInfo fiTpl = new FileInfo(@"對帳單範例.docx");

            // 爭議款結案通知書
            DocXLab.CreateDocX_WithTableTplDocument(fi, fiTpl);

            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            // Create a document.
            FileInfo fiTpl = new FileInfo(@"對帳單範例.docx");

            // show the result.
            System.Diagnostics.Process.Start(fiTpl.FullName);
        }
    }

}
