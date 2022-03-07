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
using static DocXLab.DocXLab;

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
            var workFolder = new DirectoryInfo(@"C:\Temp");

            FileInfo fi = new FileInfo(Path.Combine(workFolder.FullName, "Test.docx"));

            DocXLab.CreateDocX_Hello(fi);

            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Create a document.
            var workFolder = new DirectoryInfo(@"C:\Temp");
            FileInfo fi = new FileInfo(Path.Combine(workFolder.FullName, "Test2.docx"));

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
            var workFolder = new DirectoryInfo(@"C:\Temp");
            FileInfo fi = new FileInfo(Path.Combine(workFolder.FullName, "Test6.docx"));

            DocXLab.CreateDocX_SimpleTable(fi);

            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var workFolder = new DirectoryInfo(@"C:\Temp");
            FileInfo fi = new FileInfo(Path.Combine(workFolder.FullName, "Test7.docx"));

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

        private void button10_Click(object sender, EventArgs e)
        {
            // 模擬測試結果
            var testResultList = new List<TestResultInfo>(new TestResultInfo[]
            {
                new TestResultInfo() { ItemSn="1", Name="王＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="2", Name="陳＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="3", Name="李＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="4", Name="張＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="5", Name="鄭＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="6", Name="歐陽＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="7", Name="司徒＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="8", Name="上官＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="9", Name="諸葛＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="10", Name="LI, HSIAO-YA", TestResult="negative" },
                new TestResultInfo() { ItemSn="11", Name="Peggy S.M. Wang", TestResult="negative" },
                new TestResultInfo() { ItemSn="12", Name="歐陽＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="13", Name="司徒＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="14", Name="上官＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="15", Name="諸葛＊＊", TestResult="negative" },
                new TestResultInfo() { ItemSn="16", Name="LI, HSIAO-YA", TestResult="negative" },
                new TestResultInfo() { ItemSn="17", Name="Peggy S.M. Wang", TestResult="negative" },
            });

            var workFolder = new DirectoryInfo(@"C:\Temp"); // < ------參數化
            FileInfo fi = new FileInfo(Path.Combine(workFolder.FullName, "上傳附件範本檔.docx"));

            using (var fs = fi.OpenWrite())
            {
                string errMsg = DocXLab.CreateDocX_SimpleTable3(fs, testResultList);
                if (errMsg == "SUCCESS")
                {
                    // 成功
                }
                else
                {
                    // 若失敗! 顯示錯誤訊息。
                    MessageBox.Show(errMsg);
                }
            }

            // show the result.
            System.Diagnostics.Process.Start(fi.FullName);
        }
    }

}
