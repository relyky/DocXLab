using System;
using System.IO;
using Xceed.Words.NET;

namespace DocXLab
{
    /// <summary>
    /// DocX Labbing, ref → http://cathalscorner.blogspot.tw/
    /// </summary>
    internal class DocXLab
    {
        public static void CreateDocX_Hello(FileInfo fi)
        {
            using (DocX document = DocX.Create(fi.FullName))
            {
                // resoruce : 
                // 注意此處 font
                Xceed.Words.NET.Font font1 = new Font("Arial Black");
                Xceed.Words.NET.Font font2 = new Font("標楷體");

                // Add a new Paragraph to the document.
                Paragraph p = document.InsertParagraph()
                                        .Append("Hello World.").Font(font1).FontSize(24.0)
                                        .Append("大家好。").Font(font2).FontSize(36.0);

                // Add the 2nd paragraph to the document.
                Paragraph p2 = document.InsertParagraph()
                                        .Append("中文與英、數字體要各別選取。中文與英、數字體要各別選取。中文與英、數字體要各別選取。這事不重要只是想講三次以拉長字數。");

                // append with "Formatting".
                Formatting fmt = new Formatting();
                fmt.FontFamily = new Font("微軟正黑體");
                fmt = new Formatting()
                {
                    FontFamily = new Font("微軟正黑體"),
                    FontColor = System.Drawing.Color.Blue,
                    Size = 8d,                    
                };

                document.InsertParagraph().AppendLine().AppendLine(); // 區隔上下文字段落

                document.InsertParagraph()
                    .Append("以下段落使用了\"Formatting\"", fmt);

                document.InsertParagraph()
                    .Append("一個朋友告訴我的小故事，一直盤踞在我腦海裡。主角既非俊男亦非美女，牠只是一隻「螳螂」。故事發生在瑞芳火車站，時間是夏日的午後。"
                           , fmt);

                document.InsertParagraph()
                    .Append("等車是無聊的事，尤其是一個辦完公事的老男人，這老男人一輩子都在各公家機關洽公，這時的他必須在清冷的月台上等四十分鐘的車。他無聊到被一隻鐵軌上的螳螂吸引住視線。"
                           , fmt);

                document.InsertParagraph()
                    .Append("這隻螳螂好像比他更無聊，大熱天的午後，鐵軌熱得都快冒煙了，他彷彿看到螳螂每動一隻腳就像被燙到了似地縮了一下，用踮著腳尖的姿勢在爬行。"
                           , fmt);

                // Save the document.
                document.Save();
            }
        }

        /// <summary>
        /// append some stuff: Hyperlink, Table, Image
        /// </summary>
        public static void CreateDocX_SomeStuff(FileInfo fi)
        {
            using (DocX document = DocX.Create(fi.FullName))
            {
                // Add a hyperlink into the document.
                Hyperlink link = document.AddHyperlink("link", new Uri("http://www.google.com"));

                // Add a Table into the document.
                Table table = document.AddTable(2, 2);
                table.Design = TableDesign.ColorfulGridAccent2;
                table.Alignment = Alignment.center;
                table.Rows[0].Cells[0].Paragraphs[0].Append("1");
                table.Rows[0].Cells[1].Paragraphs[0].Append("2");
                table.Rows[1].Cells[0].Paragraphs[0].Append("3");
                table.Rows[1].Cells[1].Paragraphs[0].Append("4");

                // Add an image into the document.    
                Image image = document.AddImage("Image.jpg");

                // Create a picture (A custom view of an Image).
                Picture picture = image.CreatePicture();
                picture.Rotation = 10;
                picture.SetPictureShape(BasicShapes.cube);

                // Insert a new Paragraph into the document.
                Paragraph title = document.InsertParagraph().Append("Test").FontSize(20).Font(new Font("Comic Sans MS"));
                title.Alignment = Alignment.center;

                // Insert a new Paragraph into the document.
                Paragraph p1 = document.InsertParagraph();
                p1.AppendLine("This line contains a ").Append("bold").Bold().Append(" word.");
                p1.AppendLine("Here is a cool ").AppendHyperlink(link).Color(System.Drawing.Color.Blue).Append(".");
                p1.AppendLine();
                p1.AppendLine("Check out this picture ").AppendPicture(picture).Append(" its funky don't you think?");
                p1.AppendLine();
                p1.AppendLine("Can you check this Table of figures for me?");
                p1.AppendLine();

                // Insert the Table after Paragraph 1.
                p1.InsertTableAfterSelf(table);

                // Insert a new paragraph
                Paragraph p2 = document.InsertParagraph();
                p2.AppendLine("Is it correct?");

                // Save this document.
                document.Save();
            }
        }

        /// <summary>
        /// 產生制式化文件
        /// </summary>
        public static void CreateDocX_FormulatedDocument(FileInfo fi)
        {
            // 參數。應該自外帶入。此處為測試，寫在裡面。
            var info = new
            {
                postDate = "2015/12/07",
                custName = "歐陽重天",
                custAddr1 = "１１２６５台北市北投區",
                custAddr2 = "沒這條路三段６６號２樓之１",
                caseNo = "201205-990137",
                disputeAmt = 7530,
                flag1 = false,
                flag2 = true,
                flag3 = false,
                flag4 = true,
                otherDesc = "沒錢可付只好洗碗"
            };

            using (DocX doc = DocX.Create(fi.FullName))
            {
                //Add custom properties to document.
                doc.AddCustomProperty(new CustomProperty("CompanyName", "亞洲志遠科技"));
                doc.AddCustomProperty(new CustomProperty("Product", "DocX練習"));
                doc.AddCustomProperty(new CustomProperty("Address", "新北市中正路７５５號７樓"));
                doc.AddCustomProperty(new CustomProperty("Date", DateTime.Now));

                // resource
                Font fontC = new Font("標楷體");
                Font fontE = new Font("Tahoma");
                Font fontN = new Font("Verdana");
                Font fontG = new Font("微軟正黑體");

                // banner
                Image image = doc.AddImage("Image2.png");
                Picture pic = image.CreatePicture();
                doc.InsertParagraph()
                    .AppendPicture(pic)
                    .Alignment = Alignment.right;

                // prefix header
                doc.InsertParagraph()
                    .Append(info.postDate).Font(fontE)
                    .AppendLine()
                    .AppendLine(info.custName).Font(fontC)
                    .AppendLine()
                    .AppendLine(info.custAddr1).Font(fontC)
                    .AppendLine(info.custAddr2).Font(fontC)
                    .AppendLine()
                    .AppendLine();

                // title
                doc.InsertParagraph("爭議款結案通知書")
                    .Font(fontC)
                    .FontSize(20d)
                    .SpacingAfter(10d)
                    .Alignment = Alignment.center;

                // paragraph 1
                Paragraph p1 = doc.InsertParagraph();
                p1.Append("案件編號：").Append(info.caseNo).Font(fontN)
                    .AppendLine("爭議金額：").Append(string.Format("NTD${0:N0}元", info.disputeAmt)).Font(fontN)
                    .AppendLine();
               
                // paragraph 2
                Paragraph p2 = doc.InsertParagraph();
                p2.IndentationFirstLine = 1.0f; // 第一行縮排１公分
                p2.Append("台端於日前致電本行要求處理之爭議款項，已因下列原因結案，特發此函通知。")
                    .AppendLine();

                // item 1
                Paragraph p2a = doc.InsertParagraph();
                p2a.IndentationBefore = 1.0f; // 凸排一公分
                p2a.IndentationFirstLine = -1.0f;
                p2a.Append(info.flag1 ? "■" : "□")
                    .Append(" 1.\t商店同意退回上述爭議款款項，因本行於您提出爭議時即以「帳務調整爭議款」之科目先行調整您的信用卡帳務，現爭議款確定無須支付，本行將逕做結案處理。")
                    .SpacingAfter(10d);

                // item 2
                Paragraph p2b = doc.InsertParagraph();
                p2b.IndentationBefore = 1.0f; // 凸排一公分
                p2b.IndentationFirstLine = -1.0f;
                p2b.Append(info.flag2 ? "■" : "□")
                    .Append(" 2.\t商店主動退回上述爭議款款項，此筆退款將出現於您近期帳單中，敬請查核。")
                    .AppendLine()
                    .SpacingAfter(10d);

                // item 3
                Paragraph p2c = doc.InsertParagraph();
                p2c.IndentationBefore = 1.0f; // 凸排一公分
                p2c.IndentationFirstLine = -1.0f;
                p2c.Append(info.flag3 ? "■" : "□")
                    .Append(" 3.\t商店同意退回上述爭議款款項，本行已於近期帳單中以「帳務調整爭議款」科目退款給您。")
                    .SpacingAfter(10d);

                // item 4
                Paragraph p2d = doc.InsertParagraph();
                p2d.IndentationBefore = 1.0f; // 凸排一公分
                p2d.IndentationFirstLine = -1.0f;
                p2d.Append(info.flag4 ? "■" : "□")
                    .Append(" 4.\t其他：")
                    .Append(info.otherDesc + new string('　', 30 - info.otherDesc.Length)+".").UnderlineStyle(UnderlineStyle.singleLine)
                    .SpacingAfter(10d);

                // tail
                doc.InsertParagraph().AppendLine("謹祝　　商祺")
                    .Font(fontG);
                doc.InsertParagraph().AppendLine("千陽號銀行信用卡爭議帳款小組　敬上")
                    .Font(fontG)
                    .Alignment = Alignment.right;

                // Save this document.
                doc.Save();
            }
        }

        /// <summary>
        /// 產生制式化文件 : 套版 with 書籤
        /// </summary>
        public static void CreateDocX_WithTplDocument(FileInfo fi, FileInfo fiTpl)
        {
            // 參數。應該自外帶入。此處為測試，寫在裡面。
            var info = new
            {
                postDate = "2015/12/07",
                custName = "歐陽九重天",
                custAddr1 = "１１２６５台北市北投區",
                custAddr2 = "沒這條路三段６６號２樓之１２３",
                caseNo = "201205-990137",
                disputeAmt = 9097530,
                flag1 = false,
                flag2 = false,
                flag3 = true,
                flag4 = false,
                otherDesc = ""
            };

            // copy tpl to target
            fiTpl.CopyTo(fi.FullName, true);

            // update the bookmarks in the target
            using (DocX doc = DocX.Load(fi.FullName))
            {
                // 以「書籤」套印
                doc.Bookmarks["POST_DATE"].SetText(info.postDate);
                doc.Bookmarks["CUSTOMER_NAME"].SetText(info.custName);
                doc.Bookmarks["CUSTOMER_ADDR1"].SetText(info.custAddr1);
                doc.Bookmarks["CUSTOMER_ADDR2"].SetText(info.custAddr2);

                doc.Bookmarks["CASE_NO"].SetText(info.caseNo);
                doc.Bookmarks["DISPUTE_AMOUNT"].SetText(string.Format("NTD${0:N0}元", info.disputeAmt));
                doc.Bookmarks["FLAG_1"].SetText(info.flag1 ? "■" : "□");
                doc.Bookmarks["FLAG_2"].SetText(info.flag2 ? "■" : "□");
                doc.Bookmarks["FLAG_3"].SetText(info.flag3 ? "■" : "□");
                doc.Bookmarks["FLAG_4"].SetText(info.flag4 ? "■" : "□");
                doc.Bookmarks["OTHER_DESC"].SetText(info.otherDesc + new string('　', 30 - info.otherDesc.Length) + ".");

                // Save this document.
                doc.Save();
            }
        }

        public static void CreateDocX_SimpleTable(FileInfo fi)
        {
            // Create a document.
            using (DocX document = DocX.Create(fi.FullName))
            {
                // Add a title
                document.InsertParagraph("Inserting table")
                    .FontSize(20d)
                    .Bold()
                    .SpacingAfter(10d)
                    .Alignment = Alignment.center;

                document.InsertParagraph()
                    .Append("可在現有段落中插入Table。加入image等等。");

                // Add a Table into the document and sets its values.
                var t = document.AddTable(5, 2);
                t.Design = TableDesign.ColorfulListAccent1;
                t.Alignment = Alignment.center;
                t.Rows[0].Cells[0].Paragraphs[0].Append("Mike");
                t.Rows[0].Cells[1].Paragraphs[0].Append("65");
                t.Rows[1].Cells[0].Paragraphs[0].Append("Kevin");
                t.Rows[1].Cells[1].Paragraphs[0].Append("62");
                t.Rows[2].Cells[0].Paragraphs[0].Append("Carl");
                t.Rows[2].Cells[1].Paragraphs[0].Append("60");
                t.Rows[3].Cells[0].Paragraphs[0].Append("Michael");
                t.Rows[3].Cells[1].Paragraphs[0].Append("59");
                t.Rows[4].Cells[0].Paragraphs[0].Append("Shawn");
                t.Rows[4].Cells[1].Paragraphs[0].Append("57");

                // Add a row at the end of the table and sets its values.
                var r = t.InsertRow();
                r.Cells[0].Paragraphs[0].Append("Mario");
                r.Cells[1].Paragraphs[0].Append("54");

                // Add a row at the end of the table which is a copy of another row, and sets its values.
                var newPlayer = t.InsertRow(t.Rows[2]);
                newPlayer.ReplaceText("Carl", "Max");
                newPlayer.ReplaceText("60", "50");

                // Add an image into the document.    
                var image = document.AddImage(@"Image2.png");
                var picture = image.CreatePicture(25, 100);

                // Calculate totals points from second column in table.
                var totalPts = 0;
                foreach (var row in t.Rows)
                {
                    totalPts += int.Parse(row.Cells[1].Paragraphs[0].Text);
                }

                // Add a row at the end of the table and sets its values.
                var totalRow = t.InsertRow();
                totalRow.Cells[0].Paragraphs[0].Append("Total for ").AppendPicture(picture);
                totalRow.Cells[1].Paragraphs[0].Append(totalPts.ToString());
                totalRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;

                // Insert a new Paragraph into the document.
                var p = document.InsertParagraph("Xceed Top Players Points:");
                //p.SpacingAfter(40d);

                // Insert the Table after the Paragraph.
                p.InsertTableAfterSelf(t);

                // 
                document.Save();
            }

        }

        public static void CreateDocX_SimpleTable2(FileInfo fi)
        {
            // Create a document.
            using (DocX document = DocX.Create(fi.FullName))
            {
                // Add a title
                document.InsertParagraph("Simple table")
                    .FontSize(20d)
                    .Bold()
                    .SpacingAfter(10d)
                    .Alignment = Alignment.center;

                // Add a Table into the document and sets its values.
                var t = document.AddTable(6, 2);
                t.Design = TableDesign.TableGrid;
                t.Alignment = Alignment.center;

                // fill caption
                t.Rows[0].Cells[0].FillColor = System.Drawing.Color.LightGray;
                t.Rows[0].Cells[0].Paragraphs[0].Append("名稱");
                t.Rows[0].Cells[1].FillColor = System.Drawing.Color.LightGray;
                t.Rows[0].Cells[1].Paragraphs[0].Append("值");

                // fill content
                t.Rows[1].Cells[0].Paragraphs[0].Append("Mike");
                t.Rows[1].Cells[1].Paragraphs[0].Append("65");
                t.Rows[2].Cells[0].Paragraphs[0].Append("Kevin");
                t.Rows[2].Cells[1].Paragraphs[0].Append("62");
                t.Rows[3].Cells[0].Paragraphs[0].Append("Carl");
                t.Rows[3].Cells[1].Paragraphs[0].Append("60");
                t.Rows[4].Cells[0].Paragraphs[0].Append("Michael");
                t.Rows[4].Cells[1].Paragraphs[0].Append("59");
                t.Rows[5].Cells[0].Paragraphs[0].Append("Shawn");
                t.Rows[5].Cells[1].Paragraphs[0].Append("57");

                document.InsertParagraph()
                    .InsertTableAfterSelf(t);

                // 
                document.Save();
            }

        }

        /// <summary>
        /// 產生制式化文件 : 套版 Table
        /// </summary>
        public static void CreateDocX_WithTableTplDocument(FileInfo fi, FileInfo fiTpl)
        {
            // 參數。應該自外帶入。此處為測試，寫在裡面。
            var info = new
            {
                billDate = "2015/12/07",
                billAmt = 987654,
                items = new[] {
                    new { signDate = "10/01", enterDate = "10/25", itemDesc = "信義威秀影城", area = "Taipie", amount = 650m },
                    new { signDate = "10/06", enterDate = "10/25", itemDesc = "統一麵", area = "Taipie", amount = 3000m },
                    new { signDate = "10/22", enterDate = "10/25", itemDesc = "誠品生活股份有限公司", area = "Taipie", amount = 987654m },
                    new { signDate = "10/24", enterDate = "10/25", itemDesc = "代繳電話", area = "Kaohsiung", amount = 900m },
                    new { signDate = "10/25", enterDate = "10/25", itemDesc = "愛馬仕", area = "Kaohsiung", amount = 20900m },
                },
            };

            // copy tpl to target
            fiTpl.CopyTo(fi.FullName, true);

            // update the bookmarks in the target
            using (DocX doc = DocX.Load(fi.FullName))
            {
                //## 以「Table」依範本先把「Row」筆數準備好
                // 複製需要的筆數
                Table table1 = doc.Tables[0];
                for (int i = 0; i < info.items.Length; i++)
                {
                    Row newRow = table1.InsertRow(table1.Rows[3], 5); // 複製“第3筆”並插入到“第5筆”
                }
                table1.RemoveRow(4); // 把多餘的樣板row移除。
                table1.RemoveRow(3); // 把多餘的樣板row移除。
                table1.RemoveRow(2); // 把多餘的樣板row移除。

                //// 整理格線 : 也可程式處理 border
                //newRow.Cells[0].SetBorder(TableCellBorderType.Top, new Border(BorderStyle.Tcbs_none, BorderSize.one, 0, System.Drawing.Color.Black));

                // 再填值
                decimal totalAmount = 0m; // 小計
                for (int i = 0; i < info.items.Length; i++)
                {
                    Row row = table1.Rows[i + 2]; // 從“第２筆”開始填值。
                    row.Cells[0].Paragraphs[0].Append(info.items[i].signDate);
                    row.Cells[1].Paragraphs[0].Append(info.items[i].enterDate);
                    row.Cells[2].Paragraphs[0].Append(info.items[i].itemDesc);
                    row.Cells[3].Paragraphs[0].Append(info.items[i].area);
                    row.Cells[4].Paragraphs[0].Append(info.items[i].amount.ToString("N0"));
                    //
                    totalAmount += info.items[i].amount;
                }

                // 小計
                Row totalRow = table1.Rows[info.items.Length + 2];
                totalRow.Cells[4].Paragraphs[0].Append(totalAmount.ToString("N0"));

                //## 以「Table」套印
                Table table2 = doc.Tables[1];
                //帳單結帳日
                table2.Rows[0].Cells[1].Paragraphs[0].Append(info.billDate);
                //本期應付帳款總計
                table2.Rows[0].Cells[3].Paragraphs[0].Append(info.billAmt.ToString("N0"));

                // Save this document.
                doc.Save();
            }
        }
    }

}
