using System;
using System.Collections.Generic;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Images;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Core.Word.Tables.Models;

namespace MvvX.Open_XML_SDK.PCL
{
    public class WordReport
    {
        public IWordManager wordManager;

        public WordReport(IWordManager wordManager)
        {
            this.wordManager = wordManager;
        }

        public void GenerateReport()
        {
            var resourceName = @"C:\temp\Global.docx";
            var imagePath = @"C:\temp\circle.png";
            var headerCols = new List<string> { "Cell 1", "Cell 2", "Cell 3", "Cell 4", "Cell 5", "Cell 6" };

            // TODO for debug : use your test file :
            wordManager.OpenDocFromTemplate(resourceName, @"C:\temp\" + DateTime.Now.ToFileTime() + ".docx", true);

            //wordManager.SetTextOnBookmark("Insert_Documents", "Hi !");
            //wordManager.InsertPictureToBookmark("Insert_Documents", imagePath, ImageType.Png);

            var tableBorderModel = new TablePropertiesModel();
            tableBorderModel.TopBorder.BorderValue = BorderValue.Single;
            tableBorderModel.BottomBorder.BorderValue = BorderValue.Single;
            tableBorderModel.InsideHorizontalBorder.BorderValue = BorderValue.Single;
            tableBorderModel.InsideVerticalBorder.BorderValue = BorderValue.Single;
            tableBorderModel.LeftBorder.BorderValue = BorderValue.Single;
            tableBorderModel.RightBorder.BorderValue = BorderValue.Single;

            var tables = new List<IParagraph>();
            var row = new List<ITableRow>();

            var line1 = new List<ITableCell>();
            var photo = wordManager.CreateImage(imagePath, ImageType.Png);

            for (int i = 0; i < headerCols.Count; i++)
            {
                line1.Add(wordManager.CreateTableCell(wordManager.CreateTexte(headerCols[i]), new TableCellPropertiesModel()));
            }

            row.Add(wordManager.CreateTableRow(line1));

            var line2 = new List<ITableCell>();

            for (int i = 0; i < headerCols.Count; i++)
            {
                line2.Add(wordManager.CreateTableCell(wordManager.CreateImage(imagePath, ImageType.Png), new TableCellPropertiesModel()));
            }

            row.Add(wordManager.CreateTableRow(line2));

            var table = wordManager.CreateTable(row, tableBorderModel);
            var run = wordManager.CreateRun(table.ContentItem);
            var par = wordManager.CreateParagraph(run.ContentItem);
            tables.Add(par);
            wordManager.SetParagraphsOnBookmark("Insert_Documents", tables);

            wordManager.SaveDoc();
            wordManager.CloseDoc();
        }
    }
}
