using System;
using System.IO;
using MvvX.Open_XML_SDK.Word;
using System.Collections.Generic;
using MvvX.Open_XML_SDK.Core.Word.Tables.Models;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;

namespace MvvX.Open_XML_SDK.TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var resourceName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Global.docx");

            if (!Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results")))
                Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results"));

            using (var word = new WordManager())
            {
                // TODO for debug : use your test file :
                word.OpenDocFromTemplate(resourceName, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results", "FinalDoc_Test_OrientationParagraph-" + DateTime.Now.ToFileTime() + ".docx"), true);

                // Insertion de texte dans un bookmark
                // wordManager.SetTextOnBookmark("Insert_Documents", "Hi !");

                // Insertion d'une table dans un bookmark
                // Propriété du Tableau
                var tableProperty = new TablePropertiesModel();
                tableProperty.TopBorder = new TableBorderModel() { Color = "F7941F", Size = 20 };
                tableProperty.LeftBorder = new TableBorderModel() { Color = "F7941F", Size = 20 };
                tableProperty.RightBorder = new TableBorderModel() { Color = "F7941F", Size = 20 };
                tableProperty.BottomBorder = new TableBorderModel() { Color = "F7941F", Size = 20 };

                // Lignes du premier tableau pour les constats checked
                var lines = new List<ITableRow>();

                for (int i = 0; i < 3; i++)
                {
                    var borderTopIsOK = new TableBorderModel();
                    if (i != 0)
                        borderTopIsOK.BorderValue = BorderValues.Nil;
                    
                    // Première ligne
                    var texte = word.CreateRunForTexte("Header Numero : " + i, new RunPropertiesModel() { Bold = true, FontSize = "24", Color = "FFFFFF" });
                    var cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(texte, new TableCellPropertiesModel() { Gridspan = 2, Shading = word.GetShading(fillColor: "F7941F"), /*BorderBottom = false, BorderTop = borderTopIsOK,*/ Width = "8862",
                                                BottomBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }, TopBorder = borderTopIsOK,  }),
                        word.CreateTableCell(word.CreateEmptyRun(), new TableCellPropertiesModel() { Shading = word.GetShading(fillColor: "F7941F"), /*BorderBottom = false, BorderTop = borderTopIsOK,*/ Width = "246",
                                                BottomBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }, TopBorder = borderTopIsOK })
                    };
                    lines.Add(word.CreateTableRow(cellules, new TableRowPropertiesModel() { Height = 380 }));

                    // Deuxième ligne
                    texte = word.CreateRunForTexte("Constat et commentaire", new RunPropertiesModel() { Bold = true });
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(texte, new TableCellPropertiesModel() { Width = "4890", /*BorderBottom = false, BorderTop = false*/
                                                BottomBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }, TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil } }),
                        word.CreateTableMergeCell(word.CreateEmptyRun(), new TableCellPropertiesModel() { Fusion = true, Width = "4218", /*BorderTop = false, BorderBottom = false,*/ Gridspan = 2,
                                                BottomBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }, TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }})
                    };
                    lines.Add(word.CreateTableRow(cellules));

                    // Troisième ligne
                    texte = word.CreateRunForTexte("Texte du Constat Numero : " + i, new RunPropertiesModel());
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(texte, new TableCellPropertiesModel() { Width = "4890", /*BorderTop = false, BorderBottomColor = "FF0019"*/
                                                TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }, BottomBorder = new TableBorderModel() { Color = "FF0019" } }),
                        word.CreateTableMergeCell(word.CreateEmptyRun(), new TableCellPropertiesModel() { Fusion = true, FusionChild = true, Width = "4218", /*BorderTop = false,*/ Gridspan = 2, /*BorderBottomColor = "FF0019"*/
                                                TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }, BottomBorder = new TableBorderModel() { Color = "FF0019" } })
                    };
                    lines.Add(word.CreateTableRow(cellules));

                    // Quatrième ligne
                    texte = word.CreateRunForTexte("Risques", new RunPropertiesModel() { Bold = true });
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(texte, new TableCellPropertiesModel() { Width = "4890", /*BorderBottom = false, BorderTop = true*//*, BorderTopColor = "00FF19"*/
                                                BottomBorder = new TableBorderModel() { BorderValue = BorderValues.Nil } }),
                        word.CreateTableMergeCell(word.CreateEmptyRun(), new TableCellPropertiesModel() { Fusion = true, FusionChild = true, Width = "4218", /*BorderBottom = false,*/ Gridspan = 2, /*BorderTop = true*//*, BorderTopColor = "00FF19"*/
                                                BottomBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }})
                    };
                    lines.Add(word.CreateTableRow(cellules));

                    // Cinquième ligne
                    texte = word.CreateRunForTexte("Texte du Risque Numero : " + i, new RunPropertiesModel());
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(texte, new TableCellPropertiesModel() { Width = "4890", /*BorderTop = false*/
                                                TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil } }),
                        word.CreateTableMergeCell(word.CreateEmptyRun(), new TableCellPropertiesModel() { Fusion = true, FusionChild = true, Width = "4218", /*BorderTop = false,*/ Gridspan = 2,
                                                TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil } })
                    };
                    lines.Add(word.CreateTableRow(cellules));

                    // Sixième ligne
                    texte = word.CreateRunForTexte("Recommandations", new RunPropertiesModel() { Bold = true });
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(texte, new TableCellPropertiesModel() { Width = "4890", /*BorderBottom = false*/
                                                BottomBorder = new TableBorderModel() { BorderValue = BorderValues.Nil } }),
                        word.CreateTableMergeCell(word.CreateEmptyRun(), new TableCellPropertiesModel() { Fusion = true, FusionChild = true, Width = "4218", /*BorderBottom = false,*/ Gridspan = 2,
                                                BottomBorder = new TableBorderModel() { BorderValue = BorderValues.Nil } })
                    };
                    lines.Add(word.CreateTableRow(cellules));

                    var borderBottomIsOK = new TableBorderModel() { BorderValue = BorderValues.Nil, Color = "FF0019" };
                    if (i == 2)
                        borderBottomIsOK.BorderValue = BorderValues.Single;

                    // Septième ligne
                    texte = word.CreateRunForTexte("Texte de la Recommandation Numero : " + i, new RunPropertiesModel());
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(texte, new TableCellPropertiesModel() { Width = "4890", /*BorderTop = false, BorderBottom = borderBottomIsOK, BorderBottomColor = "FF0019"*/
                                                TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }, BottomBorder = borderBottomIsOK }),
                        word.CreateTableMergeCell(word.CreateEmptyRun(), new TableCellPropertiesModel() { Fusion = true, FusionChild = true, Width = "4218", /*BorderTop = false, BorderBottom = borderBottomIsOK,*/ Gridspan = 2, /*BorderBottomColor = "FF0019"*/
                                                TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil }, BottomBorder = borderBottomIsOK })
                    };
                    lines.Add(word.CreateTableRow(cellules));
                }

                IList<IParagraph> tables = new List<IParagraph>();                
                tables.Add(word.CreateParagraphForRun(word.CreateRunForTable(word.CreateTable(lines, tableProperty))));

                // Lignes du deuxième tableau pour les constats unchecked
                //lines = new List<TableRow>();
                
                if (tables.Count > 0)
                    word.SetParagraphsOnBookmark("Insert_Documents", tables);

                word.SaveDoc();
                word.CloseDoc();
            }
        }
    }

    internal class Run
    {
        public Run()
        {
        }
    }
}
