using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using MvvX.Plugins.OpenXMLSDK.Platform.Word;
using MvvX.Plugins.OpenXMLSDK.Word;
using MvvX.Plugins.OpenXMLSDK.Word.Models;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;
using MvvX.Plugins.OpenXMLSDK.Word.Tables.Models;

namespace MvvX.Plugins.OpenXMLSDK.TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            // Debut test report engine
            using (IWordManager word = new WordManager())
            {
                var template = GetTemplateDocument();
                var context = new ContextModel();
                context.AddItem("#KeyTest1#", new StringModel("la la la"));
                context.AddItem("#KeyTest2#", new StringModel("toto"));

                var res = word.GenerateReport(template, context);

                // test ecriture fichier
                File.WriteAllBytes("testeric.docx", res);
                Process.Start("testeric.docx");
                return;
            }

            // fin test report engine

            var resourceName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Global.docx");

            if (!Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results")))
                Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results"));

            string finalFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results", "FinalDoc_Test_OrientationParagraph-" + DateTime.Now.ToFileTime() + ".docx");
            using (IWordManager word = new WordManager())
            {
                // TODO for debug : use your test file :
                word.OpenDocFromTemplate(resourceName, finalFilePath, true);

                //    word.SaveDoc();
                //    word.CloseDoc();
                //}
                // Insertion de texte dans un bookmark
                // wordManager.SetTextOnBookmark("Insert_Documents", "Hi !");

                // Insertion de liste à puce
                //int numberId = word.CreateBulletList();

                //var p1 = word.CreateParagraphForRun(word.CreateRunForText("coucou"), new ParagraphPropertiesModel() { NumberingProperties = new NumberingPropertiesModel() { NumberingId = numberId, NumberingLevelReference = 0 } });

                //var p2 = word.CreateParagraphForRun(word.CreateRunForText("ligne2"), new ParagraphPropertiesModel() { NumberingProperties = new NumberingPropertiesModel() { NumberingId = numberId, NumberingLevelReference = 0 } });

                //var p3 = word.CreateParagraphForRun(word.CreateRunForText("ligne21"), new ParagraphPropertiesModel() { NumberingProperties = new NumberingPropertiesModel() { NumberingId = numberId, NumberingLevelReference = 1 } });

                //var pp = new List<IParagraph>() { p1, p2 , p3};
                //word.SetParagraphsOnBookmark("Insert_Documents", pp);

                // test subtemplate
                //using (IWordManager subWord = new WordManager())
                //{
                //    subWord.OpenDocFromTemplate(resourceName);
                //    // test insert html
                //    subWord.SetHtmlOnBookmark("Insert_Documents", "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\"><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" /><title>Untitled</title><style type=\"text/css\">\r\np { margin-top: 0px;margin-bottom: 12px;line-height: 1.15; } \r\nbody { font-family: 'Arial';font-style: Normal;font-weight: normal;font-size: 12px; } \r\n.Normal { telerik-style-type: paragraph;telerik-style-name: Normal;border-collapse: collapse; } \r\n.TableNormal { telerik-style-type: table;telerik-style-name: TableNormal;border-collapse: collapse; } \r\n.s_CDEA781A { telerik-style-type: local;font-family: 'Arial';font-style: Normal;font-weight: bold;font-size: 12px;color: #000000; } \r\n.s_1E7640DD { telerik-style-type: local;font-family: 'Arial';font-style: Normal;font-weight: normal;font-size: 12px;color: #000000; } \r\n.p_80A10895 { telerik-style-type: local;margin-left: 24px;text-indent: 0px; } \r\n.p_6CC438D { telerik-style-type: local;margin-right: 0px;margin-left: 24px;text-indent: 0px; } \r\n.s_A9E8602F { telerik-style-type: local;font-family: 'Arial';font-style: Italic;font-weight: bold;font-size: 12px;color: #000000; } \r\n.s_242FFA2F { telerik-style-type: local;font-family: 'Arial';font-style: Italic;font-weight: normal;font-size: 12px;color: #000000; } \r\n.s_46C5A272 { telerik-style-type: local;font-family: 'Arial';font-style: Italic;font-weight: normal;font-size: 12px;color: #000000;text-decoration: underline; } \r\n.s_D02E313C { telerik-style-type: local;font-family: 'Arial';font-style: Normal;font-weight: normal;font-size: 12px;color: #000000;text-decoration: underline; } \r\n.p_5A0704CA { telerik-style-type: local;margin-right: 0px;margin-left: 48px;text-indent: 0px; } \r\n.p_146E745D { telerik-style-type: local;margin-right: 0px;margin-left: 72px;text-indent: 0px; } \r\n.s_8795030E { telerik-style-type: local;font-style: Normal;font-weight: normal;text-decoration: underline; } </style></head><body><p class=\"Normal \">Test rich <span class=\"s_CDEA781A\">text </span><span class=\"s_1E7640DD\"></span></p><ul style=\"list-style-type:disc\"><li value=\"1\" class=\"Normal p_80A10895\"><span class=\"s_CDEA781A\">bold</span></li><li value=\"2\" class=\"Normal p_6CC438D\"><span class=\"s_A9E8602F\">italic</span><span class=\"s_242FFA2F\">ddddd</span></li><li value=\"3\" class=\"Normal p_6CC438D\"><span class=\"s_46C5A272\">u</span><span class=\"s_D02E313C\">nderline</span></li><ul style=\"list-style-type:disc\"><li value=\"1\" class=\"Normal p_5A0704CA\"><span class=\"s_D02E313C\">l</span><span class=\"s_1E7640DD\">vl2</span></li><li value=\"2\" class=\"Normal p_5A0704CA\"><span class=\"s_1E7640DD\">hop</span></li><ul style=\"list-style-type:disc\"><li value=\"1\" class=\"Normal p_146E745D\"><span class=\"s_1E7640DD\">lvl3</span></li><li value=\"2\" class=\"Normal p_146E745D\"><span class=\"s_1E7640DD\">...</span><span class=\"s_8795030E\"></span></li></ul></ul></ul></body></html>");
                //    subWord.CloseDoc();
                //    using (Stream stream = subWord.GetMemoryStream())
                //    {
                //        word.SetSubDocumentOnBookmark("Insert_Documents", stream);
                //    }
                //}

                // Insertion d'une table dans un bookmark
                // Propriété du Tableau
                var tableProperty = new TablePropertiesModel()
                {
                    TableBorders = new TableBordersModel()
                    {
                        TopBorder = new TableBorderModel() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds },
                        LeftBorder = new TableBorderModel() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
                        RightBorder = new TableBorderModel() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
                        BottomBorder = new TableBorderModel() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds }
                    },
                    TableWidth = new TableWidthModel()
                    {
                        Width = "4900",
                        Type = TableWidthUnitValues.Pct
                    },
                    Layout = new TableLayoutModel() { Type = TableLayoutValues.Fixed }
                };
                // Lignes du premier tableau pour les constats checked
                var lines = new List<ITableRow>();

                for (int i = 0; i < 3; i++)
                {
                    var borderTopIsOK = new TableBorderModel();
                    if (i != 0)
                        borderTopIsOK.BorderValue = BorderValues.Nil;

                    // Première ligne
                    var text = word.CreateRunForText("Header Number : " + i,
                            new RunPropertiesModel()
                            {
                                Bold = true,
                                FontSize = "24",
                                RunFonts = new RunFontsModel()
                                {
                                    Ascii = "Courier New",
                                    HighAnsi = "Courier New",
                                    EastAsia = "Courier New",
                                    ComplexScript = "Courier New"
                                }
                            });

                    var cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(text, new TableCellPropertiesModel() {
                            Gridspan = new GridSpanModel() { Val = 2 },
                            Shading = new ShadingModel()
                            {
                                Fill = "F7941F"
                            },
                            TableCellWidth = new TableCellWidthModel()
                            {
                                Width = "8862"
                            },
                            TableCellBorders = new TableCellBordersModel()
                            {
                                TopBorder = borderTopIsOK
                            },
                            TableVerticalAlignementValues = TableVerticalAlignmentValues.Center
                        }),
                        word.CreateTableCell(word.CreateRun(), new TableCellPropertiesModel() {
                                    TableCellWidth = new TableCellWidthModel()
                                    {
                                        Width = "246"
                                    },
                                    Shading = new ShadingModel()
                                    {
                                        Fill = "F7941F"
                                    },
                                    TableCellBorders = new TableCellBordersModel() {
                                                TopBorder = borderTopIsOK
                                    }
                        })
                    };
                    lines.Add(word.CreateTableRow(cellules, new TableRowPropertiesModel()
                    {
                        TableRowHeight = new TableRowHeightModel()
                        {
                            Val = 380
                        }
                    }));

                    //// Deuxième ligne
                    //text = word.CreateRunForText("Comments", new RunPropertiesModel() { Bold = true });
                    //cellules = new List<ITableCell>()
                    //{
                    //    word.CreateTableCell(word.CreateImage(@"c:\temp\Tulips.jpg", new Drawing.Pictures.Model.PictureModel() {
                    //        ImagePartType   = Packaging.ImagePartType.Jpeg,
                    //        MaxHeight = 10,
                    //        MaxWidth = 500
                    //    }), new TableCellPropertiesModel() {
                    //                TableCellWidth = new TableCellWidthModel()
                    //                {
                    //                    Width = "4890"
                    //                }
                    //    }),
                    //    word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
                    //                Fusion = true,
                    //                TableCellWidth = new TableCellWidthModel()
                    //                {
                    //                    Width = "4218"
                    //                },
                    //                Gridspan = new GridSpanModel() { Val = 2 }
                    //    })
                    //};
                    //lines.Add(word.CreateTableRow(cellules));

                    // Troisième ligne
                    text = word.CreateRunForText("Texte du Constat Number : " + i, new RunPropertiesModel());
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(text, new TableCellPropertiesModel() {
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4890"
                                                },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    BottomBorder = new TableBorderModel() {
                                                        Color = "FF0019"
                                                        }
                                                }
                        }),
                        word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
                                                Fusion = true,
                                                FusionChild = true,
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4218"
                                                },
                                                Gridspan = new GridSpanModel() { Val = 2 },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    BottomBorder = new TableBorderModel() {
                                                        Color = "FF0019" }
                                                    }
                        })
                    };
                    lines.Add(word.CreateTableRow(cellules));

                    // Quatrième ligne
                    text = word.CreateRunForText("Risques", new RunPropertiesModel() { Bold = true });
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(text, new TableCellPropertiesModel() {
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4890"
                                                },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    TopBorder = new TableBorderModel()
                                                    {
                                                        Color = "00FF19"
                                                    }
                                                }
                        }),
                        word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
                                                Fusion = true,
                                                FusionChild = true,
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4218"
                                                },
                                                Gridspan = new GridSpanModel() { Val = 2 },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    TopBorder = new TableBorderModel()
                                                    {
                                                        Color = "00FF19"
                                                    }
                                                }
                        })
                    };
                    lines.Add(word.CreateTableRow(cellules));

                    // Cinquième ligne
                    text = word.CreateRunForText("Texte du Risque Number : " + i, new RunPropertiesModel());
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(text, new TableCellPropertiesModel() {
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4890"
                                                },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    TopBorder = new TableBorderModel() {
                                                        BorderValue = BorderValues.Nil }
                                                }
                        }),
                        word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
                                                Fusion = true,
                                                FusionChild = true,
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4218"
                                                },
                                                Gridspan = new GridSpanModel() { Val = 2 },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    TopBorder = new TableBorderModel() {
                                                        BorderValue = BorderValues.Nil }
                                                    }
                        })
                    };
                    lines.Add(word.CreateTableRow(cellules));

                    // Sixième ligne
                    text = word.CreateRunForText("Recommandations", new RunPropertiesModel() { Bold = true });
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(text, new TableCellPropertiesModel() {
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4890"
                                                },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    BottomBorder = new TableBorderModel() {
                                                        BorderValue = BorderValues.Nil }
                                                    }
                        }),
                        word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
                                                Fusion = true,
                                                FusionChild = true,
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4218"
                                                },
                                                Gridspan = new GridSpanModel() { Val = 2 },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    BottomBorder = new TableBorderModel() {
                                                        BorderValue = BorderValues.Nil }
                                                    }
                        })
                    };
                    lines.Add(word.CreateTableRow(cellules));

                    var borderBottomIsOK = new TableBorderModel()
                    {
                        BorderValue = BorderValues.Nil,
                        Color = "FF0019"
                    };

                    if (i == 2)
                        borderBottomIsOK.BorderValue = BorderValues.Single;

                    // Septième ligne
                    text = word.CreateRunForText("Texte de la Recommandation Number : " + i, new RunPropertiesModel());
                    cellules = new List<ITableCell>()
                    {
                        word.CreateTableCell(text, new TableCellPropertiesModel() {
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4890"
                                                },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil },
                                                    BottomBorder = borderBottomIsOK }
                        }),
                        word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
                                                Fusion = true,
                                                FusionChild = true,
                                                TableCellWidth = new TableCellWidthModel()
                                                {
                                                    Width = "4218"
                                                },
                                                Gridspan = new GridSpanModel() { Val = 2 },
                                                TableCellBorders = new TableCellBordersModel() {
                                                    TopBorder = new TableBorderModel() {
                                                        BorderValue = BorderValues.Nil },
                                                    BottomBorder = borderBottomIsOK }
                        })
                    };
                    lines.Add(word.CreateTableRow(cellules));
                }

                IList<IParagraph> tables = new List<IParagraph>();
                tables.Add(word.CreateParagraphForRun(word.CreateRunForTable(word.CreateTable(lines, tableProperty))));

                tables.Add(word.CreateParagraphForRun(
                    word.CreateRunForText("Test de style"),
                        new ParagraphPropertiesModel()
                        {
                            ParagraphStyleId = new ParagraphStyleIdModel()
                            {
                                Val = "Titre01"
                            },
                            SpacingBetweenLines = new SpacingBetweenLinesModel()
                            {
                                After = "800",
                                Before = "100"
                            }
                        }));

                // Lignes du deuxième tableau pour les constats unchecked
                //lines = new List<TableRow>();

                if (tables.Count > 0)
                    word.SetParagraphsOnBookmark("Insert_Documents", tables);

                word.SaveDoc();
                word.CloseDoc();
            }

            Process.Start(finalFilePath);
        }

        private static Document GetTemplateDocument()
        {
            var doc = new Document();
            doc.Styles.Add(new Style() { StyleId = "OnSiteTitle" });
            doc.Styles.Add(new Style() { StyleId = "toto", FontColor="FFFF00", FontSize="40" });
            var page1 = new Page();
            var page2 = new Page();
            doc.Pages.Add(page1);
            doc.Pages.Add(page2);
            var paragraph = new Paragraph();
            paragraph.ChildElements.Add(new Label() { Text = "Ceci est un texte", FontSize = "30", FontName = "Arial" });
            paragraph.ChildElements.Add(new Label() { Text = "#KeyTest1#", FontSize = "40", FontColor = "FF0000", Shading = "0000FF" });
            paragraph.ChildElements.Add(new Label() { Text = "#KeyTest2#", Show = false });
            page1.ChildElements.Add(paragraph);
            var p2 = new Paragraph();
            p2.Shading = "FF0000";
            p2.ChildElements.Add(new Label() { Text = "texte paragraph2", FontSize = "20" });
            p2.ChildElements.Add(new Label() { Text = "texte2 paragraph2" });
            page1.ChildElements.Add(p2);

            // page 2
            var p21 = new Paragraph();
            p21.Justification = JustificationValues.Center;
            p21.ParagraphStyleId = "OnSiteTitle";
            p21.ChildElements.Add(new Label() { Text = "texte page2", FontName="Arial" });
            page2.ChildElements.Add(p21);
            var p22 = new Paragraph();
            p22.SpacingBefore = 800;
            p22.SpacingAfter = 800;
            p22.Justification = JustificationValues.Both;
            p22.ParagraphStyleId = "toto";
            p22.ChildElements.Add(new Label() { Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse urna augue, convallis eu enim vitae, maximus ultrices nulla. Sed egestas volutpat luctus. Maecenas sodales erat eu elit auctor, eu mattis neque maximus. Duis ac risus quis sem bibendum efficitur. Vivamus justo augue, molestie quis orci non, maximus imperdiet justo. Donec condimentum rhoncus est, ut varius lorem efficitur sed. Donec accumsan sit amet nisl vel ornare. Duis aliquet urna eu mauris porttitor facilisis. " });
            page2.ChildElements.Add(p22);
            var p23 = new Paragraph();
            p23.SpacingBetweenLines = 360;
            p23.ChildElements.Add(new Label() { Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse urna augue, convallis eu enim vitae, maximus ultrices nulla. Sed egestas volutpat luctus. Maecenas sodales erat eu elit auctor, eu mattis neque maximus. Duis ac risus quis sem bibendum efficitur. Vivamus justo augue, molestie quis orci non, maximus imperdiet justo. Donec condimentum rhoncus est, ut varius lorem efficitur sed. Donec accumsan sit amet nisl vel ornare. Duis aliquet urna eu mauris porttitor facilisis. " });
            page2.ChildElements.Add(p23);

            // page 3
            var page3 = new Page();
            var p31 = new Paragraph() { FontColor = "FF0000", FontSize = "26" };
            p31.ChildElements.Add(new Label() { Text = "test héritage" });
            var p311 = new Paragraph() { FontSize = "16" };
            p311.ChildElements.Add(new Label() { Text = "blabla" });
            p31.ChildElements.Add(p311);
            page3.ChildElements.Add(p31);
            doc.Pages.Add(page3);
            return doc;
        }
    }
}