namespace OpenXMLSDK.Engine.TestConsole
{
    class Program
    {
        static void Main()
        {
            ReportEngineTest.Test();

            //ValidateDocument();
            // fin test report engine

            //OldProgram();
        }

        //private static void ValidateDocument()
        //{
        //    // Debut test report engine
        //    var validator = new OpenXMLValidator();
        //    var results = validator.ValidateWordDocument("ExampleDocument.docx");
        //    File.Delete("ValidateDocument.txt");
        //    foreach(var result in results)
        //    {
        //        File.AppendAllText("ValidateDocument.txt", 
        //                            result.XmlPath + Environment.NewLine + 
        //                            result.Description + Environment.NewLine +
        //                            result.ErrorType.ToString() + Environment.NewLine +
        //                            " - - - - - - - - - ");
        //    }
        //}


        //private static void OldProgram()
        //{
        //    var resourceName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Global.docx");

        //    if (!Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results")))
        //        Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results"));

        //    string finalFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results", "FinalDoc_Test_OrientationParagraph-" + DateTime.Now.ToFileTime() + ".docx");
        //    using (var word = new WordManager())
        //    {
        //        //TODO for debug : use your test file :
        //        word.OpenDocFromTemplate(resourceName, finalFilePath, true);

        //        //    word.SaveDoc();
        //        //    word.CloseDoc();
        //        //}
        //        // Insertion de texte dans un bookmark
        //        // wordManager.SetTextOnBookmark("Insert_Documents", "Hi !");

        //        // Insertion de liste à puce
        //        //int numberId = word.CreateBulletList();

        //        //var p1 = word.CreateParagraphForRun(word.CreateRunForText("coucou"), new ParagraphPropertiesModel() { NumberingProperties = new NumberingPropertiesModel() { NumberingId = numberId, NumberingLevelReference = 0 } });

        //        //var p2 = word.CreateParagraphForRun(word.CreateRunForText("ligne2"), new ParagraphPropertiesModel() { NumberingProperties = new NumberingPropertiesModel() { NumberingId = numberId, NumberingLevelReference = 0 } });

        //        //var p3 = word.CreateParagraphForRun(word.CreateRunForText("ligne21"), new ParagraphPropertiesModel() { NumberingProperties = new NumberingPropertiesModel() { NumberingId = numberId, NumberingLevelReference = 1 } });

        //        //var pp = new List<Paragraph>() { p1, p2 , p3};
        //        //word.SetParagraphsOnBookmark("Insert_Documents", pp);

        //        // test subtemplate
        //        //using (IWordManager subWord = new WordManager())
        //        //{
        //        //    subWord.OpenDocFromTemplate(resourceName);
        //        //    // test insert html
        //        //    subWord.SetHtmlOnBookmark("Insert_Documents", "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\"><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" /><title>Untitled</title><style type=\"text/css\">\r\np { margin-top: 0px;margin-bottom: 12px;line-height: 1.15; } \r\nbody { font-family: 'Arial';font-style: Normal;font-weight: normal;font-size: 12px; } \r\n.Normal { telerik-style-type: paragraph;telerik-style-name: Normal;border-collapse: collapse; } \r\n.TableNormal { telerik-style-type: table;telerik-style-name: TableNormal;border-collapse: collapse; } \r\n.s_CDEA781A { telerik-style-type: local;font-family: 'Arial';font-style: Normal;font-weight: bold;font-size: 12px;color: #000000; } \r\n.s_1E7640DD { telerik-style-type: local;font-family: 'Arial';font-style: Normal;font-weight: normal;font-size: 12px;color: #000000; } \r\n.p_80A10895 { telerik-style-type: local;margin-left: 24px;text-indent: 0px; } \r\n.p_6CC438D { telerik-style-type: local;margin-right: 0px;margin-left: 24px;text-indent: 0px; } \r\n.s_A9E8602F { telerik-style-type: local;font-family: 'Arial';font-style: Italic;font-weight: bold;font-size: 12px;color: #000000; } \r\n.s_242FFA2F { telerik-style-type: local;font-family: 'Arial';font-style: Italic;font-weight: normal;font-size: 12px;color: #000000; } \r\n.s_46C5A272 { telerik-style-type: local;font-family: 'Arial';font-style: Italic;font-weight: normal;font-size: 12px;color: #000000;text-decoration: underline; } \r\n.s_D02E313C { telerik-style-type: local;font-family: 'Arial';font-style: Normal;font-weight: normal;font-size: 12px;color: #000000;text-decoration: underline; } \r\n.p_5A0704CA { telerik-style-type: local;margin-right: 0px;margin-left: 48px;text-indent: 0px; } \r\n.p_146E745D { telerik-style-type: local;margin-right: 0px;margin-left: 72px;text-indent: 0px; } \r\n.s_8795030E { telerik-style-type: local;font-style: Normal;font-weight: normal;text-decoration: underline; } </style></head><body><p class=\"Normal \">Test rich <span class=\"s_CDEA781A\">text </span><span class=\"s_1E7640DD\"></span></p><ul style=\"list-style-type:disc\"><li value=\"1\" class=\"Normal p_80A10895\"><span class=\"s_CDEA781A\">bold</span></li><li value=\"2\" class=\"Normal p_6CC438D\"><span class=\"s_A9E8602F\">italic</span><span class=\"s_242FFA2F\">ddddd</span></li><li value=\"3\" class=\"Normal p_6CC438D\"><span class=\"s_46C5A272\">u</span><span class=\"s_D02E313C\">nderline</span></li><ul style=\"list-style-type:disc\"><li value=\"1\" class=\"Normal p_5A0704CA\"><span class=\"s_D02E313C\">l</span><span class=\"s_1E7640DD\">vl2</span></li><li value=\"2\" class=\"Normal p_5A0704CA\"><span class=\"s_1E7640DD\">hop</span></li><ul style=\"list-style-type:disc\"><li value=\"1\" class=\"Normal p_146E745D\"><span class=\"s_1E7640DD\">lvl3</span></li><li value=\"2\" class=\"Normal p_146E745D\"><span class=\"s_1E7640DD\">...</span><span class=\"s_8795030E\"></span></li></ul></ul></ul></body></html>");
        //        //    subWord.CloseDoc();
        //        //    using (Stream stream = subWord.GetMemoryStream())
        //        //    {
        //        //        word.SetSubDocumentOnBookmark("Insert_Documents", stream);
        //        //    }
        //        //}

        //        // Insertion d'une table dans un bookmark
        //        // Propriété du Tableau
        //        var tableProperty = new TablePropertiesModel()
        //        {
        //            TableBorders = new TableBorders()()sModel()
        //            {
        //                TopBorder = new TableBorders() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds },
        //                LeftBorder = new TableBorders() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
        //                RightBorder = new TableBorders() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
        //                BottomBorder = new TableBorders() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds }
        //            },
        //            TableWidth = new TableWidthModel()
        //            {
        //                Width = "4900",
        //                Type = TableWidthUnitValues.Pct
        //            },
        //            Layout = new TableLayoutModel() { Type = TableLayoutValues.Fixed }
        //        };
        //        // Lignes du premier tableau pour les constats checked
        //        var lines = new List<TableRow>();

        //        for (int i = 0; i < 3; i++)
        //        {
        //            var borderTopIsOK = new TableBorders();
        //            if (i != 0)
        //                borderTopIsOK.BorderValue = BorderValues.Nil;

        //            // Première ligne
        //            var text = word.CreateRunForText("Header Number : " + i,
        //                    new RunProperties()
        //                    {
        //                        Bold = true,
        //                        FontSize = "24",
        //                        RunFonts = new RunFontsModel()
        //                        {
        //                            Ascii = "Courier New",
        //                            HighAnsi = "Courier New",
        //                            EastAsia = "Courier New",
        //                            ComplexScript = "Courier New"
        //                        }
        //                    });

        //            var cellules = new List<TableCell>()
        //            {
        //                word.CreateTableCell(text, new TableCellProperties() {
        //                    Gridspan = new GridSpan() { Val = 2 },
        //                    Shading = new ShadingModel()
        //                    {
        //                        Fill = "F7941F"
        //                    },
        //                    TableCellWidth = new TableCellWidth()
        //                    {
        //                        Width = "8862"
        //                    },
        //                    TableCellBorders = new TableCellBorders()
        //                    {
        //                        TopBorder = borderTopIsOK
        //                    },
        //                    TableVerticalAlignementValues = TableVerticalAlignmentValues.Center
        //                }),
        //                word.CreateTableCell(word.CreateRun(), new TableCellProperties() {
        //                            TableCellWidth = new TableCellWidth()
        //                            {
        //                                Width = "246"
        //                            },
        //                            Shading = new ShadingModel()
        //                            {
        //                                Fill = "F7941F"
        //                            },
        //                            TableCellBorders = new TableCellBorders() {
        //                                        TopBorder = borderTopIsOK
        //                            }
        //                })
        //            };
        //            lines.Add(word.CreateTableRow(cellules, new TableRowPropertiesModel()
        //            {
        //                TableRowHeight = new TableRowHeightModel()
        //                {
        //                    Val = 380
        //                }
        //            }));

        //            //// Deuxième ligne
        //            //text = word.CreateRunForText("Comments", new RunProperties() { Bold = true });
        //            //cellules = new List<TableCell>()
        //            //{
        //            //    word.CreateTableCell(word.CreateImage(@"c:\temp\Tulips.jpg", new Drawing.Pictures.Model.PictureModel() {
        //            //        ImagePartType   = Packaging.ImagePartType.Jpeg,
        //            //        MaxHeight = 10,
        //            //        MaxWidth = 500
        //            //    }), new TableCellProperties() {
        //            //                TableCellWidth = new TableCellWidth()
        //            //                {
        //            //                    Width = "4890"
        //            //                }
        //            //    }),
        //            //    word.CreateTableMergeCell(word.CreateRun(), new TableCellProperties() {
        //            //                Fusion = true,
        //            //                TableCellWidth = new TableCellWidth()
        //            //                {
        //            //                    Width = "4218"
        //            //                },
        //            //                Gridspan = new GridSpan() { Val = 2 }
        //            //    })
        //            //};
        //            //lines.Add(word.CreateTableRow(cellules));

        //            // Troisième ligne
        //            text = word.CreateRunForText("Texte du Constat Number : " + i, new RunProperties());
        //            cellules = new List<TableCell>()
        //            {
        //                word.CreateTableCell(text, new TableCellProperties() {
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4890"
        //                                        },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            BottomBorder = new TableBorders()()s() {
        //                                                Color = "FF0019"
        //                                                }
        //                                        }
        //                }),
        //                word.CreateTableMergeCell(word.CreateRun(), new TableCellProperties() {
        //                                        Fusion = true,
        //                                        FusionChild = true,
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4218"
        //                                        },
        //                                        Gridspan = new GridSpan() { Val = 2 },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            BottomBorder = new TableBorders() {
        //                                                Color = "FF0019" }
        //                                            }
        //                })
        //            };
        //            lines.Add(word.CreateTableRow(cellules));

        //            // Quatrième ligne
        //            text = word.CreateRunForText("Risques", new RunProperties() { Bold = true });
        //            cellules = new List<TableCell>()
        //            {
        //                word.CreateTableCell(text, new TableCellProperties() {
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4890"
        //                                        },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            TopBorder = new TableBorders()
        //                                            {
        //                                                Color = "00FF19"
        //                                            }
        //                                        }
        //                }),
        //                word.CreateTableMergeCell(word.CreateRun(), new TableCellProperties() {
        //                                        Fusion = true,
        //                                        FusionChild = true,
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4218"
        //                                        },
        //                                        Gridspan = new GridSpan() { Val = 2 },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            TopBorder = new TableBorders()
        //                                            {
        //                                                Color = "00FF19"
        //                                            }
        //                                        }
        //                })
        //            };
        //            lines.Add(word.CreateTableRow(cellules));

        //            // Cinquième ligne
        //            text = word.CreateRunForText("Texte du Risque Number : " + i, new RunProperties());
        //            cellules = new List<TableCell>()
        //            {
        //                word.CreateTableCell(text, new TableCellProperties() {
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4890"
        //                                        },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            TopBorder = new TableBorders() {
        //                                                BorderValue = BorderValues.Nil }
        //                                        }
        //                }),
        //                word.CreateTableMergeCell(word.CreateRun(), new TableCellProperties() {
        //                                        Fusion = true,
        //                                        FusionChild = true,
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4218"
        //                                        },
        //                                        Gridspan = new GridSpan() { Val = 2 },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            TopBorder = new TableBorders() {
        //                                                BorderValue = BorderValues.Nil }
        //                                            }
        //                })
        //            };
        //            lines.Add(word.CreateTableRow(cellules));

        //            // Sixième ligne
        //            text = word.CreateRunForText("Recommandations", new RunProperties() { Bold = true });
        //            cellules = new List<TableCell>()
        //            {
        //                word.CreateTableCell(text, new TableCellProperties() {
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4890"
        //                                        },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            BottomBorder = new TableBorders() {
        //                                                BorderValue = BorderValues.Nil }
        //                                            }
        //                }),
        //                word.CreateTableMergeCell(word.CreateRun(), new TableCellProperties() {
        //                                        Fusion = true,
        //                                        FusionChild = true,
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4218"
        //                                        },
        //                                        Gridspan = new GridSpan() { Val = 2 },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            BottomBorder = new TableBorders() {
        //                                                BorderValue = BorderValues.Nil }
        //                                            }
        //                })
        //            };
        //            lines.Add(word.CreateTableRow(cellules));

        //            var borderBottomIsOK = new TableBorders()
        //            {
        //                BorderValue = BorderValues.Nil,
        //                Color = "FF0019"
        //            };

        //            if (i == 2)
        //                borderBottomIsOK.BorderValue = BorderValues.Single;

        //            // Septième ligne
        //            text = word.CreateRunForText("Texte de la Recommandation Number : " + i, new RunProperties());
        //            cellules = new List<TableCell>()
        //            {
        //                word.CreateTableCell(text, new TableCellProperties() {
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4890"
        //                                        },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            TopBorder = new TableBorders() { BorderValue = BorderValues.Nil },
        //                                            BottomBorder = borderBottomIsOK },
        //                                        TableVerticalAlignementValues = TableVerticalAlignmentValues.Center
        //                }),
        //                word.CreateTableMergeCell(word.CreateRun(), new TableCellProperties() {
        //                                        Fusion = true,
        //                                        FusionChild = true,
        //                                        TableCellWidth = new TableCellWidth()
        //                                        {
        //                                            Width = "4218"
        //                                        },
        //                                        Gridspan = new GridSpan() { Val = 2 },
        //                                        TableCellBorders = new TableCellBorders() {
        //                                            TopBorder = new TableBorders() {
        //                                                BorderValue = BorderValues.Nil },
        //                                            BottomBorder = borderBottomIsOK }
        //                })
        //            };
        //            lines.Add(word.CreateTableRow(cellules));
        //        }

        //        IList<Paragraph> tables = new List<Paragraph>();
        //        tables.Add(word.CreateParagraphForRun(word.CreateRunForTable(word.CreateTable(lines, tableProperty))));

        //        tables.Add(word.CreateParagraphForRun(
        //                word.CreateRunForText("Test de style avec bordures de paragraph"),
        //                new ParagraphProperties()
        //                {
        //                    ParagraphStyleId = new ParagraphStyleId()
        //                    {
        //                        Val = "Titre01"
        //                    },
        //                    SpacingBetweenLines = new SpacingBetweenLines()
        //                    {
        //                        After = "800",
        //                        Before = "100"
        //                    },
        //                    ParagraphBorders = new ParagraphBorders()
        //                    {
        //                        TopBorder = new ParagraphBorder() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds },
        //                        LeftBorder = new ParagraphBorder() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
        //                        RightBorder = new ParagraphBorder() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
        //                        BottomBorder = new ParagraphBorder() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds }
        //                    }
        //                }));

        //        // Lignes du deuxième tableau pour les constats unchecked
        //        //lines = new List<TableRow>();

        //        word.SetParagraphsOnBookmark("Insert_Documents", tables);

        //        IList<Paragraph> cell = new List<Paragraph>();
        //        Paragraph productPathParagraph = word.CreateParagraphForRun(word.CreateRun());

        //        // add asset or location name
        //        productPathParagraph.Append(word.CreateRunForText("Txt 1", new RunProperties()
        //        {
        //            FontSize = "22",
        //            Color = "FF0000",
        //            RunFonts = new RunFontsModel()
        //            {
        //                Ascii = "Arial Rounded Light Roman"
        //            }
        //        }));

        //        productPathParagraph.Append(word.CreateRunForText(" / ", new RunProperties()
        //        {
        //            FontSize = "22",
        //            Color = "0000FF",
        //            RunFonts = new RunFontsModel()
        //            {
        //                Ascii = "Arial Rounded Light Roman"
        //            }
        //        }));

        //        productPathParagraph.Append(word.CreateRunForText(" Text 2 ", new RunProperties()
        //        {
        //            FontSize = "22",
        //            Color = "00FF00",
        //            RunFonts = new RunFontsModel()
        //            {
        //                Ascii = "Arial Rounded Light Roman"
        //            }
        //        }));
        //        cell.Add(productPathParagraph);

        //        TableCell tableCells = word.CreateTableCell(cell, new TableCellProperties()
        //        {
        //            TableVerticalAlignement = TableVerticalAlignmentValues.Center,
        //            TableCellWidth = new TableCellWidth()
        //            {
        //                Width = "800",
        //                Type = TableWidthUnitValues.Pct,
        //            },
        //            TableCellMargin = new TableCellMargin()
        //            {
        //                TopMargin = new TableWidth() { Width = "1500", Type = TableWidthUnitValues.Dxa }
        //            }
        //        });

        //        Table table = word.CreateTable(new List<TableRow>() { word.CreateTableRow(new List<TableCell>() { tableCells }) });
        //        word.SetOnBookmark("Table", word.CreateRunForTable(table));

        //        Run run = new Run();
        //        run.Append(word.CreateParagraphForRun(word.CreateRunForText("Paragraph in the shell")));
        //        word.SetOnBookmark("ParagraphInCell", run);

        //        IList<Paragraph> paragraphs = new List<Paragraph>();
        //        paragraphs.Add(word.CreateParagraphForRun(word.CreateRunForText("Text 1")));
        //        paragraphs.Add(word.CreateParagraphForRun(word.CreateRunForText("Line 2")));
        //        paragraphs.Add(word.CreateParagraphForRun(word.CreateRunForText("Working ?")));
        //        word.SetParagraphsOnBookmark("ParagraphInCell2", paragraphs);

        //        List<string> texts = new List<string>();
        //        texts.Add("first line");
        //        texts.Add("second line");
        //        texts.Add("third line");
        //        word.SetTextsOnBookmark("FormatedText", texts, true);
        //        word.SetTextsOnBookmark("UnFormatedText", texts, false);
        //        word.SaveDoc();
        //        word.CloseDoc();
        //    }

        //    Process.Start(finalFilePath);
        //}
    }
}