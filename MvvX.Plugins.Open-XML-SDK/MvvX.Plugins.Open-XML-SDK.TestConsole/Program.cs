using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using MvvX.Plugins.OpenXMLSDK.Platform.Validation;
using MvvX.Plugins.OpenXMLSDK.Platform.Word;
using MvvX.Plugins.OpenXMLSDK.Word;
using MvvX.Plugins.OpenXMLSDK.Word.Models;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs.Models;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels.Charts;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Charts;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;
using MvvX.Plugins.OpenXMLSDK.Word.Tables.Models;
using Newtonsoft.Json;

namespace MvvX.Plugins.OpenXMLSDK.TestConsole
{
    class Program
    {
        static void Main()
        {
            ReportEngineTest();

            //ValidateDocument();
            // fin test report engine

            //OldProgram();
        }

        private static void ReportEngineTest()
        {
            Console.WriteLine("Enter the path of your Json file, press enter for an example");
            var filePath = Console.ReadLine();
            var documentName = string.Empty;
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("Enter document name");
                documentName = Console.ReadLine();
            }

            Console.WriteLine("Generation in progress");
            ReportEngine(filePath, documentName);
        }

        private static void ValidateDocument()
        {
            // Debut test report engine
            var validator = new OpenXMLValidator();
            var results = validator.ValidateWordDocument("ExampleDocument.docx");
            File.Delete("ValidateDocument.txt");
            foreach (var result in results)
            {
                File.AppendAllText("ValidateDocument.txt",
                                    result.XmlPath + Environment.NewLine +
                                    result.Description + Environment.NewLine +
                                    result.ErrorType.ToString() + Environment.NewLine +
                                    " - - - - - - - - - ");
            }
        }

        private static void ReportEngine(string filePath, string documentName)
        {
            // Debut test report engine
            using (IWordManager word = new WordManager())
            {
                JsonConverter[] converters = { new JsonContextConverter() };

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    if (string.IsNullOrWhiteSpace(documentName))
                        documentName = "ExampleDocument.docx";

                    var template = GetTemplateDocument();
                    var templateJson = JsonConvert.SerializeObject(template);
                    var templateUnserialized = JsonConvert.DeserializeObject<Document>(templateJson, new JsonSerializerSettings() { Converters = converters });

                    var context = GetContext();
                    var contextJson = JsonConvert.SerializeObject(context);
                    var contextUnserialized = JsonConvert.DeserializeObject<ContextModel>(contextJson, new JsonSerializerSettings() { Converters = converters });

                    var res = word.GenerateReport(templateUnserialized, contextUnserialized, new CultureInfo("en-US"));

                    // test ecriture fichier
                    File.WriteAllBytes(documentName, res);

                    Process.Start(documentName);
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(documentName))
                        documentName = "ExampleDocument.docx";
                    if (!documentName.EndsWith(".docx"))
                        documentName = string.Concat(documentName, ".docx");

                    var stream = File.ReadAllText(filePath);
                    var report = JsonConvert.DeserializeObject<Report>(stream, new JsonSerializerSettings() { Converters = converters });

                    var res = word.GenerateReport(report.Document, report.ContextModel, new CultureInfo("en-US"));

                    // test ecriture fichier
                    File.WriteAllBytes(documentName, res);
                    Process.Start(documentName);
                }
            }
        }

        /// <summary>
        /// Generate the template
        /// Please add a new method for each new test page
        /// </summary>
        /// <returns></returns>
        private static Document GetTemplateDocument()
        {
            var doc = new Document();

            GenerateStyles(doc);
            GenerateHeaderAndFooter(doc);

            GenerateLotsOfThings(doc);
            GenerateParagraphsWithStyle(doc);
            GenerateForeachPage(doc);
            GenerateTextWithHeritage(doc);
            GenerateTableOfContent(doc);
            GenerateUniformGrid(doc);
            GenerateTable(doc);
            GenerateGraph(doc);
            GenerateTableWithFusionedCells(doc);
            GenerateTextWithEmptyLine(doc);

            return doc;
        }

        /// <summary>
        /// Generate lots of things
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateLotsOfThings(Document doc)
        {
            var page1 = new Page();
            page1.Margin = new Word.ReportEngine.Models.Attributes.SpacingModel() { Top = 845, Bottom = 1418, Left = 567, Right = 567, Header = 709, Footer = 709 };

            var paragraph = new Paragraph();
            paragraph.ChildElements.Add(new Label() { Text = "Ceci est un texte avec accents (éèàù)", FontSize = "30", FontName = "Arial" });
            paragraph.ChildElements.Add(new Label() { Text = "#KeyTest1#", FontSize = "40", FontColor = "#FontColorTestRed#", Shading = "9999FF", BoldKey = "#BoldKey#", Bold = false });
            paragraph.ChildElements.Add(new Label() { Text = "#KeyTest2#", Show = false });

            paragraph.ChildElements.Add(new Label() { Text = "Double value : #KeyTestDouble1#" });
            paragraph.ChildElements.Add(new Label() { Text = "Double value 2 : #KeyTestDouble2#" });

            paragraph.ChildElements.Add(new Label() { Text = "DateTime value : #KeyTestDatetime1#" });
            paragraph.ChildElements.Add(new Label() { Text = "Double value 2 : #KeyTestDatetime2#" });

            paragraph.Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
            {
                BorderPositions = Word.ReportEngine.Models.Attributes.BorderPositions.BOTTOM |
                                        Word.ReportEngine.Models.Attributes.BorderPositions.TOP |
                                        Word.ReportEngine.Models.Attributes.BorderPositions.LEFT,
                BorderWidthBottom = 3,
                BorderWidthLeft = 10,
                BorderWidthTop = 20,
                BorderWidthInsideVertical = 1,
                UseVariableBorders = true,
                BorderColor = "FF0000",
                BorderLeftColor = "CCCCCC",
                BorderTopColor = "123456",
                BorderRightColor = "FFEEDD",
                BorderBottomColor = "FF1234"
            };
            page1.ChildElements.Add(paragraph);
            var p2 = new Paragraph();
            p2.Shading = "#ParagraphShading#";
            p2.ChildElements.Add(new Label() { Text = "   texte paragraph2 avec espace avant", FontSize = "20", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve });
            p2.ChildElements.Add(new Label() { Text = "texte2 paragraph2 avec espace après   ", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve });
            p2.ChildElements.Add(new Label() { Text = "   texte3 paragraph2 avec espace avant et après   ", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve });
            page1.ChildElements.Add(p2);

            var table = new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                TableIndentation = new TableIndentation() { Width = 1000 },
                Rows = new List<Row>()
                {
                    new Row()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell()
                            {
                                VerticalAlignment = TableVerticalAlignmentValues.Center,
                                Justification = JustificationValues.Center,
                                ChildElements = new List<BaseElement>()
                                {
                                    new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Cell 1 - First paragraph" } }, ParagraphStyleId = "Yellow" },
                                    new Image()
                                    {
                                        Width = 50,
                                        Path = @"..\..\Resources\Desert.jpg",
                                        ImagePartType = Packaging.ImagePartType.Jpeg
                                    },
                                    new Label() { Text = "Cell 1 - Label in a cell" },
                                    new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Cell 1 - Second paragraph" } } }
                                },
                                Fusion = true
                            },
                            new Cell()
                            {
                                ChildElements = new List<BaseElement>()
                                {
                                    new Label() { Text = "Cell 2 - First label" },
                                    new Image()
                                    {
                                        Height = 10,
                                        Path = @"..\..\Resources\Desert.jpg",
                                        ImagePartType = Packaging.ImagePartType.Jpeg
                                    },
                                    new Label() { Text = "Cell 2 - Second label" }
                                },
                                Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
                                {
                                    BorderColor = "#BorderColor#",
                                    BorderWidth = 20,
                                    BorderPositions = Word.ReportEngine.Models.Attributes.BorderPositions.LEFT | Word.ReportEngine.Models.Attributes.BorderPositions.TOP
                                }
                            }
                        }
                    },
                    new Row()
                    {
                        ShowKey = "#NoRow#",
                        Cells = new List<Cell>()
                        {
                            new Cell()
                            {
                                Fusion = true,
                                FusionChild = true
                            },
                            new Cell()
                            {
                                VerticalAlignment = TableVerticalAlignmentValues.Bottom,
                                Justification = JustificationValues.Right,
                                ChildElements = new List<BaseElement>()
                                {
                                    new Label() { Text = "cellule4" }
                                }
                            }
                        }
                    }
                }
            };

            table.HeaderRow = new Row()
            {
                Cells = new List<Cell>()
                {
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "header1" } } }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "header2" }
                            }
                        }
                }
            };

            table.Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
            {
                BorderPositions = Word.ReportEngine.Models.Attributes.BorderPositions.BOTTOM | Word.ReportEngine.Models.Attributes.BorderPositions.INSIDEVERTICAL,
                BorderWidthBottom = 50,
                BorderWidthInsideVertical = 1,
                UseVariableBorders = true,
                BorderColor = "FF0000"
            };

            page1.ChildElements.Add(table);
            page1.ChildElements.Add(new Paragraph());

            if (File.Exists(@"..\..\Resources\Desert.jpg"))
                page1.ChildElements.Add(
                    new Paragraph()
                    {
                        ChildElements = new List<BaseElement>()
                        {
                        new Image()
                        {
                            MaxHeight = 100,
                            MaxWidth = 100,
                            Path = @"..\..\Resources\Desert.jpg",
                            ImagePartType = Packaging.ImagePartType.Jpeg
                        }
                        }
                    }
                );

            var tableDataSource = new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                ColsWidth = new int[2] { 750, 4250 },
                Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
                {
                    BorderPositions = (Word.ReportEngine.Models.Attributes.BorderPositions)63,
                    BorderColor = "328864",
                    BorderWidth = 20,
                },
                RowModel = new Row()
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            Shading = "FFA0FF",
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "#Cell1#" }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "#Cell2#" }
                            }
                        }
                    }
                },
                DataSourceKey = "#Datasource#"
            };

            var tableDataSourceWithPrefix = new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                ColsWidth = new int[2] { 750, 4250 },
                Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
                {
                    BorderPositions = (Word.ReportEngine.Models.Attributes.BorderPositions)63,
                    BorderColor = "328864",
                    BorderWidth = 20,
                },
                DataSourceKey = "#DatasourcePrefix#",
                AutoContextAddItemsPrefix = "DataSourcePrefix",
                RowModel = new Row()
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            Shading = "FFA0FF",
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "Item Datasource (0 index) #DataSourcePrefix_TableRow_IndexBaseZero# - ",
                                                ShowKey = "#DataSourcePrefix_TableRow_IsFirstItem#" },
                                new Label() { Text = "#Cell1#" }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "Item Datasource (1 index) #DataSourcePrefix_TableRow_IndexBaseOne# - ",
                                                ShowKey = "#DataSourcePrefix_TableRow_IsLastItem#" },
                                new Label() { Text = "#Cell2#" }
                            }
                        }
                    }
                }
            };

            page1.ChildElements.Add(tableDataSource);

            page1.ChildElements.Add(tableDataSourceWithPrefix);

            doc.Pages.Add(page1);
        }

        /// <summary>
        /// Generate paragraph with styles
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateParagraphsWithStyle(Document doc)
        {
            var page = new Page
            {
                Margin = new Word.ReportEngine.Models.Attributes.SpacingModel() { Top = 1418, Left = 845, Header = 709, Footer = 709 }
            };

            // Red Style
            var redParagrahp = new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red"
            };
            redParagrahp.ChildElements.Add(new Label() { Text = "texte page2", FontName = "Arial" });
            page.ChildElements.Add(redParagrahp);

            // Yellow Style
            var yellowParagraph = new Paragraph
            {
                SpacingBefore = 800,
                SpacingAfter = 800,
                Justification = JustificationValues.Both,
                ParagraphStyleId = "Yellow"
            };
            yellowParagraph.ChildElements.Add(new Label() { Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse urna augue, convallis eu enim vitae, maximus ultrices nulla. Sed egestas volutpat luctus. Maecenas sodales erat eu elit auctor, eu mattis neque maximus. Duis ac risus quis sem bibendum efficitur. Vivamus justo augue, molestie quis orci non, maximus imperdiet justo. Donec condimentum rhoncus est, ut varius lorem efficitur sed. Donec accumsan sit amet nisl vel ornare. Duis aliquet urna eu mauris porttitor facilisis. " });
            page.ChildElements.Add(yellowParagraph);

            // Parahraph with border
            var borderParagraph = new Paragraph();
            borderParagraph.Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
            {
                BorderPositions = (Word.ReportEngine.Models.Attributes.BorderPositions)13,
                BorderWidth = 20,
                BorderColor = "#ParagraphBorderColor#"
            };
            borderParagraph.SpacingBetweenLines = 360;
            borderParagraph.ChildElements.Add(new Label() { Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse urna augue, convallis eu enim vitae, maximus ultrices nulla. Sed egestas volutpat luctus. Maecenas sodales erat eu elit auctor, eu mattis neque maximus. Duis ac risus quis sem bibendum efficitur. Vivamus justo augue, molestie quis orci non, maximus imperdiet justo. Donec condimentum rhoncus est, ut varius lorem efficitur sed. Donec accumsan sit amet nisl vel ornare. Duis aliquet urna eu mauris porttitor facilisis. " });
            page.ChildElements.Add(borderParagraph);

            doc.Pages.Add(page);
        }

        /// <summary>
        /// Generate a foreach page
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateForeachPage(Document doc)
        {
            // Adding a foreach page :
            var foreachPage = new ForEachPage();
            foreachPage.DataSourceKey = "#DatasourceTableFusion#";

            foreachPage.Margin = new Word.ReportEngine.Models.Attributes.SpacingModel() { Top = 1418, Left = 845, Header = 709, Footer = 709 };
            var paragraph21 = new Paragraph();
            paragraph21.ChildElements.Add(new Label() { Text = "Page label : #Label#" });
            foreachPage.ChildElements.Add(paragraph21);
            var p223 = new Paragraph();
            p223.Shading = "#ParagraphShading#";
            p223.ChildElements.Add(new Label() { Text = "Texte paragraph2 avec espace avant", FontSize = "20", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve });
            foreachPage.ChildElements.Add(p223);
            doc.Pages.Add(foreachPage);
        }

        /// <summary>
        /// Generate text with different styles
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateTextWithHeritage(Document doc)
        {
            var page = new Page();

            var p31 = new Paragraph() { FontColor = "FF0000", FontSize = "26" };
            p31.ChildElements.Add(new Label() { Text = "Test the HeritFromParent" });
            var p311 = new Paragraph() { FontSize = "16" };
            p311.ChildElements.Add(new Label() { Text = " Success (not the same size)" });
            p31.ChildElements.Add(p311);
            page.ChildElements.Add(p31);

            doc.Pages.Add(page);
        }

        /// <summary>
        /// Generate a table of content
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateTableOfContent(Document doc)
        {
            var page = new Page();

            TableOfContents tableOfContents = new TableOfContents()
            {
                StylesAndLevels = new List<Tuple<string, string>>()
                {
                    new Tuple<string, string>("Red", "1"),
                    new Tuple<string, string>("Yellow", "2"),
                },
                Title = "Tessssssst !",
                TitleStyleId = "Yellow",
                ToCStylesId = new List<string>() { "Red" },
                LeaderCharValue = TabStopLeaderCharValues.underscore
            };
            page.ChildElements.Add(tableOfContents);

            doc.Pages.Add(page);
        }

        /// <summary>
        /// Generate an UniformGrid
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateUniformGrid(Document doc)
        {
            var page = new Page();
            //New page to manage UniformGrid:
            var uniformGrid = new UniformGrid()
            {
                DataSourceKey = "#UniformGridSample#",
                ColsWidth = new int[2] { 2500, 2500 },
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                CellModel = new Cell()
                {
                    VerticalAlignment = TableVerticalAlignmentValues.Center,
                    Justification = JustificationValues.Center,
                    ChildElements = new List<BaseElement>()
                        {
                            new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "#CellUniformGridTitle#" } } },
                            new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Cell 1 - Second paragraph" } } }
                        }
                },
                HeaderRow = new Row()
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "header1" } } }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "header2" }
                            }
                        }
                    }
                },
                Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
                {
                    BorderPositions = Word.ReportEngine.Models.Attributes.BorderPositions.BOTTOM | Word.ReportEngine.Models.Attributes.BorderPositions.INSIDEVERTICAL,
                    BorderWidthBottom = 50,
                    BorderWidthInsideVertical = 1,
                    UseVariableBorders = true,
                    BorderColor = "FF0000"
                }
            };

            page.ChildElements.Add(uniformGrid);

            doc.Pages.Add(page);
        }

        /// <summary>
        /// Generate a table
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateTable(Document doc)
        {
            var page = new Page();
            var tableDataSourceWithBeforeAfter = new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                ColsWidth = new int[2] { 750, 4250 },
                Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
                {
                    BorderPositions = (Word.ReportEngine.Models.Attributes.BorderPositions)63,
                    BorderColor = "328864",
                    BorderWidth = 20,
                },
                BeforeRows = new List<Row>()
                {
                    new Row()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell()
                            {
                                VerticalAlignment = TableVerticalAlignmentValues.Bottom,
                                Justification = JustificationValues.Left,
                                ChildElements = new List<BaseElement>()
                                {
                                    new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Cell 1 - A small paragraph" } }, ParagraphStyleId = "Yellow" },
                                    new Image()
                                    {
                                        MaxHeight = 100,
                                        MaxWidth = 100,
                                        Path = @"..\..\Resources\Desert.jpg",
                                        ImagePartType = Packaging.ImagePartType.Jpeg
                                    },
                                    new Label() { Text = "Custom header" },
                                    new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Cell 1 - an other paragraph" } } }
                                },
                                Fusion = true
                            },
                            new Cell()
                            {
                                ChildElements = new List<BaseElement>()
                                {
                                    new Label() { Text = "Cell 2 - an other label" },
                                    new Image()
                                    {
                                        MaxHeight = 100,
                                        MaxWidth = 100,
                                        Path = @"..\..\Resources\Desert.jpg",
                                        ImagePartType = Packaging.ImagePartType.Jpeg
                                    },
                                    new Label() { Text = "Cell 2 - an other other label" }
                                },
                                Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
                                {
                                    BorderColor = "00FF22",
                                    BorderWidth = 15,
                                    BorderPositions = Word.ReportEngine.Models.Attributes.BorderPositions.RIGHT | Word.ReportEngine.Models.Attributes.BorderPositions.TOP
                                }
                            }
                        }
                    },
                    new Row()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell()
                            {
                                Fusion = true,
                                FusionChild = true
                            },
                            new Cell()
                            {
                                VerticalAlignment = TableVerticalAlignmentValues.Bottom,
                                Justification = JustificationValues.Right,
                                ChildElements = new List<BaseElement>()
                                {
                                    new Label() { Text = "celluleX" }
                                }
                            }
                        }
                    }
                },
                RowModel = new Row()
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            Shading = "FFA2FF",
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "Cell : #Cell1#" }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "Cell : #Cell2#" }
                            }
                        }
                    }
                },
                DataSourceKey = "#Datasource#"
            };

            page.ChildElements.Add(tableDataSourceWithBeforeAfter);

            doc.Pages.Add(page);
        }

        /// <summary>
        /// Generate a graph
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateGraph(Document doc)
        {
            var page = new Page();

            var paragraph = new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new Word.ReportEngine.Models.Charts.BarModel()
                    {
                        Title = "Graph test",
                        ShowTitle = true,
                        FontSize = 23,
                        ShowBarBorder = true,
                        BarChartType = BarChartType.BarChart,
                        BarDirectionValues = BarDirectionValues.Column,
                        BarGroupingValues = BarGroupingValues.PercentStacked,
                        DataSourceKey = "#GrahSampleData#",
                        ShowMajorGridlines = true
                    }
                }
            };

            page.ChildElements.Add(paragraph);

            doc.Pages.Add(page);
        }

        /// <summary>
        /// Generate a table with fusionned cells
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateTableWithFusionedCells(Document doc)
        {
            var page = new Page();

            var tableDataSourceWithCellFusion = new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                ColsWidth = new int[3] { 1200, 1200, 1200 },
                Borders = new Word.ReportEngine.Models.Attributes.BorderModel()
                {
                    BorderPositions = (Word.ReportEngine.Models.Attributes.BorderPositions)63,
                    BorderColor = "328864",
                    BorderWidth = 20,
                },
                RowModel = new Row()
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            Shading = "FFA0FF",
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "#Cell1#" }
                            },
                            FusionKey = "#IsInGroup#",
                            FusionChildKey = "#IsNotFirstLineGroup#"
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "#Cell2#" }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "#Cell2#" }
                            }
                        }
                    }
                },
                DataSourceKey = "#DatasourceTableFusion#"
            };

            page.ChildElements.Add(tableDataSourceWithCellFusion);

            doc.Pages.Add(page);
        }

        /// <summary>
        /// Generate a paragraph with an empty line
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateTextWithEmptyLine(Document doc)
        {
            var page = new Page();
            var paragraph = new Paragraph() { FontColor = "FF0000", FontSize = "26" };
            paragraph.ChildElements.Add(new Label() { Text = "Label with" + Environment.NewLine + Environment.NewLine + "A new line" });
            page.ChildElements.Add(paragraph);

            doc.Pages.Add(page);
        }

        /// <summary>
        /// Generate Styles
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateStyles(Document doc)
        {
            doc.Styles.Add(new Style() { StyleId = "Red", FontColor = "FF0050", FontSize = "42", PrimaryStyle = true });
            doc.Styles.Add(new Style() { StyleId = "Yellow", FontColor = "FFFF00", FontSize = "40" });
        }

        /// <summary>
        /// Generate Header(s) and Footer(s)
        /// </summary>
        /// <param name="doc"></param>
        private static void GenerateHeaderAndFooter(Document doc)
        {
            // Header
            var header = new Header();
            header.Type = HeaderFooterValues.Default;
            var ph = new Paragraph();
            ph.ChildElements.Add(new Label() { Text = "Header Text" });
            if (File.Exists(@"..\..\Resources\Desert.jpg"))
                ph.ChildElements.Add(new Image()
                {
                    MaxHeight = 100,
                    MaxWidth = 100,
                    Path = @"..\..\Resources\Desert.jpg",
                    ImagePartType = Packaging.ImagePartType.Jpeg
                });
            header.ChildElements.Add(ph);
            doc.Headers.Add(header);

            // first header
            var firstHeader = new Header();
            firstHeader.Type = HeaderFooterValues.First;
            var fph = new Paragraph();
            fph.ChildElements.Add(new Label() { Text = "first header Text" });
            firstHeader.ChildElements.Add(fph);
            doc.Headers.Add(firstHeader);

            // Footer
            var footer = new Footer();
            footer.Type = HeaderFooterValues.Default;
            var pf = new Paragraph();
            pf.ChildElements.Add(new Label() { Text = "Footer Text" });
            pf.ChildElements.Add(new Label() { IsPageNumber = true });
            footer.ChildElements.Add(pf);
            doc.Footers.Add(footer);
        }

        /// <summary>
        /// Generate the context for the generated template
        /// </summary>
        /// <returns></returns>
        private static ContextModel GetContext()
        {
            ContextModel context = new ContextModel();
            context.AddItem("#NoRow#", new BooleanModel(false));
            context.AddItem("#ParagraphShading#", new StringModel("00FF00"));
            context.AddItem("#ParagraphBorderColor#", new StringModel("105296"));
            context.AddItem("#BorderColor#", new StringModel("00FF00"));
            context.AddItem("#KeyTest1#", new StringModel("Key 1"));
            context.AddItem("#KeyTest2#", new StringModel("Key 2"));

            context.AddItem("#KeyTestDouble1#", new DoubleModel(125.2345, "{0:0.##} kV"));
            context.AddItem("#KeyTestDouble2#", new DoubleModel(1025.2345, "Before - {0:0.###} - After"));

            context.AddItem("#KeyTestDatetime1#", new DateTimeModel(DateTime.Now, null));
            context.AddItem("#KeyTestDatetime2#", new DateTimeModel(DateTime.Now, "d"));

            context.AddItem("#BoldKey#", new BooleanModel(true));

            context.AddItem("#FontColorTestRed#", new StringModel("993333"));
            context.AddItem("#ParagraphStyleIdTestYellow#", new StringModel("Yellow"));

            ContextModel row1 = new ContextModel();
            row1.AddItem("#Cell1#", new StringModel("Col 1 Row 1"));
            row1.AddItem("#Cell2#", new StringModel("Col 2 Row 1"));
            row1.AddItem("#Label#", new StringModel("Label 1"));
            ContextModel row2 = new ContextModel();
            row2.AddItem("#Cell1#", new StringModel("Col 2 Row 1"));
            row2.AddItem("#Cell2#", new StringModel("Col 2 Row 2"));
            row2.AddItem("#Label#", new StringModel("Label 2"));
            ContextModel row3 = new ContextModel();
            row3.AddItem("#Cell1#", new StringModel("Col 1 Row 3"));
            row3.AddItem("#Cell2#", new StringModel("Col 2 Row 3"));
            row3.AddItem("#Label#", new StringModel("Label 1"));
            ContextModel row4 = new ContextModel();
            row4.AddItem("#Cell1#", new StringModel("Col 2 Row 4"));
            row4.AddItem("#Cell2#", new StringModel("Col 2 Row 4"));
            row4.AddItem("#Label#", new StringModel("Label 2"));

            context.AddItem("#Datasource#", new DataSourceModel()
            {
                Items = new List<ContextModel>()
                    {
                        row1, row2
                    }
            });

            context.AddItem("#DatasourcePrefix#", new DataSourceModel()
            {
                Items = new List<ContextModel>()
                    {
                        row1, row2, row3, row4
                    }
            });

            ContextModel row11 = new ContextModel();
            row11.AddItem("#IsInGroup#", new BooleanModel(true));
            row11.AddItem("#IsNotFirstLineGroup#", new BooleanModel(false));
            row11.AddItem("#Cell1#", new StringModel("Col 1 Row 1"));
            row11.AddItem("#Cell2#", new StringModel("Col 2 Row 1"));
            row11.AddItem("#Label#", new StringModel("Label 1"));
            ContextModel row12 = new ContextModel();
            row12.AddItem("#IsInGroup#", new BooleanModel(true));
            row12.AddItem("#IsNotFirstLineGroup#", new BooleanModel(true));
            row12.AddItem("#Cell1#", new StringModel("Col 1 Row 1"));
            row12.AddItem("#Cell2#", new StringModel("Col 2 Row 1"));
            row12.AddItem("#Label#", new StringModel("Label 1"));
            ContextModel row13 = new ContextModel();
            row13.AddItem("#IsInGroup#", new BooleanModel(true));
            row13.AddItem("#IsNotFirstLineGroup#", new BooleanModel(true));
            row13.AddItem("#Cell1#", new StringModel("Col 1 Row 1"));
            row13.AddItem("#Cell2#", new StringModel("Col 2 Row 1"));
            row13.AddItem("#Label#", new StringModel("Label 1"));
            ContextModel row22 = new ContextModel();
            row22.AddItem("#IsInGroup#", new BooleanModel(true));
            row22.AddItem("#IsNotFirstLineGroup#", new BooleanModel(false));
            row22.AddItem("#Cell1#", new StringModel("Col 2 Row 1"));
            row22.AddItem("#Cell2#", new StringModel("Col 2 Row 2"));
            row22.AddItem("#Label#", new StringModel("Label 2"));
            ContextModel row23 = new ContextModel();
            row23.AddItem("#IsInGroup#", new BooleanModel(true));
            row23.AddItem("#IsNotFirstLineGroup#", new BooleanModel(true));
            row23.AddItem("#Cell1#", new StringModel("Col 2 Row 1"));
            row23.AddItem("#Cell2#", new StringModel("Col 2 Row 2"));
            row23.AddItem("#Label#", new StringModel("Label 2"));

            context.AddItem("#DatasourceTableFusion#", new DataSourceModel()
            {
                Items = new List<ContextModel>()
                    {
                        row11, row12, row13, row22, row23
                    }
            });

            List<ContextModel> cellsContext = new List<ContextModel>();
            for (int i = 0; i < DateTime.Now.Day; i++)
            {
                ContextModel uniformGridContext = new ContextModel();
                uniformGridContext.AddItem("#CellUniformGridTitle#", new StringModel("Item number " + (i + 1)));
                cellsContext.Add(uniformGridContext);
            }
            context.AddItem("#UniformGridSample#", new DataSourceModel()
            {
                Items = cellsContext
            });

            context.AddItem("#GrahSampleData#", new BarChartModel()
            {
                BarChartContent = new Word.ReportEngine.BatchModels.Charts.BarModel()
                {
                    Categories = new List<BarCategoryModel>()
                    {
                        new BarCategoryModel()
                        {
                            Name = "Category 1"
                        },
                        new BarCategoryModel()
                        {
                            Name = "Category 2"
                        },
                        new BarCategoryModel()
                        {
                            Name = "Category 3"
                        },
                        new BarCategoryModel()
                        {
                            Name = "Category 4"
                        },
                        new BarCategoryModel()
                        {
                            Name = "Category 5"
                        },
                        new BarCategoryModel()
                        {
                            Name = "Category 6"
                        }
                    },
                    Series = new List<BarSerieModel>()
                    {
                        new BarSerieModel()
                        {
                            Values = new List<double?>()
                            {
                                0, 1, 2, 3, 6, null
                            },
                            Name = "Bar serie 1"
                        },
                        new BarSerieModel()
                        {
                            Values = new List<double?>()
                            {
                                5, null, 7, 8, 0, 10
                            },
                            Name = "Bar serie 2"
                        },
                        new BarSerieModel()
                        {
                            Values = new List<double?>()
                            {
                                9, 10, 11, 12, 13, 14
                            },
                            Name = "Bar serie 3"
                        },
                        new BarSerieModel()
                        {
                            Values = new List<double?>()
                            {
                                9, 10, 11, 12, 15, 25
                            },
                            Name = "Bar serie 4"
                        }
                    }
                }
            });

            return context;
        }

        private static void OldProgram()
        {
            var resourceName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Global.docx");

            if (!Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results")))
                Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results"));

            string finalFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results", "FinalDoc_Test_OrientationParagraph-" + DateTime.Now.ToFileTime() + ".docx");
            using (IWordManager word = new WordManager())
            {
                //TODO for debug : use your test file :
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
                                                    BottomBorder = borderBottomIsOK },
                                                TableVerticalAlignementValues = TableVerticalAlignmentValues.Center
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
                        word.CreateRunForText("Test de style avec bordures de paragraph"),
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
                            },
                            ParagraphBorders = new ParagraphBordersModel()
                            {
                                TopBorder = new ParagraphBorderModel() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds },
                                LeftBorder = new ParagraphBorderModel() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
                                RightBorder = new ParagraphBorderModel() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
                                BottomBorder = new ParagraphBorderModel() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds }
                            }
                        }));

                // Lignes du deuxième tableau pour les constats unchecked
                //lines = new List<TableRow>();

                word.SetParagraphsOnBookmark("Insert_Documents", tables);

                IList<IParagraph> cell = new List<IParagraph>();
                IParagraph productPathParagraph = word.CreateParagraphForRun(word.CreateRun());

                // add asset or location name
                productPathParagraph.Append(word.CreateRunForText("Txt 1", new RunPropertiesModel()
                {
                    FontSize = "22",
                    Color = "FF0000",
                    RunFonts = new RunFontsModel()
                    {
                        Ascii = "Arial Rounded Light Roman"
                    }
                }));

                productPathParagraph.Append(word.CreateRunForText(" / ", new RunPropertiesModel()
                {
                    FontSize = "22",
                    Color = "0000FF",
                    RunFonts = new RunFontsModel()
                    {
                        Ascii = "Arial Rounded Light Roman"
                    }
                }));

                productPathParagraph.Append(word.CreateRunForText(" Text 2 ", new RunPropertiesModel()
                {
                    FontSize = "22",
                    Color = "00FF00",
                    RunFonts = new RunFontsModel()
                    {
                        Ascii = "Arial Rounded Light Roman"
                    }
                }));
                cell.Add(productPathParagraph);

                ITableCell tableCells = word.CreateTableCell(cell, new TableCellPropertiesModel()
                {
                    TableVerticalAlignementValues = TableVerticalAlignmentValues.Center,
                    TableCellWidth = new TableCellWidthModel()
                    {
                        Width = "800",
                        Type = TableWidthUnitValues.Pct,
                    },
                    TableCellMargin = new TableCellMarginModel()
                    {
                        TopMargin = new TableWidthModel() { Width = "1500", Type = TableWidthUnitValues.Dxa }
                    }
                });

                ITable table = word.CreateTable(new List<ITableRow>() { word.CreateTableRow(new List<ITableCell>() { tableCells }) });
                word.SetOnBookmark("Table", word.CreateRunForTable(table));

                IRun run = new PlatformRun();
                run.Append(word.CreateParagraphForRun(word.CreateRunForText("Paragraph in the shell")));
                word.SetOnBookmark("ParagraphInCell", run);

                IList<IParagraph> paragraphs = new List<IParagraph>();
                paragraphs.Add(word.CreateParagraphForRun(word.CreateRunForText("Text 1")));
                paragraphs.Add(word.CreateParagraphForRun(word.CreateRunForText("Line 2")));
                paragraphs.Add(word.CreateParagraphForRun(word.CreateRunForText("Working ?")));
                word.SetParagraphsOnBookmark("ParagraphInCell2", paragraphs);

                List<string> texts = new List<string>();
                texts.Add("first line");
                texts.Add("second line");
                texts.Add("third line");
                word.SetTextsOnBookmark("FormatedText", texts, true);
                word.SetTextsOnBookmark("UnFormatedText", texts, false);
                word.SaveDoc();
                word.CloseDoc();
            }

            Process.Start(finalFilePath);
        }
    }
}