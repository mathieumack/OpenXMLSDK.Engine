using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Newtonsoft.Json;
using OpenXMLSDK.Engine.interfaces.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.ReportEngine.DataContext.Charts;
using OpenXMLSDK.Engine.ReportEngine.DataContext.FluentExtensions;
using OpenXMLSDK.Engine.Word;
using OpenXMLSDK.Engine.Word.Models;
using OpenXMLSDK.Engine.Word.ReportEngine;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.ExtendedModels;
using OpenXMLSDK.Engine.Word.Tables;
using OpenXMLSDK.Engine.Word.Tables.Models;

namespace OpenXMLSDK.UnitTest.ReportEngine
{
    public static class ReportEngineTest
    {
        private const string Lorem_Ipsum = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse urna augue, convallis eu enim vitae, maximus ultrices nulla. Sed egestas volutpat luctus. Maecenas sodales erat eu elit auctor, eu mattis neque maximus. Duis ac risus quis sem bibendum efficitur. Vivamus justo augue, molestie quis orci non, maximus imperdiet justo. Donec condimentum rhoncus est, ut varius lorem efficitur sed. Donec accumsan sit amet nisl vel ornare. Duis aliquet urna eu mauris porttitor facilisis. ";

        public static void ReportEngine(string filePath, string documentName, bool useSeveralReports = false)
        {
            // Debut test report engine
            var serializerSettings = new JsonSerializerSettings() { Converters = { new JsonContextConverter() } };
            IList<Report> reports = default;
            Document reportDocument = default;
            ContextModel reportContext = default;

            // Format document name
            documentName = FormatDocumentName(documentName);

            // Report from test data
            if (string.IsNullOrWhiteSpace(filePath))
            {
                var template = GetTemplateDocument();
                var templateJson = JsonConvert.SerializeObject(template);
                reportDocument = JsonConvert.DeserializeObject<Document>(templateJson, serializerSettings);

                var context = GetContext();
                var contextJson = JsonConvert.SerializeObject(context);
                reportContext = JsonConvert.DeserializeObject<ContextModel>(contextJson, serializerSettings);

                if (useSeveralReports)
                {
                    // Duplicate reports
                    reports = new List<Report>()
                    {
                        new Report()
                        {
                            AddPageBreak = true,
                            ContextModel = reportContext,
                            Document = reportDocument
                        },
                        new Report()
                        {
                            ContextModel = reportContext,
                            Document = reportDocument
                        }
                    };
                }
            }
            // Report from imput data
            else
            {
                var fileContent = File.ReadAllText(filePath);

                if (useSeveralReports)
                {
                    reports = JsonConvert.DeserializeObject<IList<Report>>(fileContent, serializerSettings);
                }
                else
                {
                    var report = JsonConvert.DeserializeObject<Report>(fileContent, serializerSettings);
                    reportDocument = report.Document;
                    reportContext = report.ContextModel;
                }
            }

            // Generate report
            byte[] res;
            using (var word = new WordManager())
            {
                if (useSeveralReports)
                {
                    res = word.GenerateReport(reports, true, new CultureInfo("en-US"));
                }
                else
                {
                    res = word.GenerateReport(reportDocument, reportContext, new CultureInfo("en-US"));
                }
            }

            // Write test file
            File.WriteAllBytes(documentName, res);
        }

        private static string FormatDocumentName(string documentName)
        {
            if (string.IsNullOrWhiteSpace(documentName))
            {
                documentName = "ExampleDocument.docx";
            }

            if (!documentName.EndsWith(".docx"))
            {
                documentName = string.Concat(documentName, ".docx");
            }

            return documentName;
        }

        public static void Test()
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

        /// <summary>
        /// Generate the context for the generated template
        /// </summary>
        /// <returns></returns>
        private static ContextModel GetContext()
        {
            ContextModel context = new ContextModel()
                        .AddBoolean("#NoRow#", false)
                        .AddString("#ParagraphShading#", "00FF00")
                        .AddString("#ParagraphBorderColor#", "105296")
                        .AddString("#BorderColor#", "00FF00")
                        .AddString("#KeyTest1#", "Key 1")
                        .AddString("#KeyTest2#", "Key 2")
                        .AddBoolean("#BoldKey#", true)
                        .AddString("#FontColorTestRed#", "993333")
                        .AddString("#ParagraphStyleIdTestYellow#", "Yellow");

            GenerateForeachContext(context);

            GenerateForeachPageContext(context);

            GenerateUniformGridContext(context);

            GenerateTableContext(context);

            GenerateSubstitutableStringContext(context);

            GeneratePieChartContext(context);

            GenerateBarGraphContext(context);

            GenerateLineGraphContext(context);

            GenerateScatterGraphContext(context);

            GenerateCombineGraphContext(context);

            GenerateMultipleColumnsContext(context);

            return context;
        }

        private static SimpleField PageCrossReference(string anchor)
        {
            return new SimpleField()
            {
                Instruction = anchor
            };
        }

        /// <summary>
        /// Generate the template
        /// </summary>
        /// <returns></returns>
        private static Document GetTemplateDocument()
        {
            var doc = new Document
            {
                Margin = new SpacingModel() { Top = 845, Bottom = 1418, Left = 567, Right = 567, Header = 709, Footer = 709 }
            };
            doc.Styles.Add(new Style() { StyleId = "Red", FontColor = "FF0050", FontSize = "42" });
            doc.Styles.Add(new Style() { StyleId = "Yellow", FontColor = "FFFF00", FontSize = "40" });
            doc.Styles.Add(new Style() { StyleId = "TOC1", FontColor = "8A3459", FontSize = "30", FontName = "Arial" });
            doc.Styles.Add(new Style() { StyleId = "TOC2", FontColor = "8A7934", FontSize = "20", FontName = "Arial" });

            // Paragraphs
            doc.Pages.Add(GenerateParagraphPage());
            // Second page to have different margins
            doc.Pages.Add(GenerateParagraphSecondPage());

            // Foreach
            doc.Pages.Add(GenerateForeachPage(doc));

            // Foreach page
            doc.Pages.Add(GenerateForeachPagePage());

            // Table of content
            doc.Pages.Add(GenerateTableOfContent());

            // Uniform grid
            doc.Pages.Add(GenerateUniformGridPage());

            // Tables
            doc.Pages.Add(GenerateTablesPage());

            // substitutable strings
            doc.Pages.Add(GenerateSubstitutableStringPage());

            // Pie charts
            doc.Pages.Add(GeneratePieChartPage());

            // Bar graphs
            doc.Pages.Add(GenerateBarChartPage());

            // Curve graphs
            doc.Pages.Add(GenerateLineGraphsPage());

            // Scatter graphs
            doc.Pages.Add(GenerateScatterGraphsPage());

            // Combine graphs (Line and Bar)
            doc.Pages.Add(GenerateCombineGraphsPage());

            // Split page on 2 columns
            doc.Pages.Add(GenerateTableOn1stPage());
            doc.Pages.Add(Generate2ColmunOnSamePage());

            // Manage headers and footers
            ManageHeadersAndFooters(doc);

            return doc;
        }

        #region Paragraphs and labels

        /// <summary>
        /// Generate paragraphs page
        /// </summary>
        /// <returns></returns>
        private static Page GenerateParagraphPage()
        {
            var page = new Page();

            // Paragraph with space inside labels and Shading
            page.ChildElements.Add(new Paragraph
            {
                Shading = "#ParagraphShading#",
                ChildElements = new List<BaseElement>
                {
                    new Label() { Text = "   Paragraph with space before", FontSize = "20", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve },
                    new Label() { Text = Environment.NewLine },
                    new Label() { Text = "Paragraph with space after   ", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve },
                    new Label() { Text = Environment.NewLine },
                    new Label() { Text = "   Paragraph2 with space before and after   ", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve }
                }
            });

            // Paragraph with text from context
            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Label()
                    {
                        Text = "#KeyTest1#",
                        FontSize = "40",
                        TransformOperations = new List<LabelTransformOperation>()
                        {
                            new LabelTransformOperation()
                            {
                                TransformOperationType = LabelTransformOperationType.ToUpper
                            }
                        },
                        FontColor = "#FontColorTestRed#",
                        Shading = "9999FF",
                        BoldKey = "#BoldKey#",
                        Bold = false
                    },
                    // This label will not be displayed
                    new Label()
                    {
                        Text = "#KeyTest2#",
                        Show = false
                    }
                }
            });

            // Paragraph with style heritage
            page.ChildElements.Add(new Paragraph
            {
                FontColor = "FF0000",
                FontSize = "26",
                ChildElements = new List<BaseElement>
                {
                    new Label { Text = "Test the HeritFromParent" },
                    new Paragraph
                    {
                        FontSize = "16",
                        ChildElements = new List<BaseElement>
                        {
                            new Label
                            {
                                Text = " Success (not the same size)",
                                SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve
                            }
                        }
                    }
                }
            });

            // Bookmark
            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Hyperlink
                    {
                        Anchor = "bmk",
                        Text = new Label
                        {
                            Text = "Link to the table of content ",
                            SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve
                        }
                    },
                    PageCrossReference("PAGEREF bmk")
                }
            });

            // Some specials characters
            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Label() { Text = "Label with special character (éèàù).", FontSize = "30", FontName = "Arial" }
                }
            });

            //hyperlink
            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Hyperlink()
                    {
                        Text = new Label()
                        {
                            Text = "Go to github.",
                            FontSize = "30",
                            FontName = "Arial",
                            FontColor = "40A6DB",
                            Underline = new UnderlineModel
                            {
                                Color = "40A6DB",
                                Val = UnderlineValues.DashedHeavy
                            }
                        },
                        WebSiteUri = "https://www.github.com/"
                    }
                }
            });

            // Indentation
            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Label() { Text = "This paragraph is indent from the left and the right", FontSize = "30", FontName = "Arial" }
                },
                Indentation = new ParagraphIndentationModel()
                {
                    Left = "300",
                    Right = "6000"
                }
            });

            // Paragraph with borders 1/2
            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Label() { Text = "This paragraph has borders", FontSize = "30", FontName = "Arial" }
                },
                Borders = new BorderModel()
                {
                    BorderPositions = BorderPositions.BOTTOM | BorderPositions.TOP | BorderPositions.LEFT,
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
                }
            });

            // Paragraph with borders 2/2 and space between lines
            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Label() { Text = Lorem_Ipsum }
                },
                Borders = new BorderModel()
                {
                    BorderPositions = (BorderPositions)13,
                    BorderWidth = 20,
                    BorderColor = "#ParagraphBorderColor#"
                },
                SpacingBetweenLines = 360
            });

            // Paragraph with tabulation and style 1/4
            page.ChildElements.Add(new Paragraph()
            {
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>()
                {
                    new Label()
                    {
                        Text = "Tabulation 1"
                    },
                    new Label()
                    {
                        TabulationProperties = new TabulationPropertiesModel()
                        {
                            TabStopPosition = 2500,
                            Leader = TabStopLeaderCharValues.dot,
                            Alignment = TabAlignmentValues.Right
                        }
                    },
                    new Label()
                    {
                        TabulationProperties = new TabulationPropertiesModel()
                        {
                            TabStopPosition = 5000,
                            Leader = TabStopLeaderCharValues.underscore,
                            Alignment = TabAlignmentValues.Right
                        }
                    }
                }
            });

            // Paragraph with tabulation and style 2/4
            page.ChildElements.Add(new Paragraph()
            {
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>()
                {
                    new Label(){Text = "Tabulation 2" },
                    new Label()
                    {
                        TabulationProperties = new TabulationPropertiesModel()
                    },
                    new Label(){Text = "After tabulation"}
                }
            });

            // Paragraph with tabulation and style 3/4
            page.ChildElements.Add(new Paragraph()
            {
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>()
                {
                    new Label(){ Text = "Tabulation 3"},
                    new Label(){Text = "test",FontColor = "FFFF00"},
                    new Label()
                    {
                        TabulationProperties = new TabulationPropertiesModel()
                        {
                            TabStopPosition = 5000,
                            Leader = TabStopLeaderCharValues.dot,
                            Alignment = TabAlignmentValues.Right
                        },
                        FontColor = "FFFF00"
                    },
                    new Label()
                    {
                        TabulationProperties = new TabulationPropertiesModel()
                        {
                            TabStopPosition = 10000,
                            Leader = TabStopLeaderCharValues.middleDot,
                            Alignment = TabAlignmentValues.Right
                        },
                        FontColor = "0000FF"
                    },
                    new Label(){ Text = "After 2 Tabulations"}
                }
            });

            // Paragraph with tabulation and style 4/4
            page.ChildElements.Add(new Paragraph()
            {
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>()
                {
                    new Label()
                    {
                        TabulationProperties = new TabulationPropertiesModel()
                        {
                            TabStopPosition = 10000,
                            Leader = TabStopLeaderCharValues.underscore,
                            Alignment = TabAlignmentValues.Right,
                        },
                        FontColor = "FFFF00"
                    },
                    new Label(){Text = "Tabulation 4"}
                }
            });

            // Image
            if (File.Exists(@"Resources\Desert.jpg"))
                page.ChildElements.Add(new Paragraph()
                {
                    ChildElements = new List<BaseElement>()
                    {
                        new Image()
                        {
                            MaxHeight = 100,
                            MaxWidth = 100,
                            Path = @"Resources\Desert.jpg",
                            ImagePartType = Engine.Packaging.ImagePartType.Jpeg
                        }
                    }
                });

            return page;
        }

        /// <summary>
        /// Generate paragraphs page with different page margin
        /// </summary>
        /// <returns></returns>
        private static Page GenerateParagraphSecondPage()
        {
            var page = new Page
            {
                Margin = new SpacingModel() { Top = 2500, Left = 845, Header = 1500, Footer = 709 }
            };

            // Paragraph with justification
            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label() { Text = "Text page 2", FontName = "Arial" }
                }
            });

            // Paragraphs with spacing before, after and a style
            page.ChildElements.Add(new Paragraph
            {
                SpacingBefore = 800,
                SpacingAfter = 800,
                Justification = JustificationValues.Both,
                ParagraphStyleId = "Yellow",
                ChildElements = new List<BaseElement>
                {
                    new Label() { Text = Lorem_Ipsum }
                }
            });

            return page;
        }

        #endregion

        #region Foreach

        /// <summary>
        /// Generate foreach context
        /// </summary>
        /// <param name="context"></param>
        public static void GenerateForeachContext(ContextModel context)
        {
            context.AddItem("#ForEachParagraph#", new DataSourceModel()
            {
                Items = new List<ContextModel>()
                {
                    new ContextModel().AddString("#TemplateKey#", "Template 1").AddString("#ForeachKeyTemplate1#", "This is the first template"),
                    new ContextModel().AddString("#TemplateKey#", "Template 2").AddString("#ForeachKeyTemplate2#", "This is the second template"),
                    new ContextModel().AddString("#TemplateKey#", "Template 1").AddString("#ForeachKeyTemplate1#", "This is the first template again")
                }
            });
        }

        /// <summary>
        /// Generate foreach and templates
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Page GenerateForeachPage(Document document)
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "Foreach test page"
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "This is an example of a list of template used in a foreach"
                    }
                }
            });

            document.TemplateDefinitions.Add(new TemplateDefinition()
            {
                TemplateId = "Template 1",
                Note = "Sample paragraph",
                ChildElements = new List<BaseElement>()
                {
                    new Paragraph
                    {
                        ChildElements = new List<BaseElement>
                        {
                            new Label { Text = "#ForeachKeyTemplate1#" }
                        }
                    }
                }
            });
            document.TemplateDefinitions.Add(new TemplateDefinition()
            {
                TemplateId = "Template 2",
                Note = "Sample paragraph",
                ChildElements = new List<BaseElement>()
                {
                    new Paragraph
                    {
                        Shading = "9EB5BA",
                        FontName = "Chiller",
                        ChildElements = new List<BaseElement>
                        {
                            new Label { Text = "#ForeachKeyTemplate2#" }
                        }
                    }
                }
            });

            // Foreach with template model
            page.ChildElements.Add(new ForEach()
            {
                DataSourceKey = "#ForEachParagraph#",
                ItemTemplate = new List<BaseElement>()
                {
                    new TemplateModel() { TemplateId = "#TemplateKey#" }
                }
            });

            return page;
        }

        #endregion

        #region Foreach page

        /// <summary>
        /// Generate foreach page context
        /// </summary>
        /// <param name="context"></param>
        public static void GenerateForeachPageContext(ContextModel context)
        {
            ContextModel page1 = new ContextModel().AddString("#Label#", "Foreach page First page");
            ContextModel page2 = new ContextModel().AddString("#Label#", "Foreach page Second page");

            context.AddCollection("#ForeachPageDataSource#", page1, page2);
        }

        /// <summary>
        /// Generate foreach page page
        /// </summary>
        /// <returns></returns>
        public static ForEachPage GenerateForeachPagePage()
        {
            return new ForEachPage
            {
                DataSourceKey = "#ForeachPageDataSource#",
                Margin = new SpacingModel() { Top = 1418, Left = 845, Header = 709, Footer = 709 },
                ChildElements = new List<BaseElement>
                {
                    new Paragraph
                    {
                        ParagraphStyleId = "Red",
                        ChildElements = new List<BaseElement>
                        {
                            new Label() { Text = "#Label#" }
                        }
                    },
                    new Paragraph
                    {
                        ChildElements = new List<BaseElement>
                        {
                            new Label() { Text = "           Text with space before", FontSize = "20", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve }
                        }
                    }
                }
            };
        }

        #endregion

        #region Table of content

        /// <summary>
        /// Generate table of content
        /// </summary>
        /// <returns></returns>
        public static Page GenerateTableOfContent()
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Label()
                    {
                        Shading = "2B3C4F",
                        FontSize = "26",
                        Text = "Table of content bookmark",
                    },
                    new BookmarkStart() {Id = "bmk", Name = "bmk" },
                    new BookmarkEnd(){Id = "bmk"}
                }
            });

            TableOfContents tableOfContents = new TableOfContents()
            {
                StylesAndLevels = new List<Tuple<string, string>>()
                {
                    new Tuple<string, string>("Red", "1"),
                    new Tuple<string, string>("Yellow", "2"),
                },
                Title = "Table of content!",
                TitleStyleId = "Red",
                ToCStylesId = new List<string>() { "TOC1", "TOC2" },
                LeaderCharValue = TabStopLeaderCharValues.underscore
            };
            page.ChildElements.Add(tableOfContents);

            return page;
        }

        #endregion

        #region Uniform grid

        /// <summary>
        /// Generate uniform grid context
        /// </summary>
        /// <param name="context"></param>
        private static void GenerateUniformGridContext(ContextModel context)
        {
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
        }

        /// <summary>
        /// Generate uniform gid page
        /// </summary>
        /// <returns></returns>
        private static Page GenerateUniformGridPage()
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "Uniform grid test page"
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "For this uniform grid, the number of cells is defined by the actual day from the beginning of the month"
                    }
                }
            });

            page.ChildElements.Add(new UniformGrid()
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
                                new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Header 1" } } }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "Header 2" }
                            }
                        }
                    }
                },
                Borders = new BorderModel()
                {
                    BorderPositions = BorderPositions.BOTTOM | BorderPositions.INSIDEVERTICAL,
                    BorderWidthBottom = 50,
                    BorderWidthInsideVertical = 1,
                    UseVariableBorders = true,
                    BorderColor = "FF0000"
                }
            });

            return page;
        }

        #endregion

        #region Tables

        /// <summary>
        /// Generate table context
        /// </summary>
        /// <param name="context"></param>
        private static void GenerateTableContext(ContextModel context)
        {
            // Classic generation
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

            // Fluent samples
            ContextModel row1 = new ContextModel()
                        .AddString("#Cell1#", "Col 1 Row 1")
                        .AddString("#Cell2#", "Col 2 Row 1")
                        .AddString("#Label#", "Label 1")
                        .AddDouble("#ColSpan#", 2, "{0}");

            context.AddCollection("#TableDataSource#", row1, row2)
                   .AddCollection("#DatasourcePrefix#", row1, row2, row3, row4);


            ContextModel row11 = new ContextModel()
                         .AddBoolean("#IsInGroup#", true)
                         .AddBoolean("#IsNotFirstLineGroup#", false)
                         .AddString("#Cell1#", "Col 1 Row 1")
                         .AddString("#Cell2#", "Col 2 Row 1")
                         .AddString("#Label#", "Label 1");
            ContextModel row12 = new ContextModel()
                         .AddBoolean("#IsInGroup#", true)
                         .AddBoolean("#IsNotFirstLineGroup#", true)
                         .AddString("#Cell1#", "Col 1 Row 1")
                         .AddString("#Cell2#", "Col 2 Row 1")
                         .AddString("#Label#", "Label 1");
            ContextModel row13 = new ContextModel()
                         .AddBoolean("#IsInGroup#", true)
                         .AddBoolean("#IsNotFirstLineGroup#", true)
                         .AddString("#Cell1#", "Col 1 Row 1")
                         .AddString("#Cell2#", "Col 2 Row 1")
                         .AddString("#Label#", "Label 1");
            ContextModel row22 = new ContextModel()
                         .AddBoolean("#IsInGroup#", true)
                         .AddBoolean("#IsNotFirstLineGroup#", false)
                         .AddString("#Cell1#", "Col 2 Row 1")
                         .AddString("#Cell2#", "Col 2 Row 2")
                         .AddString("#Label#", "Label 2");
            ContextModel row23 = new ContextModel()
                         .AddBoolean("#IsInGroup#", true)
                         .AddBoolean("#IsNotFirstLineGroup#", true)
                         .AddString("#Cell1#", "Col 2 Row 1")
                         .AddString("#Cell2#", "Col 2 Row 2")
                         .AddString("#Label#", "Label 2");

            context.AddCollection("#TableWithFusedCellsInContext#", row11, row12, row13, row22, row23);
        }

        /// <summary>
        /// Generate table page
        /// </summary>
        /// <returns></returns>
        private static Page GenerateTablesPage()
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "Table test page"
                    }
                }
            });

            // Table
            page.ChildElements.Add(new Table()
            {
                TableWidth = new TableWidthModel() { Width = "4000", Type = TableWidthUnitValues.Pct },
                TableIndentation = new TableIndentation() { Width = 1000 },
                Rows = new List<Row>()
                {
                    new Row()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell()
                            {
                                NoWrap = true,
                                VerticalAlignment = TableVerticalAlignmentValues.Center,
                                Justification = JustificationValues.Center,
                                ChildElements = new List<BaseElement>()
                                {
                                    new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Cell 1 - First paragraph" } }, ParagraphStyleId = "Yellow" },
                                    new Image()
                                    {
                                        Width = 50,
                                        Path = @"Resources\Desert.jpg",
                                        ImagePartType = Engine.Packaging.ImagePartType.Jpeg
                                    },
                                    new Label() { Text = "Cell 1 No Wrap - Label in a cell" },
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
                                        Path = @"Resources\Desert.jpg",
                                        ImagePartType = Engine.Packaging.ImagePartType.Jpeg
                                    },
                                    new Label() { Text = "Cell 2 - Second label" }
                                },
                                Borders = new BorderModel()
                                {
                                    BorderColor = "#BorderColor#",
                                    BorderWidth = 20,
                                    BorderPositions = BorderPositions.LEFT | BorderPositions.TOP
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
                },
                HeaderRow = new Row()
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Header from paragraph" } } }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "Header from label" }
                            }
                        }
                    }
                },
                Borders = new BorderModel()
                {
                    BorderPositions = BorderPositions.BOTTOM | BorderPositions.INSIDEVERTICAL,
                    BorderWidthBottom = 50,
                    BorderWidthInsideVertical = 1,
                    UseVariableBorders = true,
                    BorderColor = "FF0000"
                }
            });

            // Add a paragraph to avoid Word merging tables  
            page.ChildElements.Add(new Paragraph { ChildElements = new List<BaseElement>() { new Label() { Text = "Table with ColSpan in context ⏬" } } });

            // Table with ColSpan in context
            page.ChildElements.Add(new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                ColsWidth = new int[2] { 750, 4250 },
                Borders = new BorderModel()
                {
                    BorderPositions = (BorderPositions)63,
                    BorderColor = "328864",
                    BorderWidth = 20,
                },
                RowModel = new Row()
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            ColSpanKey = "#ColSpan#",
                            Shading = "FFA0FF",
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "#Cell1#" },
                                new Label() { Text = "ColSpan 2" },
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
                DataSourceKey = "#TableDataSource#"
            });

            // Add a paragraph to avoid Word merging tables 
            page.ChildElements.Add(new Paragraph { ChildElements = new List<BaseElement>() { new Label { Text = "Table with datasource with prefixs ⏬" } } });

            // Table with datasource with prefixs
            page.ChildElements.Add(new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                ColsWidth = new int[2] { 750, 4250 },
                Borders = new BorderModel()
                {
                    BorderPositions = (BorderPositions)63,
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
                                new Label()
                                {
                                    Text = "Item Datasource (0 index) #DataSourcePrefix_TableRow_IndexBaseZero# - ",
                                    ShowKey = "#DataSourcePrefix_TableRow_IsFirstItem#"
                                },
                                new Label() { Text = "#Cell1#" }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label()
                                {
                                    Text = "Item Datasource (1 index) #DataSourcePrefix_TableRow_IndexBaseOne# - ",
                                    ShowKey = "#DataSourcePrefix_TableRow_IsLastItem#"
                                },
                                new Label() { Text = "#Cell2#" }
                            }
                        }
                    }
                }
            });

            // Add a paragraph to avoid Word merging tables 
            page.ChildElements.Add(new Paragraph { ChildElements = new List<BaseElement>() { new Label { Text = "Table with before and after rows ⏬" } } });

            // Table with before and after rows
            page.ChildElements.Add(new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                ColsWidth = new int[2] { 750, 4250 },
                Borders = new BorderModel()
                {
                    BorderPositions = (BorderPositions)63,
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
                                    new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Merged with below" } }, ParagraphStyleId = "Yellow" },
                                },
                                Fusion = true
                            },
                            new Cell()
                            {
                                ChildElements = new List<BaseElement>()
                                {
                                    new Label() { Text = "A label ", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve },
                                    new Image()
                                    {
                                        MaxHeight = 75,
                                        MaxWidth = 75,
                                        Path = @"Resources\Desert.jpg",
                                        ImagePartType = Engine.Packaging.ImagePartType.Jpeg
                                    },
                                    new Label() { Text = " with an image", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve }
                                },
                                Borders = new BorderModel()
                                {
                                    BorderColor = "00FF22",
                                    BorderWidth = 15,
                                    BorderPositions = BorderPositions.RIGHT | BorderPositions.TOP
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
                                    new Label() { Text = "Text set at bottom right position" }
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
                                new Label() { Text = "Cell: #Cell1#" }
                            }
                        },
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "Cell: #Cell2#" }
                            }
                        }
                    }
                },
                AfterRows = new List<Row>()
                {
                    new Row()
                    {
                        Cells = new List<Cell>()
                        {
                            new Cell()
                            {
                                ColSpan = 2,
                                Justification = JustificationValues.Center,
                                ChildElements = new List<BaseElement>()
                                {
                                    new Label() { Text = "After row" }
                                }
                            }
                        }
                    }
                },
                DataSourceKey = "#TableDataSource#"
            });

            // Add a paragraph to avoid Word merging tables  
            page.ChildElements.Add(new Paragraph { ChildElements = new List<BaseElement>() { new Label { Text = "Fused table ⏬" } } });

            // Fused table
            page.ChildElements.Add(new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                ColsWidth = new int[3] { 1200, 1200, 1200 },
                Borders = new BorderModel()
                {
                    BorderPositions = (BorderPositions)63,
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
                DataSourceKey = "#TableWithFusedCellsInContext#"
            });

            return page;
        }

        #endregion

        #region Substitutable string

        /// <summary>
        /// Generate substitutable string context
        /// </summary>
        /// <param name="context"></param>
        private static void GenerateSubstitutableStringContext(ContextModel context)
        {
            string textToDisplay = "DateTimeModel : {0}\n DoubleModel : {1}\n StringModel : {2}\n";
            ContextModel rowSubstitutable = new ContextModel();
            rowSubstitutable.AddItem("#SubstitutableStringData#",
                new SubstitutableStringModel(
                    textToDisplay,
                    new ContextModel()
                        .AddDateTime("#Val1#", DateTime.Now, "D")
                        .AddDouble("#Val2#", 5.4, "This number is displayed with a render pattern : {0}")
                        .AddString("#Val3#", "This text is substituted")
                )
            );
            rowSubstitutable.AddItem("#SubstitutableStringDataWithLessParameters#",
                new SubstitutableStringModel(
                    textToDisplay,
                    new ContextModel()
                        .AddDateTime("#Val1#", DateTime.Now, null)
                )
            );
            rowSubstitutable.AddItem("#SubstitutableStringDataWithMoreParameters#",
                new SubstitutableStringModel(
                    textToDisplay,
                    new ContextModel()
                        .AddDateTime("#Val1#", DateTime.Now, null)
                        .AddDouble("#Val2#", 5.4, null)
                        .AddString("#Val3#", "This text is substituted")
                        .AddDouble("#Val4#", 75, null)
                        .AddString("#Val5#", "Not displayed string")
                )
            );

            context.AddCollection("#SubstitutableStringDataSourceModel#", rowSubstitutable);
        }

        /// <summary>
        /// Generate substitutable string page
        /// </summary>
        /// <returns></returns>
        private static Page GenerateSubstitutableStringPage()
        {
            var page = new Page();

            // Substitutable string
            var pargraphTitle = new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red"
            };
            pargraphTitle.ChildElements.Add(new Label() { Text = "Substitutable string", FontName = "Arial" });
            page.ChildElements.Add(pargraphTitle);

            page.ChildElements.Add(new Table()
            {
                RowModel = new Row()
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label()
                                {
                                    Text = "Matching of supplied parameters and expected parameters: \n",
                                    Bold = true,
                                    Underline = new UnderlineModel () { Val = UnderlineValues.Single }
                                },
                                new Label() { Text = "#SubstitutableStringData#" },

                                new Label() { Text = Environment.NewLine },
                                new Label()
                                {
                                    Text = "Less supplied parameters than expected parameters: \n",
                                    Bold = true,
                                    Underline = new UnderlineModel () { Val = UnderlineValues.Single }
                                },
                                new Label() { Text = "#SubstitutableStringDataWithLessParameters#" },

                                new Label() { Text = Environment.NewLine },
                                new Label()
                                {
                                    Text = "More supplied parameters than expected parameters: \n",
                                    Bold = true,
                                    Underline = new UnderlineModel () { Val = UnderlineValues.Single }
                                },
                                new Label() { Text = "#SubstitutableStringDataWithMoreParameters#" }
                            }
                        }
                    }
                },
                DataSourceKey = "#SubstitutableStringDataSourceModel#"
            });

            return page;
        }

        #endregion

        #region Charts

        #region Pie chart

        /// <summary>
        /// Generate context for pie charts
        /// </summary>
        /// <param name="context"></param>
        private static void GeneratePieChartContext(ContextModel context)
        {
            context.AddItem("#PieGraphSampleData#", new SingleSerieChartModel()
            {
                ChartContent = new SingleSeriesModel()
                {
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel()
                        {
                            Name = "A Category",
                            Color = "9FA0A4"
                        },
                        new CategoryModel()
                        {
                            Name = "B Category",
                            Color = "32AD3C"
                        },
                        new CategoryModel()
                        {
                            Name = "C Category",
                            Color = "E47F00"
                        },
                        new CategoryModel()
                        {
                            Name = "D Category",
                            Color = "DC0A0A"
                        },
                        new CategoryModel()
                        {
                            Name = "E Category"
                        },
                        new CategoryModel()
                        {
                            Name = "F Category"
                        }
                    },
                    Serie = new SerieModel()
                    {
                        Values = new List<double?>()
                        {
                            10, 20, 5, 50, 15, null
                        },
                        Name = "Serie 1",
                        HasBorder = true,
                        BorderColor = "#FFFFFF",
                        Color = "#000000",
                        BorderWidth = 63500
                    }
                }
            });
        }

        /// <summary>
        /// Generate Pie graph page templates
        /// </summary>
        /// <returns></returns>
        private static Page GeneratePieChartPage()
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "Pie graphs test page"
                    }
                }
            });

            var pieChartPr = new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new PieModel()
                    {
                        Title = "Pie Chart test",
                        ShowTitle = true,
                        ShowChartBorder = true,
                        PieChartType = PieChartType.PieChart,
                        DataSourceKey = "#PieGraphSampleData#",
                        DataLabel = new DataLabelModel()
                        {
                            ShowCatName = true,
                            ShowPercent = true,
                            Separator = "\n",
                            FontSize = 8
                        },
                        DataLabelColor = "#FFFF00"
                    }
                }
            };

            page.ChildElements.Add(pieChartPr);

            return page;
        }

        #endregion

        #region Bar chart

        /// <summary>
        /// Generate context for Bar Graphs
        /// </summary>
        /// <param name="context"></param>
        private static void GenerateBarGraphContext(ContextModel context)
        {
            // Old bar graph objects
            context.AddItem("#OldBarGraphSampleData#", new BarChartModel()
            {
                BarChartContent = new Engine.ReportEngine.DataContext.Charts.BarModel()
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
                            Name = "Bar serie 1",
                            Color = "9FA0A4"
                        },
                        new BarSerieModel()
                        {
                            Values = new List<double?>()
                            {
                                5, null, 7, 8, 0, 10
                            },
                            Name = "Bar serie 2",
                            Color = "32AD3C"
                        },
                        new BarSerieModel()
                        {
                            Values = new List<double?>()
                            {
                                9, 10, 11, 12, 13, 14
                            },
                            Name = "Bar serie 3",
                            Color = "E47F00"
                        },
                        new BarSerieModel()
                        {
                            Values = new List<double?>()
                            {
                                9, 10, 11, 12, 15, 25
                            },
                            Name = "Bar serie 4",
                            Color = "DC0A0A"
                        }
                    }
                }
            });

            context.AddItem("#BarGraphSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel()
                        {
                            Name = "Category 1"
                        },
                        new CategoryModel()
                        {
                            Name = "Category 2"
                        },
                        new CategoryModel()
                        {
                            Name = "Category 3"
                        },
                        new CategoryModel()
                        {
                            Name = "Category 4"
                        },
                        new CategoryModel()
                        {
                            Name = "Category 5"
                        },
                        new CategoryModel()
                        {
                            Name = "Category 6"
                        }
                    },
                    Series = new List<SerieModel>()
                    {
                        new SerieModel()
                        {
                            Values = new List<double?>()
                            {
                                0, 1, 2, 3, 6, null
                            },
                            Name = "Bar serie 1",
                            Color = "9FA0A4",
                            HasBorder = true,
                            BorderColor = "#FF00FF",
                            BorderWidth = 63500
                        },
                        new SerieModel()
                        {
                            Values = new List<double?>()
                            {
                                5, null, 7, 8, 0, 10
                            },
                            Name = "Bar serie 2",
                            Color = "32AD3C",
                            HasBorder = true,
                            BorderColor = "#0000FF",
                            BorderWidth = 63500
                        },
                        new SerieModel()
                        {
                            Values = new List<double?>()
                            {
                                9, 10, 11, 12, 13, 14
                            },
                            Name = "Bar serie 3",
                            Color = "E47F00"
                        },
                        new SerieModel()
                        {
                            Values = new List<double?>()
                            {
                                9, 10, 11, 12, 15, 25
                            },
                            Name = "Bar serie 4",
                            Color = "DC0A0A"
                        }
                    }
                }
            });

            context.AddItem("#SingleStackedBarGraphSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel()
                        {
                            Name = "Cat name"
                        }
                    },
                    Series = new List<SerieModel>()
                    {
                        new SerieModel()
                        {
                            Name = "Serie 1",
                            Color = "9FA0A4",
                            Values = new List<double?>{ 98 }
                        },
                        new SerieModel()
                        {
                            Name = "Serie 2",
                            Color = "E47F00",
                            Values = new List<double?>{ 2 }
                        }
                    }
                }
            });

            context.AddItem("#BarGraphNumericCategoriesSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    CategoryType = CategoryType.NumberReference,
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel() { Value = 1.24680135790987654321 },
                        new CategoryModel() { Value = 2.24680135790987654321 },
                        new CategoryModel() { Value = 3.24680135790987654321 },
                        new CategoryModel() { Value = 4.24680135790987654321 },
                        new CategoryModel() { Value = 5.24680135790987654321 },
                        new CategoryModel() { Value = 6.24680135790987654321 }
                    },
                    Series = new List<SerieModel>()
                    {
                        new SerieModel()
                        {
                            Values = new List<double?>()
                            {
                                0, 1, 2, 3, 6, null
                            },
                            Name = "Bar serie 1",
                            Color = "F5E642",
                        },
                        new SerieModel()
                        {
                            Values = new List<double?>()
                            {
                                5, null, 7, 8, 0, 10
                            },
                            Name = "Bar serie 2",
                            Color = "75F23F",
                        },
                        new SerieModel()
                        {
                            Values = new List<double?>()
                            {
                                9, 10, 11, 12, 13, 14
                            },
                            Name = "Bar serie 3",
                            Color = "28C7BC"
                        },
                        new SerieModel()
                        {
                            Values = new List<double?>()
                            {
                                9, 10, 11, 12, 15, 25
                            },
                            Name = "Bar serie 4",
                            Color = "C728C7"
                        }
                    }
                }
            });
        }

        /// <summary>
        /// Generate Bar graph page templates
        /// </summary>
        /// <returns></returns>
        private static Page GenerateBarChartPage()
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "Bar graphs test page"
                    }
                }
            });

            // Old bar graph objects
            page.ChildElements.Add(new Paragraph()
            {
                ChildElements = new List<BaseElement>()
                {
                    new Engine.Word.ReportEngine.Models.Charts.BarModel()
                    {
                        Title = "Bar test",
                        ShowTitle = true,
                        BarChartType = BarChartType.BarChart,
                        BarDirectionValues = BarDirectionValues.Column,
                        BarGroupingValues = BarGroupingValues.PercentStacked,
                        ValuesAxisModel = new ChartAxisModel
                        {
                            ShowMajorGridlines = true
                        },
                        DataSourceKey = "#OldBarGraphSampleData#",
                        MaxHeight = 320
                    }
                }
            });

            page.ChildElements.Add(new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new Engine.Word.ReportEngine.Models.Charts.BarModel()
                    {
                        Title = "Bordered serie bar test",
                        ShowTitle = true,
                        FontSize = "23",
                        BarChartType = BarChartType.BarChart,
                        BarDirectionValues = BarDirectionValues.Column,
                        BarGroupingValues = BarGroupingValues.PercentStacked,
                        DataSourceKey = "#BarGraphSampleData#",
                        ValuesAxisModel = new ChartAxisModel
                        {
                            ShowMajorGridlines = true
                        },
                        MaxHeight = 320
                    }
                }
            });

            page.ChildElements.Add(new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new Engine.Word.ReportEngine.Models.Charts.BarModel()
                    {
                        Title = "Single stacked Graph without min-max",
                        ShowTitle = true,
                        MaxHeight = 100,
                        CategoriesAxisModel = new ChartAxisModel
                        {
                            DeleteAxis = true
                        },
                        ValuesAxisModel = new ChartAxisModel
                        {
                            DeleteAxis = true
                        },
                        ShowLegend = false,
                        HasBorder = false,
                        DataSourceKey = "#SingleStackedBarGraphSampleData#"
                    }
                }
            });

            page.ChildElements.Add(new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new Engine.Word.ReportEngine.Models.Charts.BarModel()
                    {
                        Title = "Single stacked Graph with min-max",
                        ShowTitle = true,
                        MaxHeight = 100,
                        CategoriesAxisModel = new ChartAxisModel
                        {
                            DeleteAxis = true
                        },
                        ValuesAxisModel = new ChartAxisModel
                        {
                            DeleteAxis = true
                        },
                        ShowLegend = false,
                        HasBorder = false,
                        DataSourceKey = "#SingleStackedBarGraphSampleData#",
                        ValuesAxisScaling = new BarChartScalingModel()
                        {
                            MinAxisValue = 0,
                            MaxAxisValue = 100
                        }
                    }
                }
            });

            page.ChildElements.Add(new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new Engine.Word.ReportEngine.Models.Charts.BarModel()
                    {
                        Title = "Graph with numeric categories test",
                        ShowTitle = true,
                        FontSize = "23",
                        BarChartType = BarChartType.BarChart,
                        BarDirectionValues = BarDirectionValues.Column,
                        BarGroupingValues = BarGroupingValues.PercentStacked,
                        DataSourceKey = "#BarGraphNumericCategoriesSampleData#",
                        MaxHeight = 320
                    }
                }
            });

            return page;
        }

        #endregion

        #region Line Graph

        /// <summary>
        /// Generate context for line graphs
        /// </summary>
        /// <param name="context"></param>
        private static void GenerateLineGraphContext(ContextModel context)
        {
            context.AddItem("#LineGraphStandardSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel() { Name = "1" },
                        new CategoryModel() { Name = "2" },
                        new CategoryModel() { Name = "3" },
                        new CategoryModel() { Name = "4" },
                        new CategoryModel() { Name = "5" }
                    },
                    Series = new List<SerieModel>
                    {
                        new SerieModel()
                        {
                            Name = "Multiple of two",
                            Color = "9FA0A4",
                            Values = new List<double?> { 2, 4, 6, 8, 10 }
                        }
                    },
                    CategoriesAxisModel = new AxisModel
                    {
                        Title = "From context",
                        Color = "9FA0A4"
                    }
                }
            });

            context.AddItem("#LineGraphStandardSecondaryAxisSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel() { Name = "1" },
                        new CategoryModel() { Name = "2" },
                        new CategoryModel() { Name = "3" },
                        new CategoryModel() { Name = "4" },
                        new CategoryModel() { Name = "5" }
                    },
                    Series = new List<SerieModel>
                    {
                        new SerieModel()
                        {
                            Name = "Multiple of two",
                            Color = "#874054",
                            Values = new List<double?> { -2, -4, 6, 8, 10 }
                        },
                        new SerieModel()
                        {
                            Name = "Multiple of three",
                            Color = "#080890",
                            Values = new List<double?> { 3, 6, -9, 12, 15 },
                            UseSecondaryAxis = true,
                            SmoothCurve = true,
                            PresetLineDashValues = PresetLineDashValues.DashDot
                        }
                    },
                    ValuesAxisModel = new AxisModel
                    {
                        Title = "Gauche",
                        Color = "874054"
                    },
                    SecondaryValuesAxisModel = new AxisModel
                    {
                        Title = "Droite",
                        Color = "080890"
                    }
                }
            });

            context.AddItem("#LineGraphWithNumericCategoriesSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel() { Value = 1.1234567890 },
                        new CategoryModel() { Value = 3.1234567890 },
                        new CategoryModel() { Value = 5.1234567890 },
                        new CategoryModel() { Value = 7.1234567890 },
                        new CategoryModel() { Value = 9.1234567890 }
                    },
                    Series = new List<SerieModel>
                    {
                        new SerieModel()
                        {
                            Name = "Multiple of two",
                            Color = "9FA0A4",
                            Values = new List<double?> { 2, 4, 6, 8, 10 }
                        }
                    },
                    CategoriesAxisModel = new AxisModel
                    {
                        Title = "From context",
                        Color = "9FA0A4"
                    }
                }
            });
        }

        /// <summary>
        ///  Generate Line graph page templates
        /// </summary>
        /// <returns></returns>
        private static Page GenerateLineGraphsPage()
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "Line graphs test page"
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement> {
                    new LineModel
                    {
                        Title = "Line graph test",
                        ShowTitle = true,
                        FontSize = "23",
                        DataSourceKey = "#LineGraphStandardSampleData#",
                        MaxHeight = 320,
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        ShowLegend = true,
                        LegendPosition = LegendPositionValues.Bottom,
                        ValuesAxisModel = new ChartAxisModel
                        {
                            ShowMajorGridlines = true,
                            MajorGridlinesColor = "FF0000"
                        }
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement> {
                    new LineModel
                    {
                        Title = "Line graph with secondary axis test",
                        ShowTitle = true,
                        DataSourceKey = "#LineGraphStandardSecondaryAxisSampleData#",
                        MaxHeight = 320,
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        ValuesAxisModel = new ChartAxisModel
                        {
                            ShowMajorGridlines = true,
                            MajorGridlinesColor = "48C9B0",
                            ShowAxisCurve = true,
                            AxisCurveColor = "00FF00"
                        },
                        CategoriesAxisModel = new ChartAxisModel
                        {
                            ShowMajorGridlines = true,
                            MajorGridlinesColor = "48C9B0",
                            Title = "Categories !"
                        }
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement> {
                    new LineModel
                    {
                        CategoryType = CategoryType.NumberReference,
                        Title = "Line graph with numeric categories test",
                        ShowTitle = true,
                        FontSize = "23",
                        DataSourceKey = "#LineGraphWithNumericCategoriesSampleData#",
                        MaxHeight = 320,
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        ShowLegend = true,
                        LegendPosition = LegendPositionValues.Bottom,
                        ValuesAxisModel = new ChartAxisModel
                        {
                            ShowMajorGridlines = true,
                            MajorGridlinesColor = "151515"
                        }
                    }
                }
            });

            return page;
        }

        #endregion

        #region Scatter Graph

        /// <summary>
        /// Generate context for scatter graphs
        /// </summary>
        /// <param name="context"></param>
        private static void GenerateScatterGraphContext(ContextModel context)
        {
            context.AddItem("#ScatterGraphStandardSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    ScatterSeries = new List<ScatterSerieModel>
                    {
                        new ScatterSerieModel
                        {
                            Name = "Cloud",
                            HideCurve = false,
                            Color = "2CADD4",
                            Values = new List<CurvePoint>
                            {
                                new CurvePoint{ X = 1, Y = 1 },
                                new CurvePoint{ X = 2, Y = 5 },
                                new CurvePoint{ X = 3, Y = 8 },
                                new CurvePoint{ X = 4, Y = 2 },
                                new CurvePoint{ X = 5, Y = 6 }
                            }
                        }
                    },
                    CategoriesAxisModel = new AxisModel
                    {
                        Title = "Cloud from context",
                        Color = "9FA0A4"
                    }
                }
            });

            context.AddItem("#ScatterGraphCloudSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    ScatterSeries = new List<ScatterSerieModel>
                    {
                        new ScatterSerieModel
                        {
                            Name = "Cloud",
                            HideCurve = true,
                            Color = "2CADD4",
                            Values = new List<CurvePoint>
                            {
                                new CurvePoint{ X = 1, Y = 1 },
                                new CurvePoint{ X = 2, Y = 5 },
                                new CurvePoint{ X = 3, Y = 8 },
                                new CurvePoint{ X = 4, Y = 2 },
                                new CurvePoint{ X = 5, Y = 6 }
                            },
                            SerieMarker = new SerieMarker
                            {
                                MarkerStyleValues = MarkerStyleValues.Star
                            }
                        }
                    }
                }
            });

            context.AddItem("#ScatterGraphCurvesWithDifferentXAxisSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    ScatterSeries = new List<ScatterSerieModel>
                    {
                        new ScatterSerieModel
                        {
                            Name = "5 points",
                            Color = "2CADD4",
                            Values = new List<CurvePoint>
                            {
                                new CurvePoint{ X = 1, Y = 1 },
                                new CurvePoint{ X = 2, Y = 5 },
                                new CurvePoint{ X = 3, Y = 8 },
                                new CurvePoint{ X = 4, Y = 2 },
                                new CurvePoint{ X = 5, Y = 6 }
                            }
                        },
                        new ScatterSerieModel
                        {
                            Name = "10 points",
                            Color = "32ED76",
                            UseSecondaryAxis = true,
                            Values = new List<CurvePoint>
                            {
                                new CurvePoint{ X = 1, Y = 10 },
                                new CurvePoint{ X = 2, Y = 20 },
                                new CurvePoint{ X = 3, Y = 5 },
                                new CurvePoint{ X = 4, Y = 8 },
                                new CurvePoint{ X = 5, Y = 15 },
                                new CurvePoint{ X = 6, Y = 18 },
                                new CurvePoint{ X = 7, Y = 1 },
                                new CurvePoint{ X = 8, Y = 7 },
                                new CurvePoint{ X = 9, Y = 3 },
                                new CurvePoint{ X = 10, Y = 20 }
                            }
                        }
                    },
                    CategoriesAxisModel = new AxisModel
                    {
                        Title = "Cloud from context",
                        Color = "9FA0A4",
                        MinimumValue = 1,
                        MaximumValue = 5
                    },
                    SecondaryCategoriesAxisModel = new AxisModel
                    {
                        MinimumValue = 1,
                        MaximumValue = 10
                    }
                }
            });
        }

        /// <summary>
        ///  Generate Scatter graph page templates
        /// </summary>
        /// <returns></returns>
        private static Page GenerateScatterGraphsPage()
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "Scatter graphs test page"
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement> {
                    new ScatterModel
                    {
                        Title = "Scatter graph - 1 curve",
                        ShowTitle = true,
                        FontSize = "23",
                        DataSourceKey = "#ScatterGraphStandardSampleData#",
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        ShowLegend = true,
                        LegendPosition = LegendPositionValues.Bottom,
                        ValuesAxisModel = new ChartAxisModel
                        {
                            ShowMajorGridlines = true,
                            MajorGridlinesColor = "F2308B"
                        }
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement> {
                    new ScatterModel
                    {
                        Title = "Scatter graph - Cloud Points",
                        ShowTitle = true,
                        FontSize = "23",
                        DataSourceKey = "#ScatterGraphCloudSampleData#",
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        ShowLegend = true,
                        LegendPosition = LegendPositionValues.Bottom
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement> {
                    new ScatterModel
                    {
                        Title = "Scatter graph - 2 curves with different X number",
                        ShowTitle = true,
                        FontSize = "23",
                        DataSourceKey = "#ScatterGraphCurvesWithDifferentXAxisSampleData#",
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        ShowLegend = true,
                        LegendPosition = LegendPositionValues.Bottom,
                        SecondaryCategoriesAxisModel = new ChartAxisModel
                        {
                            ShowAxisCurve = false,
                            TickLabelPosition = TickLabelPositionValues.None
                        }
                    }
                }
            });

            return page;
        }

        #endregion

        #region Combine Graph

        /// <summary>
        /// Generate context for Combined graphs
        /// </summary>
        /// <param name="context"></param>
        private static void GenerateCombineGraphContext(ContextModel context)
        {
            context.AddItem("#CombineGraphOnlyLineSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel() { Name = "1" },
                        new CategoryModel() { Name = "3" },
                        new CategoryModel() { Name = "5" },
                        new CategoryModel() { Name = "7" },
                        new CategoryModel() { Name = "9" }
                    },
                    Series = new List<SerieModel>
                    {
                        new SerieModel()
                        {
                            Name = "Multiple of two",
                            Color = "9FA0A4",
                            Values = new List<double?> { 2, 4, 6, 8, 10 }
                        },
                        new SerieModel()
                        {
                            Name = "Multiple of three",
                            Color = "6C8BE0",
                            Values = new List<double?> { 1, 3, 5, 7, 9 }
                        }
                    },
                    CategoriesAxisModel = new AxisModel
                    {
                        Title = "From context",
                        Color = "9FA0A4"
                    }
                }
            });

            context.AddItem("#CombineGraphOnlyBarSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel() { Name = "1" },
                        new CategoryModel() { Name = "2" },
                        new CategoryModel() { Name = "3" },
                        new CategoryModel() { Name = "4" },
                        new CategoryModel() { Name = "5" }
                    },
                    Series = new List<SerieModel>
                    {
                        new SerieModel()
                        {
                            Name = "Multiple of two",
                            Color = "9FA0A4",
                            Values = new List<double?> { 2, 4, 6, 8, 10 },
                            SerieChartType = SerieChartType.Bar
                        },
                        new SerieModel()
                        {
                            Name = "Multiple of three",
                            Color = "6C8BE0",
                            Values = new List<double?> { 1, 3, 5, 7, 9 },
                            SerieChartType = SerieChartType.Bar
                        }
                    },
                    CategoriesAxisModel = new AxisModel
                    {
                        Title = "From context",
                        Color = "9FA0A4",
                        CrossesAt = 3.5
                    }
                }
            });

            context.AddItem("#CombineGraphFrankensteinSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel() { Name = "1" },
                        new CategoryModel() { Name = "2.2" },
                        new CategoryModel() { Name = "3" },
                        new CategoryModel() { Name = "4" },
                        new CategoryModel() { Name = "5" },
                        new CategoryModel() { Name = "6" }
                    },
                    Series = new List<SerieModel>
                    {
                        new SerieModel()
                        {
                            Name = "Multiple of two",
                            Color = "EB347A",
                            Values = new List<double?> { 2, 4, 6, 8, 10, 12 },
                            SerieChartType = SerieChartType.Line
                        },
                        new SerieModel()
                        {
                            Name = "Orange",
                            Color = "E38812",
                            Values = new List<double?> { 2, 3, 5, 9, 10, 15 },
                            SerieChartType = SerieChartType.Line,
                            UseSecondaryAxis = true
                        },
                        new SerieModel()
                        {
                            Name = "Multiple of two",
                            Color = "12E3E3",
                            Values = new List<double?> { 2, 4, -6, 8, 10, 12 },
                            SerieChartType = SerieChartType.Bar
                        }
                    },
                    CategoriesAxisModel = new AxisModel
                    {
                        Title = "From context",
                        LabelFormat = "{0} with 'unit'",
                        Color = "9FA0A4",
                        CrossesAt = 2.3
                    },
                    ValuesAxisModel = new AxisModel
                    {
                        CrossesAt = 3
                    },
                    SecondaryValuesAxisModel = new AxisModel
                    {
                        CrossesAt = 5
                    }
                }
            });

            context.AddItem("#CombineGraphNumericFrankensteinSampleData#", new MultipleSeriesChartModel()
            {
                ChartContent = new MultipleSeriesModel()
                {
                    CategoryType = CategoryType.NumberReference,
                    Categories = new List<CategoryModel>()
                    {
                        new CategoryModel() { Value = 1.24680135790987654321 },
                        new CategoryModel() { Value = 2.24680135790987654321 },
                        new CategoryModel() { Value = 3.24680135790987654321 },
                        new CategoryModel() { Value = 4.24680135790987654321 },
                        new CategoryModel() { Value = 5.24680135790987654321 },
                        new CategoryModel() { Value = 6.24680135790987654321 }
                    },
                    Series = new List<SerieModel>
                    {
                        new SerieModel()
                        {
                            Name = "Multiple of two",
                            Color = "EB4934",
                            Values = new List<double?> { 2, 4, 6, 8, 10, 12 },
                            SerieChartType = SerieChartType.Line
                        },
                        new SerieModel()
                        {
                            Name = "Second Line",
                            Color = "16979C",
                            Values = new List<double?> { 2, 3, 5, 9, 10, 15 },
                            SerieChartType = SerieChartType.Line
                        },
                        new SerieModel()
                        {
                            Name = "Multiple of two",
                            Color = "D80DDB",
                            Values = new List<double?> { 2, 4, -6, 8, 10, 12 },
                            SerieChartType = SerieChartType.Bar
                        }
                    },
                    CategoriesAxisModel = new AxisModel
                    {
                        Title = "From context",
                        LabelFormat = "{0} with 'unit'",
                        Color = "9FA0A4",
                        CrossesAt = 2.3
                    }
                }
            });
        }

        /// <summary>
        /// Generate Combined graph templates
        /// </summary>
        /// <returns></returns>
        private static Page GenerateCombineGraphsPage()
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "Combine graphs test page"
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new CombineChartModel
                    {
                        Title = "Combine Chart Model - Only Line",
                        ShowTitle = true,
                        DataSourceKey = "#CombineGraphOnlyLineSampleData#",
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        ValuesAxisModel = new ChartAxisModel
                        {
                            ShowMajorGridlines = true,
                            MajorGridlinesColor = "34EBC6"
                        }
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new CombineChartModel
                    {
                        Title = "Combine Chart Model - Only Bar",
                        ShowTitle = true,
                        DataSourceKey = "#CombineGraphOnlyBarSampleData#",
                        MaxHeight = 320,
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        Overlap = 0,
                        ShowLegend = true,
                        LegendPosition = LegendPositionValues.Bottom
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new CombineChartModel
                    {
                        Title = "Combine Chart Model - Frankenstein",
                        ShowTitle = true,
                        DataSourceKey = "#CombineGraphFrankensteinSampleData#",
                        MaxHeight = 320,
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        ShowLegend = true,
                        FontFamilyLegend = "Arial",
                        LegendPosition = LegendPositionValues.Top
                    }
                }
            });

            page.ChildElements.Add(new Paragraph
            {
                ChildElements = new List<BaseElement>
                {
                    new CombineChartModel
                    {
                        Title = "Combine Chart Model - Numeric Frankenstein",
                        ShowTitle = true,
                        DataSourceKey = "#CombineGraphNumericFrankensteinSampleData#",
                        MaxHeight = 320,
                        DataLabel = new DataLabelModel { ShowDataLabel = false },
                        ShowLegend = true,
                        FontFamilyLegend = "Arial",
                        LegendPosition = LegendPositionValues.Right
                    }
                }
            });

            return page;
        }

        #endregion

        #endregion Charts

        #region Multiple columns 

        /// <summary>
        /// Generate the multiple columns context
        /// </summary>
        /// <param name="context"></param>
        private static void GenerateMultipleColumnsContext(ContextModel context)
        {
            context.AddDouble("#ColumnNumber#", 2.0, null);
        }

        /// <summary>
        /// Create table on the first page for the multiple columns example
        /// </summary>
        /// <returns></returns>
        private static Page GenerateTableOn1stPage()
        {
            var page = new Page();

            page.ChildElements.Add(new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red",
                ChildElements = new List<BaseElement>
                {
                    new Label
                    {
                        Text = "Multiple columns test page"
                    }
                }
            });

            var tableDataSourceWithBeforeAfter = new Table()
            {
                TableWidth = new TableWidthModel() { Width = "5000", Type = TableWidthUnitValues.Pct },
                ColsWidth = new int[2] { 1500, 3400 },
                Borders = new BorderModel()
                {
                    BorderPositions = (BorderPositions)63,
                    BorderColor = "328864",
                    BorderWidth = 8,
                },
                HeaderRow = new Row
                {
                    Cells = new List<Cell>
                    {
                        new Cell
                        {
                            Margin = new MarginModel { Left = 100 },
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "Multiple columns Header table" }
                            }
                        },
                        new Cell
                        {
                            Margin = new MarginModel { Left = 100 },
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "This table will be full width" }
                            }
                        }
                    }
                },
                Rows = new List<Row>()
                {
                    new Row
                    {
                        Cells = new List<Cell>
                        {
                            new Cell
                            {
                                Shading = "BCF5F3"
                            },
                            new Cell
                            {
                                Margin = new MarginModel { Left = 100 },
                                ChildElements = new List<BaseElement>()
                                {
                                    new Label() { Text = "The text after this table will be splited on multiple columns" }
                                }
                            }
                        }
                    }
                }
            };

            page.ChildElements.Add(tableDataSourceWithBeforeAfter);

            return page;
        }

        /// <summary>
        /// Create paragraph on second page merged with the previous one for multiple columns example
        /// </summary>
        /// <returns></returns>
        private static Page Generate2ColmunOnSamePage()
        {
            var page = new Page
            {
                ColumnNumberKey = "#ColumnNumber#"
            };

            // Define Paragraph
            var p2 = new Paragraph
            {
                SpacingBefore = 800,
                SpacingAfter = 800,
                Justification = JustificationValues.Both
            };
            string wideText = Lorem_Ipsum;
            wideText += "\n\n" + Lorem_Ipsum;
            wideText += "\n\n" + Lorem_Ipsum;
            wideText += "\n\n";
            p2.ChildElements.Add(new Label() { Text = wideText + wideText });
            page.ChildElements.Add(p2);

            return page;
        }

        #endregion Multiple columns 

        /// <summary>
        /// Manage headers and footers
        /// </summary>
        /// <param name="doc"></param>
        private static void ManageHeadersAndFooters(Document doc)
        {
            // Header
            var header = new Header
            {
                Type = HeaderFooterValues.Default
            };
            var ph = new Paragraph
            {
                ChildElements = new List<BaseElement>()
                {
                    new Label()
                    {
                        Text = "Header Text ",
                        SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve
                    }
                }
            };
            if (File.Exists(@"Resources\Desert.jpg"))
                ph.ChildElements.Add(new Image()
                {
                    MaxHeight = 100,
                    MaxWidth = 100,
                    Path = @"Resources\Desert.jpg",
                    ImagePartType = Engine.Packaging.ImagePartType.Jpeg
                });
            header.ChildElements.Add(ph);
            doc.Headers.Add(header);

            // first header
            var firstHeader = new Header
            {
                Type = HeaderFooterValues.First
            };
            var fph = new Paragraph();
            fph.ChildElements.Add(new Label() { Text = "First header Text" });
            firstHeader.ChildElements.Add(fph);
            doc.Headers.Add(firstHeader);

            // Footer
            var footer = new Footer
            {
                Type = HeaderFooterValues.Default
            };
            var pf = new Paragraph();
            pf.ChildElements.Add(new Label() { Text = "Footer Text" });
            pf.ChildElements.Add(new Label() { IsPageNumber = true });
            footer.ChildElements.Add(pf);
            doc.Footers.Add(footer);
        }
    }
}