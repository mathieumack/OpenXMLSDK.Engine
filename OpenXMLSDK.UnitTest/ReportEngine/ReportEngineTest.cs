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

            ContextModel context = new ContextModel()
                        .AddBoolean("#NoRow#", false)
                        .AddString("#ParagraphShading#", "00FF00")
                        .AddString("#ParagraphBorderColor#", "105296")
                        .AddString("#BorderColor#", "00FF00")
                        .AddString("#KeyTest1#", "Key 1")
                        .AddString("#KeyTest2#", "Key 2")
                        .AddBoolean("#BoldKey#", true)
                        .AddString("#FontColorTestRed#", "993333")
                        .AddString("#ParagraphStyleIdTestYellow#", "Yellow")
                        .AddCollection("#Datasource#", row1, row2)
                        .AddCollection("#DatasourcePrefix#", row1, row2, row3, row4);

            // For each with template model
            context.AddItem("#ForEachParagraph#", new DataSourceModel()
            {
                Items = new List<ContextModel>()
                {
                    new ContextModel().AddString("#TemplateKey#", "Template 1").AddString("#KeyTest1#", "foreach"),
                    new ContextModel().AddString("#TemplateKey#", "Template 1").AddString("#KeyTest1#", "foreach")
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

            byte[] numbers = { 0, 16, 104, 213 };

            string textToDisplay = "Base64ContentModel : {0}\n BooleanModel : {1}\n ByteContentModel : {2}\n DateTimeModel : {3}\n DoubleModel : {4}\n StringModel : {5}\n";
            ContextModel rowSubstitutable = new ContextModel();
            rowSubstitutable.AddItem("#SubstitutableStringData#",
                new SubstitutableStringModel(
                    textToDisplay,
                    new ContextModel()
                        .AddBase64Content("#Val1#", "OBFZDTcPCxlCKhdXCQ0kMQhKPh9uIgYIAQxALBtZAwUeOzcdcUEeW0dMO1kbPElWCV1ISFFKZ0kdWFlLAURPZhEFQVseXVtPOUUICVhMAzcfZ14AVEdIVVgfAUIBWVpOUlAeaUVMXFlKIy9rGUN0VF08Oz1POxFfTCcVFw1LMQNbBQYWAQ==")
                        .AddBoolean("#Val2#", false)
                        .AddByteContent("#Val3#", numbers)
                        .AddDateTime("#Val4#", DateTime.Now, null)
                        .AddDouble("#Val5#", 5.4, null)
                        .AddString("#Val6#", "TestString")
                )
            );
            rowSubstitutable.AddItem("#SubstitutableStringDataWithLessParameters#",
                new SubstitutableStringModel(
                    textToDisplay,
                    new ContextModel()
                        .AddBase64Content("#Val1#", "OBFZDTcPCxlCKhdXCQ0kMQhKPh9uIgYIAQxALBtZAwUeOzcdcUEeW0dMO1kbPElWCV1ISFFKZ0kdWFlLAURPZhEFQVseXVtPOUUICVhMAzcfZ14AVEdIVVgfAUIBWVpOUlAeaUVMXFlKIy9rGUN0VF08Oz1POxFfTCcVFw1LMQNbBQYWAQ==")
                        .AddBoolean("#Val2#", false)
                        .AddByteContent("#Val3#", numbers)
                        .AddDateTime("#Val4#", DateTime.Now, null)
                )
            );
            rowSubstitutable.AddItem("#SubstitutableStringDataWithMoreParameters#",
                new SubstitutableStringModel(
                    textToDisplay,
                    new ContextModel()
                        .AddBase64Content("#Val1#", "OBFZDTcPCxlCKhdXCQ0kMQhKPh9uIgYIAQxALBtZAwUeOzcdcUEeW0dMO1kbPElWCV1ISFFKZ0kdWFlLAURPZhEFQVseXVtPOUUICVhMAzcfZ14AVEdIVVgfAUIBWVpOUlAeaUVMXFlKIy9rGUN0VF08Oz1POxFfTCcVFw1LMQNbBQYWAQ==")
                        .AddBoolean("#Val2#", false)
                        .AddByteContent("#Val3#", numbers)
                        .AddDateTime("#Val4#", DateTime.Now, null)
                        .AddDouble("#Val5#", 5.4, null)
                        .AddString("#Val6#", "TestString")
                        .AddDouble("#Val7#", 5.4, null)
                        .AddString("#Val8#", "TestString")
                )
            );

            context.AddCollection("#SubstitutableStringDataSourceModel#", rowSubstitutable);

            GeneratePieChartContext(context);

            GenerateBarGraphContext(context);

            GenerateLineGraphContext(context);

            GenerateScatterGraphContext(context);

            GenerateCombineGraphContext(context);

            return context;
        }

        private static Paragraph CreateTableofContentItem()
        {
            return new Paragraph
            {
                ParagraphStyleId = "TableOfContent",
                SpacingAfter = 0,
                SpacingBefore = 0,
                ChildElements = new List<BaseElement>()
                {
                    new Label()
                    {
                        Text ="Table of content simply",
                        SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve
                    }
                }
            };
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
            var doc = new Document();
            doc.Styles.Add(new Style() { StyleId = "Red", FontColor = "FF0050", FontSize = "42" });
            doc.Styles.Add(new Style() { StyleId = "Yellow", FontColor = "FFFF00", FontSize = "40" });
            doc.Styles.Add(new Style()
            {
                StyleId = "TableOfContent",
                Type = StyleValues.Paragraph,
                CustomStyle = true,
                FontName = "Arial",
                FontSize = "20",
                //PrimaryStyle = true,
                FontColor = "FFFF00",
            });
            doc.Styles.Add(new Style()
            {
                StyleId = "TOC 1",
                Type = StyleValues.Paragraph,
                PrimaryStyle = false,
                CustomStyle = false,
                FontName = "Arial",
                FontSize = "30",
                FontColor = "FF0050",
            });

            var page1 = new Page();
            page1.Margin = new SpacingModel() { Top = 845, Bottom = 1418, Left = 567, Right = 567, Header = 709, Footer = 709 };
            var page2 = new Page();
            page2.Margin = new SpacingModel() { Top = 1418, Left = 845, Header = 709, Footer = 709 };
            doc.Pages.Add(page1);
            doc.Pages.Add(page2);

            // Template 1 :
            page1.ChildElements.Add(
                new Paragraph
                {
                    SpacingAfter = 0,
                    SpacingBefore = 0,
                    ChildElements = new List<BaseElement>
                    {
                           new SimpleField()
                            {
                                Instruction = @"TOC \t TableOfContent;1;",
                                IsDirty = true,
                                HintText = new Label()
                                {
                                    FontColor = "0000FF",
                                    Text = "Default text"
                                }
                            }
                    }
                }
            );

            page1.ChildElements.Add(
                new Paragraph
                {
                    ChildElements = new List<BaseElement>
                    {
                       new Hyperlink(){Anchor = "bmk", Text = new Label(){Text = "link to bookmark with Page Ref : ", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve } },
                       PageCrossReference("PAGEREF bmk")
                    }
                }
            );

            page1.ChildElements.Add(CreateTableofContentItem());

            var paragraph = new Paragraph();
            paragraph.ChildElements.Add(new Label() { Text = "Label without special character (éèàù).", FontSize = "30", FontName = "Arial" });
            paragraph.ChildElements.Add(new Hyperlink()
            {
                Text = new Label()
                {
                    Text = "Go to github.",
                    FontSize = "20",
                    FontName = "Arial"
                },
                WebSiteUri = "https://www.github.com/"
            });
            paragraph.Indentation = new ParagraphIndentationModel()
            {
                Left = "300",
                Right = "6000"
            };
            paragraph.ChildElements.Add(new Label() { Text = "Ceci est un texte avec accents (éèàù)", FontSize = "30", FontName = "Arial" });
            paragraph.ChildElements.Add(new Label()
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
            });
            paragraph.ChildElements.Add(new Label()
            {
                Text = "#KeyTest2#",
                Show = false
            });
            paragraph.Borders = new BorderModel()
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
            };

            var templateDefinition = new TemplateDefinition()
            {
                TemplateId = "Template 1",
                Note = "Sample paragraph",
                ChildElements = new List<BaseElement>() { paragraph }
            };
            doc.TemplateDefinitions.Add(templateDefinition);

            page1.ChildElements.Add(paragraph);
            page1.ChildElements.Add(new TemplateModel() { TemplateId = "Template 1" });

            // Foreach with template model
            var forEach = new ForEach()
            {
                DataSourceKey = "#ForEachParagraph#",
                ItemTemplate = new List<BaseElement>()
                {
                    new TemplateModel() { TemplateId = "#TemplateKey#" }
                }
            };

            page1.ChildElements.Add(forEach);

            page1.ChildElements.Add(new Paragraph()
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

            page1.ChildElements.Add(new Paragraph()
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

            page1.ChildElements.Add(new Paragraph()
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

            page1.ChildElements.Add(new Paragraph()
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
                                        ImagePartType = OpenXMLSDK.Engine.Packaging.ImagePartType.Jpeg
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
                                        ImagePartType = OpenXMLSDK.Engine.Packaging.ImagePartType.Jpeg
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

            table.Borders = new BorderModel()
            {
                BorderPositions = BorderPositions.BOTTOM | BorderPositions.INSIDEVERTICAL,
                BorderWidthBottom = 50,
                BorderWidthInsideVertical = 1,
                UseVariableBorders = true,
                BorderColor = "FF0000"
            };

            page1.ChildElements.Add(table);
            page1.ChildElements.Add(new Paragraph());

            if (File.Exists(@"Resources\Desert.jpg"))
                page1.ChildElements.Add(
                    new Paragraph()
                    {
                        ChildElements = new List<BaseElement>()
                        {
                        new Image()
                        {
                            MaxHeight = 100,
                            MaxWidth = 100,
                            Path = @"Resources\Desert.jpg",
                            ImagePartType = OpenXMLSDK.Engine.Packaging.ImagePartType.Jpeg
                        }
                        }
                    }
                );

            var tableDataSource = new Table()
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
                DataSourceKey = "#Datasource#"
            };

            var tableDataSourceWithPrefix = new Table()
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

            // page 2
            var p21 = new Paragraph();
            p21.Justification = JustificationValues.Center;
            p21.ParagraphStyleId = "Red";
            p21.ChildElements.Add(new Label() { Text = "texte page2", FontName = "Arial" });
            page2.ChildElements.Add(p21);

            var p22 = new Paragraph();
            p22.SpacingBefore = 800;
            p22.SpacingAfter = 800;
            p22.Justification = JustificationValues.Both;
            p22.ParagraphStyleId = "Yellow";
            p22.ChildElements.Add(new Label() { Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse urna augue, convallis eu enim vitae, maximus ultrices nulla. Sed egestas volutpat luctus. Maecenas sodales erat eu elit auctor, eu mattis neque maximus. Duis ac risus quis sem bibendum efficitur. Vivamus justo augue, molestie quis orci non, maximus imperdiet justo. Donec condimentum rhoncus est, ut varius lorem efficitur sed. Donec accumsan sit amet nisl vel ornare. Duis aliquet urna eu mauris porttitor facilisis. " });
            page2.ChildElements.Add(p22);

            var p23 = new Paragraph();
            p23.Borders = new BorderModel()
            {
                BorderPositions = (BorderPositions)13,
                BorderWidth = 20,
                BorderColor = "#ParagraphBorderColor#"
            };
            p23.SpacingBetweenLines = 360;
            p23.ChildElements.Add(new Label() { Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse urna augue, convallis eu enim vitae, maximus ultrices nulla. Sed egestas volutpat luctus. Maecenas sodales erat eu elit auctor, eu mattis neque maximus. Duis ac risus quis sem bibendum efficitur. Vivamus justo augue, molestie quis orci non, maximus imperdiet justo. Donec condimentum rhoncus est, ut varius lorem efficitur sed. Donec accumsan sit amet nisl vel ornare. Duis aliquet urna eu mauris porttitor facilisis. " });
            page2.ChildElements.Add(p23);

            // Adding a foreach page :
            var foreachPage = new ForEachPage();
            foreachPage.DataSourceKey = "#DatasourceTableFusion#";

            foreachPage.Margin = new SpacingModel() { Top = 1418, Left = 845, Header = 709, Footer = 709 };
            var paragraph21 = new Paragraph();
            paragraph21.ChildElements.Add(new Label() { Text = "Page label : #Label#" });
            foreachPage.ChildElements.Add(paragraph21);
            var p223 = new Paragraph();
            p223.Shading = "#ParagraphShading#";
            p223.ChildElements.Add(new Label() { Text = "Texte paragraph2 avec espace avant", FontSize = "20", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve });
            foreachPage.ChildElements.Add(p223);
            doc.Pages.Add(foreachPage);

            // page 3
            var page3 = new Page();
            var p31 = new Paragraph() { FontColor = "FF0000", FontSize = "26" };
            p31.ChildElements.Add(new Label() { Text = "Test the HeritFromParent" });
            var p311 = new Paragraph() { FontSize = "16" };
            p311.ChildElements.Add(new Label() { Text = " Success (not the same size)" });
            p31.ChildElements.Add(p311);
            page3.ChildElements.Add(p31);

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
            page3.ChildElements.Add(tableOfContents);

            paragraph = new Paragraph()
            {
                ParagraphStyleId = "#ParagraphStyleIdTestYellow#"
            };
            paragraph.ChildElements.Add(new Label() { Text = "Ceci est un test de paragraph avec Style", FontSize = "30", FontName = "Arial" });
            page3.ChildElements.Add(paragraph);

            doc.Pages.Add(page3);

            // page 4
            var page4 = new Page();
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
                Borders = new BorderModel()
                {
                    BorderPositions = BorderPositions.BOTTOM | BorderPositions.INSIDEVERTICAL,
                    BorderWidthBottom = 50,
                    BorderWidthInsideVertical = 1,
                    UseVariableBorders = true,
                    BorderColor = "FF0000"
                }
            };

            page4.ChildElements.Add(uniformGrid);

            doc.Pages.Add(page4);

            // page 5
            var page5 = new Page();
            var tableDataSourceWithBeforeAfter = new Table()
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
                                    new Paragraph() { ChildElements = new List<BaseElement>() { new Label() { Text = "Cell 1 - A small paragraph" } }, ParagraphStyleId = "Yellow" },
                                    new Image()
                                    {
                                        MaxHeight = 100,
                                        MaxWidth = 100,
                                        Path = @"Resources\Desert.jpg",
                                        ImagePartType = OpenXMLSDK.Engine.Packaging.ImagePartType.Jpeg
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
                                        Path = @"Resources\Desert.jpg",
                                        ImagePartType = OpenXMLSDK.Engine.Packaging.ImagePartType.Jpeg
                                    },
                                    new Label() { Text = "Cell 2 - an other other label" }
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

            page5.ChildElements.Add(tableDataSourceWithBeforeAfter);

            doc.Pages.Add(page5);

            // page 6
            var page6 = new Page();

            var tableDataSourceWithCellFusion = new Table()
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
                DataSourceKey = "#DatasourceTableFusion#"
            };

            page6.ChildElements.Add(tableDataSourceWithCellFusion);

            doc.Pages.Add(page6);

            // page 7
            var page7 = new Page();
            var p7 = new Paragraph() { FontColor = "FF0000", FontSize = "26" };
            p7.ChildElements.Add(new Label() { Text = "Label with" + Environment.NewLine + Environment.NewLine + "A new line" });
            page7.ChildElements.Add(p7);

            page7.ChildElements.Add(
                new Paragraph
                {
                    ChildElements = new List<BaseElement>
                        {
                           new Label()
                           {
                               Text = "Page Ref bookmark",
                           },
                           new BookmarkStart() {Id = "bmk", Name = "bmk" },
                           new BookmarkEnd(){Id = "bmk"}
                        }
                }
            );

            // Substitutable string
            var pargraphTitle = new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red"
            };
            pargraphTitle.ChildElements.Add(new Label() { Text = "Substitutable string", FontName = "Arial" });
            page7.ChildElements.Add(pargraphTitle);

            var substitutableTableDataSource = new Table()
            {
                RowModel = new Row()
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            ChildElements = new List<BaseElement>()
                            {
                                new Label() { Text = "Matching of supplied parameters and expected parameters : \n"
                                            , Bold = true, Underline = new UnderlineModel () { Val = UnderlineValues.Single } },
                                new Label() { Text = "#SubstitutableStringData#" },

                                new Label() { Text = "\n" },
                                new Label() { Text = "Less supplied parameters than expected parameters : \n"
                                            , Bold = true, Underline = new UnderlineModel () { Val = UnderlineValues.Single } },
                                new Label() { Text = "#SubstitutableStringDataWithLessParameters#" },

                                new Label() { Text = "\n" },
                                new Label() { Text = "More supplied parameters than expected parameters : \n"
                                            , Bold = true, Underline = new UnderlineModel () { Val = UnderlineValues.Single } },
                                new Label() { Text = "#SubstitutableStringDataWithMoreParameters#" }
                            }
                        }
                    }
                }
                ,
                DataSourceKey = "#SubstitutableStringDataSourceModel#"
            };

            page7.ChildElements.Add(substitutableTableDataSource);

            doc.Pages.Add(page7);

            // page 8 -> PieChart
            doc.Pages.Add(GeneratePieChartPage());

            // page 9 -> BarChart
            doc.Pages.Add(GenerateBarChartPage());

            // Page 10 Curve graphs
            doc.Pages.Add(GenerateLineGraphsPage());

            // Page 11 Scatter graphs
            doc.Pages.Add(GenerateScatterGraphsPage());

            // Page 12 Combine graphs (Line and Bar)
            doc.Pages.Add(GenerateCombineGraphsPage());

            // Header
            var header = new Header();
            header.Type = HeaderFooterValues.Default;
            var ph = new Paragraph();
            ph.ChildElements.Add(new Label() { Text = "Header Text" });
            if (File.Exists(@"Resources\Desert.jpg"))
                ph.ChildElements.Add(new Image()
                {
                    MaxHeight = 100,
                    MaxWidth = 100,
                    Path = @"Resources\Desert.jpg",
                    ImagePartType = OpenXMLSDK.Engine.Packaging.ImagePartType.Jpeg
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

            return doc;
        }

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
                        ShowMajorGridlines = true,
                        DataLabel = new DataLabelModel()
                        {
                            //ShowDataLabel = true,
                            ShowCatName = true,
                            ShowPercent = true,
                            Separator = "\n",
                            FontSize = 8
                        }
                        ,
                        DataLabelColor = "#FFFF00"//Yellow
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
                    CategoryType = CategoryTypes.NumberReference,
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

            page.ChildElements.Add(new Paragraph()
            {
                ChildElements = new List<BaseElement>()
                {
                    new Engine.Word.ReportEngine.Models.Charts.BarModel()
                    {
                        Title = "Graph test",
                        ShowTitle = true,
                        ShowBarBorder = true,
                        BarChartType = BarChartType.BarChart,
                        BarDirectionValues = BarDirectionValues.Column,
                        BarGroupingValues = BarGroupingValues.PercentStacked,
                        DataSourceKey = "#OldBarGraphSampleData#",
                        ShowMajorGridlines = true,
                        MaxHeight = 320
                    }
                }
            });

            page.ChildElements.Add(new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new Engine.Word.ReportEngine.Models.Charts.BarModel()
                    {
                        Title = "Graph test",
                        ShowTitle = true,
                        FontSize = "23",
                        BarChartType = BarChartType.BarChart,
                        BarDirectionValues = BarDirectionValues.Column,
                        BarGroupingValues = BarGroupingValues.PercentStacked,
                        DataSourceKey = "#BarGraphSampleData#",
                        ShowMajorGridlines = true,
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
                        DeleteAxeCategory = true,
                        DeleteAxeValue = true,
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
                        DeleteAxeCategory = true,
                        DeleteAxeValue = true,
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
                        CategoryType = CategoryTypes.NumberReference,
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
                    CategoryType = CategoryTypes.NumberReference,
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

        #endregion
    }
}