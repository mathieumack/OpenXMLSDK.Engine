using System;
using System.Collections.Generic;
using System.IO;
using ReportEngine.Core.DataContext.FluentExtensions;
using ReportEngine.Core.DataContext.Charts;
using ReportEngine.Core.DataContext;
using CUsing = ReportEngine.Core.Template.Charts;
using ReportEngine.Core.Template.ExtendedModels;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Images;
using ReportEngine.Core.Template.Text;
using ReportEngine.Core.Template.Tables;
using ReportEngine.Core.Template.Tables.Models;
using ReportEngine.Core.Template.Styles;
using ReportEngine.Core.Template.Charts;

namespace SampleTest.Content
{
    public static class ReportEngineSample
    {
        /// <summary>
        /// Generate the context for the generated template
        /// </summary>
        /// <returns></returns>
        public static ContextModel GetContext()
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
                        .AddString("#Label#", "Label 1");

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

            context.AddItem("#OldBarGraphSampleData#", new BarChartModel()
            {
                BarChartContent = new ReportEngine.Core.DataContext.Charts.BarModel()
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

            return context;
        }

        /// <summary>
        /// Generate the template
        /// </summary>
        /// <returns></returns>
        public static Document GetTemplateDocument()
        {
            var doc = new Document()
            {
                FontName = "Times-Roman",
                FontEncoding = "Cp1252",
                FontSize = 24
            };

            doc.Styles.Add(new Style() { StyleId = "Red", FontColor = "FF0050", FontSize = 42 });
            doc.Styles.Add(new Style() { StyleId = "Yellow", FontColor = "FFFF00", FontSize = 40 });

            var page1 = new Page();
            page1.Margin = new SpacingModel() { Top = 845, Bottom = 1418, Left = 567, Right = 567, Header = 709, Footer = 709 };
            var page2 = new Page();
            page2.Margin = new SpacingModel() { Top = 1418, Left = 845, Header = 709, Footer = 709 };
            doc.Pages.Add(page1);
            doc.Pages.Add(page2);

            // Template 1 :

            var paragraph = new Paragraph()
            {
                Indentation = new ParagraphIndentationModel()
                {
                    Left = 300,
                    Right = 6000
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
            };
            paragraph.ChildElements.Add(new Label() { Text = "Label without special character (éèàù).", FontSize = 30 });
            paragraph.ChildElements.Add(new Hyperlink()
            {
                Text = new Label()
                {
                    Text = "Go to github.",
                    FontSize = 20
                },
                WebSiteUri = "https://www.github.com/"
            });
            paragraph.ChildElements.Add(new Label() { Text = "Ceci est un texte avec accents (éèàù)", FontSize = 22 });
            paragraph.ChildElements.Add(new Label()
            {
                Text = "#KeyTest1#",
                FontSize = 40,
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

            var templateDefinition = new TemplateDefinition()
            {
                TemplateId = "Template 1",
                Note = "Sample paragraph",
                ChildElements = new List<BaseElement>() { paragraph }
            };
            doc.TemplateDefinitions.Add(templateDefinition);

            page1.ChildElements.Add(new TemplateModel() { TemplateId = "Template 1" });
            page1.ChildElements.Add(paragraph);
            page1.ChildElements.Add(new TemplateModel() { TemplateId = "Template 1" });
            page1.ChildElements.Add(new TemplateModel() { TemplateId = "Template 1" });

            var p2 = new Paragraph();
            p2.Shading = "#ParagraphShading#";
            p2.ChildElements.Add(new Label() { Text = "   texte paragraph2 avec espace avant", FontSize = 20, SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve });
            p2.ChildElements.Add(new Label() { Text = "texte2 paragraph2 avec espace après   ", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve });
            p2.ChildElements.Add(new Label() { Text = "   texte3 paragraph2 avec espace avant et après   ", SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve });
            page1.ChildElements.Add(p2);

            var table = new Table()
            {
                TableWidth = new TableWidthModel() { Width = 5000, Type = TableWidthUnitValues.Pct },
                TableIndentation = new TableIndentation() { Width = 1000 },
                ColsWidth = new int[] { 2500, 2500 },
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
                                        Path = @"Resources\Desert.jpg",
                                        ImagePartType = ImagePartType.Jpeg
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
                                        Path = @"Resources\Desert.jpg",
                                        ImagePartType = ImagePartType.Jpeg
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
                                Fusion = false,
                                FusionChild = false,
                                ChildElements = new List<BaseElement>()
                                {
                                    new Label() { Text = "cellule4" }
                                }
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
                                MaxHeight = 250,
                                MaxWidth = 250,
                                Path = @"Resources\Desert.jpg",
                                ImagePartType = ImagePartType.Jpeg
                            }
                        }
                    }
                );

            var tableDataSource = new Table()
            {
                TableWidth = new TableWidthModel() { Width = 5000, Type = TableWidthUnitValues.Pct },
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

            page1.ChildElements.Add(tableDataSource);

            var tableDataSourceWithPrefix = new Table()
            {
                TableWidth = new TableWidthModel() { Width = 5000, Type = TableWidthUnitValues.Pct },
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
            
            page1.ChildElements.Add(tableDataSourceWithPrefix);

            // page 2
            var p21 = new Paragraph();
            p21.Justification = JustificationValues.Center;
            p21.ParagraphStyleId = "Red";
            p21.ChildElements.Add(new Label() { Text = "texte page2" });
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
            p223.ChildElements.Add(new Label() { Text = "Texte paragraph2 avec espace avant", FontSize = 20, SpaceProcessingModeValue = SpaceProcessingModeValues.Preserve });
            foreachPage.ChildElements.Add(p223);
            doc.Pages.Add(foreachPage);

            // page 3
            var page3 = new Page();
            var p31 = new Paragraph() { FontColor = "FF0000", FontSize = 26 };
            p31.ChildElements.Add(new Label() { Text = "Test the HeritFromParent" });
            var p311 = new Paragraph() { FontSize = 16 };
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
            paragraph.ChildElements.Add(new Label() { Text = "Ceci est un test de paragraph avec Style", FontSize = 30 });
            page3.ChildElements.Add(paragraph);

            doc.Pages.Add(page3);

            // page 4
            var page4 = new Page();
            //New page to manage UniformGrid:
            var uniformGrid = new UniformGrid()
            {
                DataSourceKey = "#UniformGridSample#",
                ColsWidth = new int[2] { 2500, 2500 },
                TableWidth = new TableWidthModel() { Width = 5000, Type = TableWidthUnitValues.Pct },
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
                TableWidth = new TableWidthModel() { Width = 5000, Type = TableWidthUnitValues.Pct },
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
                                        ImagePartType = ImagePartType.Jpeg
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
                                        ImagePartType = ImagePartType.Jpeg
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

            // page 6 -> BarChart
            var page6 = new Page();

            var oldpr = new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new CUsing.BarModel()
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
            };

            page6.ChildElements.Add(oldpr);

            var pr = new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new CUsing.BarModel()
                    {
                        Title = "Graph test",
                        ShowTitle = true,
                        FontSize = 23,
                        BarChartType = BarChartType.BarChart,
                        BarDirectionValues = BarDirectionValues.Column,
                        BarGroupingValues = BarGroupingValues.PercentStacked,
                        DataSourceKey = "#BarGraphSampleData#",
                        ShowMajorGridlines = true,
                        MaxHeight = 320
                    }
                }
            };

            page6.ChildElements.Add(pr);

            var singleStackedBarGraph = new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new CUsing.BarModel()
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
            };

            page6.ChildElements.Add(singleStackedBarGraph);

            var singleStackedBarGraphWithMinMax = new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new CUsing.BarModel()
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
            };

            page6.ChildElements.Add(singleStackedBarGraphWithMinMax);

            doc.Pages.Add(page6);

            // page 7
            var page7 = new Page();

            var tableDataSourceWithCellFusion = new Table()
            {
                TableWidth = new TableWidthModel() { Width = 5000, Type = TableWidthUnitValues.Pct },
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

            page7.ChildElements.Add(tableDataSourceWithCellFusion);

            doc.Pages.Add(page7);

            // page 8
            var page8 = new Page();
            var p8 = new Paragraph() { FontColor = "FF0000", FontSize = 26 };
            p8.ChildElements.Add(new Label() { Text = "Label with" + Environment.NewLine + Environment.NewLine + "A new line" });
            page8.ChildElements.Add(p8);

            doc.Pages.Add(page8);

            // page 9 -> PieChart
            var page9 = new Page();

            var pieChartPr = new Paragraph()
            {
                ChildElements = new List<BaseElement>() {
                    new CUsing.PieModel()
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
                        //DataLabelColor = "#000000"//Black
                    }
                }
            };

            page9.ChildElements.Add(pieChartPr);

            // Substitutable string
            var pargraphTitle = new Paragraph
            {
                Justification = JustificationValues.Center,
                ParagraphStyleId = "Red"
            };
            pargraphTitle.ChildElements.Add(new Label() { Text = "Substitutable string" });
            page9.ChildElements.Add(pargraphTitle);

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

            page9.ChildElements.Add(substitutableTableDataSource);

            doc.Pages.Add(page9);

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
                    ImagePartType = ImagePartType.Jpeg
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
    }
}