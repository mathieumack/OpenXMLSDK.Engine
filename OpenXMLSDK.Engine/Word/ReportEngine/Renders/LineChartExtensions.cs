using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.ReportEngine.DataContext.Charts;
using OpenXMLSDK.Engine.Word.Charts;
using OpenXMLSDK.Engine.Word.Extensions;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using DC = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    /// <summary>
    /// Line graph extension
    /// </summary>
    public static class LineChartExtensions
    {
        /// <summary>
        ///  Render the graph
        /// </summary>
        /// <param name="lineModel"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static Run Render(this LineModel lineModel, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(lineModel, formatProvider);

            Run runItem = null;

            // We construct categories and series from the context object
            if (!string.IsNullOrWhiteSpace(lineModel.DataSourceKey) && context.TryGetItem(lineModel.DataSourceKey, out MultipleSeriesChartModel contextModel))
            {
                if (contextModel.ChartContent is null || contextModel.ChartContent.Categories is null
                   || contextModel.ChartContent.Series is null)
                    return runItem;

                // Update categories object :
                lineModel.Categories = contextModel.ChartContent.Categories.Select(e => new LineCategory()
                {
                    Name = e.Name,
                    Color = e.Color
                }).ToList();

                // We update
                lineModel.Series = contextModel.ChartContent.Series.Select(e => new LineSerie()
                {
                    Name = e.Name,
                    Values = e.Values,
                    Color = e.Color,
                    DataLabelColor = e.DataLabelColor,
                    LabelFormatString = e.LabelFormatString,
                    HasBorder = e.HasBorder,
                    BorderColor = e.BorderColor,
                    BorderWidth = e.BorderWidth,
                    UseSecondaryAxis = e.UseSecondaryAxis
                }).ToList();

                // Update Axes
                UpdateAxisFromcontext(lineModel.CategoriesAxisModel, contextModel.ChartContent.CategoriesAxisModel);
                UpdateAxisFromcontext(lineModel.ValuesAxisModel, contextModel.ChartContent.ValuesAxisModel);
                UpdateAxisFromcontext(lineModel.SecondaryValuesAxisModel, contextModel.ChartContent.SecondaryValuesAxisModel);
            }

            runItem = CreateGraph(lineModel, documentPart);

            if (runItem != null)
                parent.AppendChild(runItem);

            return runItem;
        }

        #region Internal methods

        /// <summary>
        /// Update template axis model with context values
        /// </summary>
        /// <param name="template"></param>
        /// <param name="context"></param>
        private static void UpdateAxisFromcontext(ChartAxisModel template, AxisModel context)
        {
            if (!string.IsNullOrWhiteSpace(context.Title))
                template.Title = context.Title;

            if (!string.IsNullOrWhiteSpace(context.Color))
                template.TitleColor = context.Color;
        }

        /// <summary>
        /// Create the graph
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="documentPart"></param>
        /// <returns></returns>
        private static Run CreateGraph(LineModel chartModel, OpenXmlPart documentPart)
        {
            if (chartModel.Categories == null)
                throw new ArgumentNullException("categories of chartModel must not be null");
            if (chartModel.Series == null)
                throw new ArgumentNullException("series of chartModel must be not null");

            if (chartModel.Series.Any(e => e.Values.Count != chartModel.Categories.Count))
                throw new ChartModelException("Error in series. Serie values must have same count as categories.", "004-001");

            // Add a new chart and set the chart language to English-US.
            ChartPart chartPart = documentPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.AppendChild(new EditingLanguage { Val = new StringValue("en-US") });
            chartPart.ChartSpace.AppendChild(new RoundedCorners { Val = new BooleanValue(chartModel.RoundedCorner) });
            Chart chart = chartPart.ChartSpace.AppendChild(new Chart());

            // Add graph title.
            if (chartModel.ShowTitle)
            {
                Title titleChart = chart.AppendChild(new Title());
                titleChart.AppendChild(
                    new ChartText(
                        new RichText(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text(chartModel.Title)))
                        )
                    )
                );
                titleChart.AppendChild(new Overlay() { Val = false });
            }

            // Create a new clustered column chart.
            PlotArea plotArea = chart.AppendChild(new PlotArea());
            plotArea.AppendChild(new Layout());


            uint i = 0;
            ManageLineChart(chartModel, plotArea, new UInt32Value(48650112U), new UInt32Value(48672768U), ref i);
            if (chartModel.Series.Any(s => s.UseSecondaryAxis))
                ManageLineChart(chartModel, plotArea, new UInt32Value(48650108U), new UInt32Value(48672708U), ref i, true);

            ManageLegend(chartModel, chart);

            chart.Append(
                new PlotVisibleOnly() { Val = new BooleanValue(true) },
                new DisplayBlanksAs() { Val = new EnumValue<DisplayBlanksAsValues>(DisplayBlanksAsValues.Gap) },
                new ShowDataLabelsOverMaximum() { Val = false });

            ManageGraphBorders(chartModel, chartPart);

            Run element = SaveChart(chartModel, documentPart, chartPart);

            return element;
        }

        /// <summary>
        /// Manage Line chart
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="plotArea"></param>
        /// <param name="categoryAxisId"></param>
        /// <param name="valuesAxisId"></param>
        /// <param name="i"></param>
        /// <param name="secondaryAxis"></param>
        private static void ManageLineChart(LineModel chartModel, PlotArea plotArea, UInt32Value categoryAxisId, UInt32Value valuesAxisId, ref uint i, bool secondaryAxis = false)
        {
            LineChart lineChart = plotArea.AppendChild(
                new LineChart(
                    new Grouping { Val = new EnumValue<DC.GroupingValues>((DC.GroupingValues)(int)chartModel.GroupingValues) },
                    new VaryColors { Val = new BooleanValue(chartModel.VaryColors) }));

            uint p = 0;
            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.
            foreach (var serie in chartModel.Series.Where(s => s.UseSecondaryAxis.Equals(secondaryAxis)))
            {
                // Series.
                LineChartSeries lineChartSeries = lineChart.AppendChild(
                    new LineChartSeries(
                        new Index() { Val = i },
                        new Order() { Val = i },
                        new Marker
                        {
                            Symbol = new Symbol { Val = new EnumValue<DC.MarkerStyleValues>((DC.MarkerStyleValues)(int)serie.LineSerieMarker.MarkerStyleValues) },
                            Size = new Size { Val = serie.LineSerieMarker.Size }
                        },
                        new SeriesText(
                            new StringReference(
                                new StringCache(
                                    new PointCount() { Val = new UInt32Value(1U) },
                                    new StringPoint() { Index = (uint)0, NumericValue = new NumericValue() { Text = serie.Name } })))));

                // Serie color.
                if (!string.IsNullOrWhiteSpace(serie.Color))
                {
                    string color = serie.Color;
                    color = color.Replace("#", "");
                    if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                        throw new Exception("Error in color of serie.");

                    lineChartSeries.AppendChild(new ChartShapeProperties(new A.Outline(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } })));
                }

                // Categories.
                StringReference strLit = lineChartSeries.AppendChild(new CategoryAxisData()).AppendChild(new StringReference());
                strLit.AppendChild(new StringCache());
                strLit.StringCache.AppendChild(new PointCount() { Val = (uint)chartModel.Categories.Count });
                // Category list.
                foreach (var categorie in chartModel.Categories)
                {
                    strLit.StringCache.AppendChild(new StringPoint() { Index = p, NumericValue = new NumericValue(categorie.Name) });
                    p++;
                }
                p = 0;

                // Values.
                NumberReference numLit = lineChartSeries.AppendChild(new Values()).AppendChild(new NumberReference());
                numLit.AppendChild(new NumberingCache());
                numLit.NumberingCache.AppendChild(new FormatCode("General"));
                numLit.NumberingCache.AppendChild(new PointCount() { Val = (uint)serie.Values.Count });
                foreach (var value in serie.Values)
                {
                    numLit.NumberingCache.AppendChild(new NumericPoint() { Index = p, NumericValue = new NumericValue(value != null ? value.ToString() : string.Empty) });
                    p++;
                }
                i++;
            }

            ManageDataLabels(chartModel, lineChart);

            if (chartModel.SpaceBetweenLineCategories.HasValue)
                lineChart.AppendChild(new GapWidth() { Val = (ushort)chartModel.SpaceBetweenLineCategories.Value });
            else
                lineChart.AppendChild(new GapWidth() { Val = 55 });

            lineChart.AppendChild(new Overlap() { Val = 100 });

            lineChart.AppendChild(new AxisId() { Val = categoryAxisId });
            lineChart.AppendChild(new AxisId() { Val = valuesAxisId });

            // Add the Category Axis.
            var catAxis = new CategoryAxis(
                new AxisId() { Val = categoryAxisId },
                new Scaling() { Orientation = new Orientation() { Val = new EnumValue<OrientationValues>(OrientationValues.MinMax) } },
                new Delete() { Val = secondaryAxis || chartModel.CategoriesAxisModel.DeleteAxis },
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                new MajorTickMark() { Val = TickMarkValues.None },
                new MinorTickMark() { Val = TickMarkValues.None },
                new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = valuesAxisId },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new AutoLabeled() { Val = new BooleanValue(true) },
                new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                new LabelOffset() { Val = new UInt16Value((ushort)100) },
                new NoMultiLevelLabels() { Val = false },
                new MajorGridlines(ManageShapeProperties(chartModel.CategoriesAxisModel.ShowMajorGridlines, chartModel.CategoriesAxisModel.MajorGridlinesColor)),
                ManageShapeProperties(chartModel.CategoriesAxisModel.ShowAxisCurve, chartModel.CategoriesAxisModel.AxisCurveColor));
            if (!secondaryAxis && !string.IsNullOrWhiteSpace(chartModel.CategoriesAxisModel.Title))
                catAxis.Title = ManageTitle(chartModel.CategoriesAxisModel.Title, chartModel.CategoriesAxisModel.TitleColor);
            plotArea.AppendChild(catAxis);

            // Add the Value Axis.
            var axixModel = secondaryAxis ? chartModel.SecondaryValuesAxisModel : chartModel.ValuesAxisModel;
            var valueAxis = new ValueAxis(
                new AxisId() { Val = valuesAxisId },
                chartModel.ValuesAxisScaling?.GetScaling() ?? new Scaling() { Orientation = new Orientation() { Val = new EnumValue<OrientationValues>(OrientationValues.MinMax) } },
                new Delete() { Val = axixModel.DeleteAxis },
                new AxisPosition() { Val = secondaryAxis ? new EnumValue<AxisPositionValues>(AxisPositionValues.Right) : new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                new DC.NumberingFormat()
                {
                    FormatCode = new StringValue("General"),
                    SourceLinked = new BooleanValue(true)
                },
                new MajorTickMark() { Val = TickMarkValues.None },
                new MinorTickMark() { Val = TickMarkValues.None },
                new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = categoryAxisId },
                new Crosses() { Val = secondaryAxis ? new EnumValue<CrossesValues>(CrossesValues.Maximum) : new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) },
                ManageShapeProperties(axixModel.ShowAxisCurve, axixModel.AxisCurveColor));
            if (!secondaryAxis)
                valueAxis.MajorGridlines = new MajorGridlines
                {
                    ChartShapeProperties = ManageShapeProperties(axixModel.ShowMajorGridlines, axixModel.MajorGridlinesColor)
                };
            if (!string.IsNullOrWhiteSpace(axixModel.Title))
                valueAxis.Title = ManageTitle(axixModel.Title, axixModel.TitleColor);
            plotArea.AppendChild(valueAxis);
        }

        /// <summary>
        /// Manage DataLabels
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="lineChart"></param>
        private static void ManageDataLabels(LineModel chartModel, LineChart lineChart)
        {
            DataLabels dLbls = new DataLabels(
                new ShowLegendKey() { Val = false },
                new ShowValue() { Val = chartModel.DataLabel != null && chartModel.DataLabel.ShowDataLabel },
                new ShowCategoryName() { Val = false },
                new ShowSeriesName() { Val = false },
                new ShowPercent() { Val = false },
                new ShowBubbleSize() { Val = false });

            // DataLabel
            string dataLabelColor = "#000000"; //Black by default
            if (!string.IsNullOrWhiteSpace(chartModel.DataLabelColor))
                dataLabelColor = chartModel.DataLabelColor;
            dataLabelColor = dataLabelColor.Replace("#", "");
            if (!Regex.IsMatch(dataLabelColor, "^[0-9-A-F]{6}$"))
                throw new Exception("Error in dataLabel color.");

            var fontSize = chartModel.DataLabel.FontSize * 100; // word size x 100 for XML FontSize
            TextProperties txtPr = new TextProperties
            (
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph
                (
                    new A.ParagraphProperties
                    (
                        new A.DefaultRunProperties
                        (
                            new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = dataLabelColor } }
                        )
                        { Baseline = 0, FontSize = fontSize }
                    )
                )
            );

            dLbls.AppendChild(txtPr);
            lineChart.AppendChild(dLbls);
        }

        /// <summary>
        ///  Manage axes titles
        /// </summary>
        /// <param name="title"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        private static Title ManageTitle(string title, string color)
        {
            Title titleElement = new Title();

            if (string.IsNullOrWhiteSpace(title))
                return titleElement;

            var rpr = new A.RunProperties
            {
                FontSize = 1000
            };

            if (!string.IsNullOrWhiteSpace(color))
            {
                color = color.Replace("#", "");
                if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                    throw new Exception("Error in color of serie.");

                rpr.AppendChild(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } });
            }

            var paragraph = new A.Paragraph();
            paragraph.AppendChild(new A.Run
            {
                Text = new A.Text(title),
                RunProperties = rpr
            });

            titleElement.AppendChild(
                new ChartText
                {
                    RichText = new RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        paragraph)
                });
            titleElement.AppendChild(new Overlay() { Val = false });

            return titleElement;
        }

        /// <summary>
        /// Manage ShapeProperties
        /// </summary>
        /// <param name="show"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        private static ChartShapeProperties ManageShapeProperties(bool show, string color)
        {
            if (!show)
                return new ChartShapeProperties(new A.Outline(new A.NoFill()));

            if (!string.IsNullOrWhiteSpace(color))
            {
                color = color.Replace("#", "");
                if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                    throw new Exception("Error in color of grid lines.");
                return new ChartShapeProperties(new A.Outline(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } }));
            }

            return new ChartShapeProperties();
        }

        /// <summary>
        /// Manage legend
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="chart"></param>
        private static void ManageLegend(LineModel chartModel, Chart chart)
        {
            // Add the chart Legend.
            if (chartModel.ShowLegend)
            {
                var textProperty = new TextProperties();
                if (!string.IsNullOrEmpty(chartModel.FontFamilyLegend))
                {
                    textProperty = new TextProperties
                    (
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(new A.LatinFont() { CharacterSet = 0, Typeface = chartModel.FontFamilyLegend }) { Baseline = 0 }))
                    );
                }

                chart.AppendChild(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
                new Overlay() { Val = false },
                new Layout(),
                textProperty));
            }
        }

        /// <summary>
        /// Manage graph borders
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="chartPart"></param>
        private static void ManageGraphBorders(LineModel chartModel, ChartPart chartPart)
        {
            // Graph borders.
            if (chartModel.HasBorder)
            {
                chartModel.BorderWidth = chartModel.BorderWidth.HasValue ? chartModel.BorderWidth.Value : 12700;

                if (!string.IsNullOrEmpty(chartModel.BorderColor))
                {
                    var color = chartModel.BorderColor.Replace("#", "");
                    if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                        throw new Exception("Error in color of chart borders.");
                    chartPart.ChartSpace.AppendChild(new ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = color })) { Width = chartModel.BorderWidth.Value }));
                }
                else
                {
                    chartPart.ChartSpace.AppendChild(new ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" })) { Width = chartModel.BorderWidth.Value }));
                }
            }
            else
            {
                chartPart.ChartSpace.AppendChild(new ChartShapeProperties(new A.Outline(new A.NoFill())));
            }
        }

        /// <summary>
        /// Save the chart
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="documentPart"></param>
        /// <param name="chartPart"></param>
        /// <returns></returns>
        private static Run SaveChart(LineModel chartModel, OpenXmlPart documentPart, ChartPart chartPart)
        {
            // Save the chart part.
            chartPart.ChartSpace.Save();

            // Get the grap Id for the drawing element.
            string relationshipId = documentPart.GetIdOfPart(chartPart);

            // Resize the graph.
            long imageWidth = 5486400;
            long imageHeight = 3200400;

            if (chartModel.MaxWidth.HasValue)
                // convert pixel in EMU (English Metric Unit normalement c'est : EMU = pixel * 914400 / 96) --> 914400 / 96 = 9525.
                imageWidth = (long)chartModel.MaxWidth * 9525;
            if (chartModel.MaxHeight.HasValue)
                imageHeight = (long)chartModel.MaxHeight * 9525;

            // Drawing element creation.
            var element = new Run(
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = imageWidth, Cy = imageHeight },
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value)1U,
                            Name = "Chart 1"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                // Lien avec l'Id du graphique
                                new ChartReference() { Id = relationshipId }
                                )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
                    )
                )
            );
            return element;
        }

        #endregion
    }
}
