using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.Charts;
using OpenXMLSDK.Engine.Word.Extensions;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using dc = DocumentFormat.OpenXml.Drawing.Charts;
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
            if (!string.IsNullOrWhiteSpace(lineModel.DataSourceKey) && context.TryGetItem(lineModel.DataSourceKey, out MultipleSeriesChartModel multipleSeriesContextModel))
            {
                if (multipleSeriesContextModel.ChartContent != null && multipleSeriesContextModel.ChartContent.Categories != null
                   && multipleSeriesContextModel.ChartContent.Series != null)
                {
                    // Update categories object :
                    lineModel.Categories = multipleSeriesContextModel.ChartContent.Categories.Select(e => new LineCategory()
                    {
                        Name = e.Name,
                        Color = e.Color
                    }).ToList();

                    // We update
                    lineModel.Series = multipleSeriesContextModel.ChartContent.Series.Select(e => new LineSerie()
                    {
                        LabelFormatString = e.LabelFormatString,
                        Color = e.Color,
                        DataLabelColor = e.DataLabelColor,
                        Values = e.Values,
                        Name = e.Name,
                        HasBorder = e.HasBorder,
                        BorderColor = e.BorderColor,
                        BorderWidth = e.BorderWidth
                    }).ToList();
                }
                else
                    return runItem;
            }

            runItem = CreateGraph(lineModel, documentPart);

            if (runItem != null)
                parent.AppendChild(runItem);

            return runItem;
        }

        #region Internal methods

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

            int countCategories = chartModel.Categories.Count;

            if (chartModel.Series.Any(e => e.Values.Count != countCategories))
                throw new ChartModelException("Error in series. Serie values must have same count as categories.", "004-001");

            // Add a new chart and set the chart language to English-US.
            ChartPart chartPart = documentPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new dc.ChartSpace();
            chartPart.ChartSpace.Append(new dc.EditingLanguage { Val = new StringValue("en-US") });
            chartPart.ChartSpace.Append(new dc.RoundedCorners { Val = new BooleanValue(chartModel.RoundedCorner) });
            dc.Chart chart = chartPart.ChartSpace.AppendChild(new dc.Chart());

            // Add graph title.
            if (chartModel.ShowTitle)
            {
                dc.Title titleChart = chart.AppendChild(new dc.Title());
                titleChart.AppendChild(new dc.ChartText(new dc.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(chartModel.Title))))));
                titleChart.AppendChild(new dc.Overlay() { Val = false });
            }

            // Create a new clustered column chart.
            dc.PlotArea plotArea = chart.AppendChild(new dc.PlotArea());
            plotArea.AppendChild(new dc.Layout());
            dc.LineChart lineChart = plotArea.AppendChild(
                new dc.LineChart(
                    new dc.Grouping { Val = new EnumValue<dc.GroupingValues>((dc.GroupingValues)(int)chartModel.GroupingValues) },
                    new dc.VaryColors { Val = new BooleanValue(chartModel.VaryColors) }));

            uint i = 0;
            uint p = 0;
            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.
            foreach (var serie in chartModel.Series)
            {
                // Series.
                dc.LineChartSeries lineChartSeries = lineChart.AppendChild(
                    new dc.LineChartSeries(
                        new dc.Index() { Val = i },
                        new dc.Order() { Val = i },
                        new dc.Marker
                        {
                            Symbol = new dc.Symbol { Val = new EnumValue<dc.MarkerStyleValues>((dc.MarkerStyleValues)(int)serie.LineSerieMarker.MarkerStyleValues) },
                            Size = new dc.Size { Val = serie.LineSerieMarker.Size }
                        },
                        new dc.SeriesText(
                            new dc.StringReference(
                                new dc.StringCache(
                                    new dc.PointCount() { Val = new UInt32Value(1U) },
                                    new dc.StringPoint() { Index = (uint)0, NumericValue = new dc.NumericValue() { Text = serie.Name } })))));

                // Serie color.
                A.ShapeProperties shapeProperties = new A.ShapeProperties();

                if (!string.IsNullOrWhiteSpace(serie.Color))
                {
                    string color = serie.Color;
                    color = color.Replace("#", "");
                    if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                        throw new Exception("Error in color of serie.");

                    shapeProperties.AppendChild(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } });
                }

                // Border of all categories.
                if (serie.HasBorder)
                {
                    serie.BorderWidth = serie.BorderWidth.HasValue ? serie.BorderWidth.Value : 12700;

                    serie.BorderColor = !string.IsNullOrEmpty(serie.BorderColor) ? serie.BorderColor : "000000";
                    serie.BorderColor = serie.BorderColor.Replace("#", "");
                    if (!Regex.IsMatch(serie.BorderColor, "^[0-9-A-F]{6}$"))
                        throw new Exception("Error in color of serie.");

                    shapeProperties.AppendChild(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = serie.BorderColor })) { Width = serie.BorderWidth.Value });
                }

                if (shapeProperties.HasChildren)
                    lineChartSeries.AppendChild(shapeProperties);

                // Categories.
                dc.StringReference strLit = lineChartSeries.AppendChild
                        (new dc.CategoryAxisData()).AppendChild(new dc.StringReference());
                strLit.AppendChild(new dc.StringCache());
                strLit.StringCache.AppendChild(new dc.PointCount() { Val = (uint)countCategories });
                // Category list.
                foreach (var categorie in chartModel.Categories)
                {
                    strLit.StringCache.AppendChild(new dc.StringPoint() { Index = p, NumericValue = new dc.NumericValue(categorie.Name) });
                    p++;
                }
                p = 0;

                // Values
                dc.NumberReference numLit = lineChartSeries.AppendChild
                    (new dc.Values())
                        .AppendChild(new dc.NumberReference());
                numLit.AppendChild(new dc.NumberingCache());
                numLit.NumberingCache.AppendChild(new dc.FormatCode("General"));
                numLit.NumberingCache.AppendChild(new dc.PointCount() { Val = (uint)serie.Values.Count });
                foreach (var value in serie.Values)
                {
                    numLit.NumberingCache.AppendChild(new dc.NumericPoint() { Index = p, NumericValue = new dc.NumericValue(value != null ? value.ToString() : string.Empty) });
                    p++;
                }
                i++;
            }

            dc.DataLabels dLbls = new dc.DataLabels(
                new dc.ShowLegendKey() { Val = false },
                new dc.ShowValue() { Val = chartModel.DataLabel == null ? false : chartModel.DataLabel.ShowDataLabel },
                new dc.ShowCategoryName() { Val = false },
                new dc.ShowSeriesName() { Val = false },
                new dc.ShowPercent() { Val = false },
                new dc.ShowBubbleSize() { Val = false });

            // DataLabel
            string dataLabelColor = "#000000"; //Black by default
            if (!string.IsNullOrWhiteSpace(chartModel.DataLabelColor))
                dataLabelColor = chartModel.DataLabelColor;
            dataLabelColor = dataLabelColor.Replace("#", "");
            if (!Regex.IsMatch(dataLabelColor, "^[0-9-A-F]{6}$"))
                throw new Exception("Error in dataLabel color.");

            var fontSize = chartModel.DataLabel.FontSize * 100; // word size x 100 for XML FontSize
            dc.TextProperties txtPr = new dc.TextProperties(
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
            dLbls.Append(txtPr);

            lineChart.Append(dLbls);

            if (chartModel.SpaceBetweenLineCategories.HasValue)
                lineChart.Append(new dc.GapWidth() { Val = (UInt16)chartModel.SpaceBetweenLineCategories.Value });
            else
                lineChart.Append(new dc.GapWidth() { Val = 55 });

            lineChart.Append(new dc.Overlap() { Val = 100 });

            lineChart.Append(new dc.AxisId() { Val = new UInt32Value(48650112U) });
            lineChart.Append(new dc.AxisId() { Val = new UInt32Value(48672768U) });

            // Set ShapeProperties.
            dc.ShapeProperties dcSP = null;
            if (chartModel.ShowMajorGridlines)
            {
                if (!string.IsNullOrWhiteSpace(chartModel.MajorGridlinesColor))
                {
                    string color = chartModel.MajorGridlinesColor;
                    color = color.Replace("#", "");
                    if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                        throw new Exception("Error in color of grid lines.");
                    dcSP = new dc.ShapeProperties(new A.Outline(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } }));
                }
                else
                {
                    dcSP = new dc.ShapeProperties();
                }
            }
            else
            {
                dcSP = new dc.ShapeProperties(new A.Outline(new A.NoFill()));
            }

            // Add the Category Axis.
            plotArea.AppendChild(new dc.CategoryAxis(new dc.AxisId() { Val = new UInt32Value(48650112u) },
                new dc.Scaling() { Orientation = new dc.Orientation() { Val = new EnumValue<dc.OrientationValues>(dc.OrientationValues.MinMax) } },
                new dc.Delete() { Val = chartModel.DeleteAxeCategory },
                new dc.AxisPosition() { Val = new EnumValue<dc.AxisPositionValues>(dc.AxisPositionValues.Left) },
                new dc.MajorTickMark() { Val = dc.TickMarkValues.None },
                new dc.MinorTickMark() { Val = dc.TickMarkValues.None },
                new dc.TickLabelPosition() { Val = new EnumValue<dc.TickLabelPositionValues>(dc.TickLabelPositionValues.NextTo) },
                new dc.CrossingAxis() { Val = new UInt32Value(48672768U) },
                new dc.Crosses() { Val = new EnumValue<dc.CrossesValues>(dc.CrossesValues.AutoZero) },
                new dc.AutoLabeled() { Val = new BooleanValue(true) },
                new dc.LabelAlignment() { Val = new EnumValue<dc.LabelAlignmentValues>(dc.LabelAlignmentValues.Center) },
                new dc.LabelOffset() { Val = new UInt16Value((ushort)100) },
                new dc.NoMultiLevelLabels() { Val = false },
                dcSP));

            // Add the Value Axis.
            plotArea.AppendChild(new dc.ValueAxis(new dc.AxisId() { Val = new UInt32Value(48672768u) },
                chartModel.ValuesAxisScaling?.GetScaling() ??
                new dc.Scaling() { Orientation = new dc.Orientation() { Val = new EnumValue<dc.OrientationValues>(dc.OrientationValues.MinMax) } },
                new dc.Delete() { Val = chartModel.DeleteAxeValue },
                new dc.AxisPosition() { Val = new EnumValue<dc.AxisPositionValues>(dc.AxisPositionValues.Bottom) },
                new dc.NumberingFormat()
                {
                    FormatCode = new StringValue("General"),
                    SourceLinked = new BooleanValue(true)
                },
                new dc.MajorTickMark() { Val = dc.TickMarkValues.None },
                new dc.MinorTickMark() { Val = dc.TickMarkValues.None },
                new dc.TickLabelPosition() { Val = new EnumValue<dc.TickLabelPositionValues>(dc.TickLabelPositionValues.NextTo) },
                new dc.CrossingAxis() { Val = new UInt32Value(48650112U) },
                new dc.Crosses() { Val = new EnumValue<dc.CrossesValues>(dc.CrossesValues.AutoZero) },
                new dc.CrossBetween() { Val = new EnumValue<dc.CrossBetweenValues>(dc.CrossBetweenValues.Between) },
                new dc.MajorGridlines(dcSP.CloneNode(true)),
                dcSP.CloneNode(true)));

            // Add the chart Legend.
            if (chartModel.ShowLegend)
            {
                var textProperty = new dc.TextProperties();
                if (!string.IsNullOrEmpty(chartModel.FontFamilyLegend))
                {
                    textProperty = new dc.TextProperties(new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(new A.LatinFont() { CharacterSet = 0, Typeface = chartModel.FontFamilyLegend }) { Baseline = 0 })));
                }

                dc.Legend legend = chart.AppendChild(new dc.Legend(new dc.LegendPosition() { Val = new EnumValue<dc.LegendPositionValues>(dc.LegendPositionValues.Right) },
                new dc.Overlay() { Val = false },
                new dc.Layout(),
                textProperty));
            }

            chart.Append(new dc.PlotVisibleOnly() { Val = new BooleanValue(true) },
                new dc.DisplayBlanksAs() { Val = new EnumValue<dc.DisplayBlanksAsValues>(dc.DisplayBlanksAsValues.Gap) },
                new dc.ShowDataLabelsOverMaximum() { Val = false });

            // Graph borders.
            if (chartModel.HasBorder)
            {
                chartModel.BorderWidth = chartModel.BorderWidth.HasValue ? chartModel.BorderWidth.Value : 12700;

                if (!string.IsNullOrEmpty(chartModel.BorderColor))
                {
                    var color = chartModel.BorderColor.Replace("#", "");
                    if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                        throw new Exception("Error in color of chart borders.");
                    chartPart.ChartSpace.Append(new dc.ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = color })) { Width = chartModel.BorderWidth.Value }));
                }
                else
                {
                    chartPart.ChartSpace.Append(new dc.ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" })) { Width = chartModel.BorderWidth.Value }));
                }
            }
            else
            {
                chartPart.ChartSpace.Append(new dc.ChartShapeProperties(new A.Outline(new A.NoFill())));
            }

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
                                new dc.ChartReference() { Id = relationshipId }
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
