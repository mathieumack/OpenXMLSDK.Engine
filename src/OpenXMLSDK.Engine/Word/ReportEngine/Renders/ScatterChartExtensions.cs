using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.ReportEngine.DataContext.Charts;
using OpenXMLSDK.Engine.ReportEngine.Validations;
using OpenXMLSDK.Engine.Word.Extensions;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using DC = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class ScatterChartExtensions
    {
        /// <summary>
        /// Graph render
        /// </summary>
        /// <param name="scatterModel"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static Run Render(this ScatterModel scatterModel, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(scatterModel, formatProvider);

            Run runItem = null;

            // We construct categories and series from the context object
            if (!string.IsNullOrWhiteSpace(scatterModel.DataSourceKey) && context.TryGetItem(scatterModel.DataSourceKey, out MultipleSeriesChartModel contextModel))
            {
                if (contextModel.ChartContent is null || contextModel.ChartContent.ScatterSeries is null)
                    return runItem;

                // Series update
                scatterModel.Series = contextModel.ChartContent.ScatterSeries.Select(e => new ScatterSerie()
                {
                    Name = e.Name,
                    Values = e.Values,
                    Color = e.Color,
                    DataLabelColor = e.DataLabelColor,
                    LabelFormatString = e.LabelFormatString,
                    HasBorder = e.HasBorder,
                    UseSecondaryAxis = e.UseSecondaryAxis,
                    SmoothCurve = e.SmoothCurve,
                    PresetLineDashValues = e.PresetLineDashValues,
                    HideCurve = e.HideCurve,
                    SerieMarker = e.SerieMarker
                }).ToList();

                // Update Axes
                UpdateAxisFromcontext(scatterModel.CategoriesAxisModel, contextModel.ChartContent.CategoriesAxisModel);
                UpdateAxisFromcontext(scatterModel.ValuesAxisModel, contextModel.ChartContent.ValuesAxisModel);
                UpdateAxisFromcontext(scatterModel.SecondaryCategoriesAxisModel, contextModel.ChartContent.SecondaryCategoriesAxisModel);
                UpdateAxisFromcontext(scatterModel.SecondaryValuesAxisModel, contextModel.ChartContent.SecondaryValuesAxisModel);
            }

            runItem = CreateGraph(scatterModel, documentPart);

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

            if (context.CrossesAt.HasValue)
                template.CrossesAt = context.CrossesAt;

            if (!string.IsNullOrWhiteSpace(context.LabelFormat))
                template.LabelFormat = context.LabelFormat;

            if (context.InvertAxisOrder.HasValue)
                template.ScalingModel.Orientation = context.InvertAxisOrder.Value ? OrientationType.MaxMin : OrientationType.MinMax;
            if (context.MinimumValue.HasValue)
                template.ScalingModel.MinAxisValue = context.MinimumValue.Value;
            if (context.MaximumValue.HasValue)
                template.ScalingModel.MaxAxisValue = context.MaximumValue.Value;
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
        public static void ManageScatterChart(ScatterModel chartModel, PlotArea plotArea, UInt32Value categoryAxisId, UInt32Value valuesAxisId, ref uint i, bool secondaryAxis = false)
        {
            ScatterChart chart = plotArea.AppendChild(
                new ScatterChart(
                    new ScatterStyle { Val = ScatterStyleValues.Line },
                    new VaryColors { Val = new BooleanValue(chartModel.VaryColors) }));

            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.
            foreach (var serie in chartModel.Series.Where(s => s.UseSecondaryAxis.Equals(secondaryAxis)))
            {
                // Series.
                ScatterChartSeries scatterChartSeries = chart.AppendChild(
                    new ScatterChartSeries(
                        new DC.Index() { Val = i },
                        new Order() { Val = i },
                        new Marker
                        {
                            Symbol = new Symbol { Val = new DC.MarkerStyleValues(serie.SerieMarker.MarkerStyleValues.ToString().ToLower()) },
                            Size = new Size { Val = serie.SerieMarker.Size }
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
                    color.CheckColorFormat();

                    if (serie.HideCurve.HasValue && serie.HideCurve.Value)
                    {
                        scatterChartSeries.AppendChild(new ChartShapeProperties(new A.Outline(new A.NoFill())));
                        scatterChartSeries.Marker.ChartShapeProperties = new ChartShapeProperties(
                                new A.Outline(
                                    new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } },
                                    new A.PresetDash() { Val = new A.PresetLineDashValues(serie.PresetLineDashValues.ToString().ToLower()) }));
                    }
                    else
                        scatterChartSeries.AppendChild(
                            new ChartShapeProperties(
                                new A.Outline(
                                    new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } },
                                    new A.PresetDash() { Val = new A.PresetLineDashValues(serie.PresetLineDashValues.ToString().ToLower()) })));
                }
                else if (serie.HideCurve.HasValue && serie.HideCurve.Value)
                    scatterChartSeries.AppendChild(new ChartShapeProperties(new A.Outline(new A.NoFill())));

                // Values.
                var xValues = scatterChartSeries.AppendChild(new XValues()).AppendChild(new NumberReference());
                xValues.AppendChild(new NumberingCache());
                xValues.NumberingCache.AppendChild(new FormatCode("General"));
                xValues.NumberingCache.AppendChild(new PointCount() { Val = (uint)serie.Values.Count });

                var yValues = scatterChartSeries.AppendChild(new YValues()).AppendChild(new NumberReference());
                yValues.AppendChild(new NumberingCache());
                yValues.NumberingCache.AppendChild(new FormatCode("General"));
                yValues.NumberingCache.AppendChild(new PointCount() { Val = (uint)serie.Values.Count });

                uint p = 0;
                foreach (var value in serie.Values)
                {
                    xValues.NumberingCache.AppendChild(new NumericPoint() { Index = p, NumericValue = new NumericValue(value.X.HasValue ? value.X.Value.ToString(CultureInfo.InvariantCulture) : string.Empty) });
                    yValues.NumberingCache.AppendChild(new NumericPoint() { Index = p, NumericValue = new NumericValue(value.Y.HasValue ? value.Y.Value.ToString(CultureInfo.InvariantCulture) : string.Empty) });
                    p++;
                }

                // Smooth
                scatterChartSeries.AppendChild(new Smooth { Val = new BooleanValue(serie.SmoothCurve) });
                i++;
            }

            ManageDataLabels(chartModel, chart);

            if (chartModel.SpaceBetweenLineCategories.HasValue)
                chart.AppendChild(new GapWidth() { Val = (ushort)chartModel.SpaceBetweenLineCategories.Value });
            else
                chart.AppendChild(new GapWidth() { Val = 55 });

            chart.AppendChild(new AxisId() { Val = categoryAxisId });
            chart.AppendChild(new AxisId() { Val = valuesAxisId });

            // Add the X Axis. 
            var xAxixModel = secondaryAxis ? chartModel.SecondaryCategoriesAxisModel : chartModel.CategoriesAxisModel;
            var xAxis = new ValueAxis(
                new AxisId() { Val = categoryAxisId },
                xAxixModel.ScalingModel.GetScaling(),
                new Delete() { Val = xAxixModel.DeleteAxis },
                new AxisPosition() { Val = secondaryAxis ? new EnumValue<AxisPositionValues>(AxisPositionValues.Top) : new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                new MajorTickMark() { Val = TickMarkValues.None },
                new MinorTickMark() { Val = TickMarkValues.None },
                new TickLabelPosition() { Val = xAxixModel.TickLabelPosition.HasValue ? new DC.TickLabelPositionValues(xAxixModel.TickLabelPosition.ToString().ToLower()) : new EnumValue<DC.TickLabelPositionValues>(DC.TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = valuesAxisId },
                new AutoLabeled() { Val = new BooleanValue(true) },
                new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                new LabelOffset() { Val = new UInt16Value((ushort)100) },
                new NoMultiLevelLabels() { Val = false },
                ManageShapeProperties(xAxixModel.ShowAxisCurve, xAxixModel.AxisCurveColor));
            if (!secondaryAxis)
                xAxis.MajorGridlines = new MajorGridlines
                {
                    ChartShapeProperties = ManageShapeProperties(xAxixModel.ShowMajorGridlines, xAxixModel.MajorGridlinesColor)
                };
            if (!string.IsNullOrWhiteSpace(xAxixModel.Title))
                xAxis.Title = ManageTitle(xAxixModel.Title, xAxixModel.LabelFormat, xAxixModel.TitleColor);
            if (xAxixModel.CrossesAt != null)
                xAxis.AppendChild(new CrossesAt() { Val = new DoubleValue(xAxixModel.CrossesAt) });
            else
                xAxis.AppendChild(new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) });
            plotArea.AppendChild(xAxis);

            // Add the Y Axis.
            var yAxixModel = secondaryAxis ? chartModel.SecondaryValuesAxisModel : chartModel.ValuesAxisModel;
            var yAxis = new ValueAxis(
                new AxisId() { Val = valuesAxisId },
                yAxixModel.ScalingModel.GetScaling(),
                new Delete() { Val = yAxixModel.DeleteAxis },
                new AxisPosition() { Val = secondaryAxis ? new EnumValue<AxisPositionValues>(AxisPositionValues.Right) : new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                new DC.NumberingFormat()
                {
                    FormatCode = new StringValue("General"),
                    SourceLinked = new BooleanValue(true)
                },
                new MajorTickMark() { Val = TickMarkValues.None },
                new MinorTickMark() { Val = TickMarkValues.None },
                new TickLabelPosition() { Val = yAxixModel.TickLabelPosition.HasValue ? yAxixModel.TickLabelPosition.Value.ToOOxml() : DC.TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = categoryAxisId },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) },
                ManageShapeProperties(yAxixModel.ShowAxisCurve, yAxixModel.AxisCurveColor));
            if (!secondaryAxis)
                yAxis.MajorGridlines = new MajorGridlines
                {
                    ChartShapeProperties = ManageShapeProperties(yAxixModel.ShowMajorGridlines, yAxixModel.MajorGridlinesColor)
                };
            if (!string.IsNullOrWhiteSpace(yAxixModel.Title))
                yAxis.Title = ManageTitle(yAxixModel.Title, yAxixModel.LabelFormat, yAxixModel.TitleColor);
            if (yAxixModel.CrossesAt != null)
                yAxis.AppendChild(new CrossesAt() { Val = new DoubleValue(yAxixModel.CrossesAt) });
            else
                yAxis.AppendChild(new Crosses() { Val = secondaryAxis ? CrossesValues.Maximum : CrossesValues.AutoZero });
            plotArea.AppendChild(yAxis);
        }

        /// <summary>
        /// Create graph
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="documentPart"></param>
        /// <returns></returns>
        private static Run CreateGraph(ScatterModel chartModel, OpenXmlPart documentPart)
        {
            if (chartModel.Series == null)
                throw new ArgumentNullException("series of chartModel must be not null");

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
            ManageScatterChart(chartModel, plotArea, new UInt32Value(48650112U), new UInt32Value(48672768U), ref i);
            if (chartModel.Series.Any(s => s.UseSecondaryAxis))
                ManageScatterChart(chartModel, plotArea, new UInt32Value(48650113U), new UInt32Value(48672708U), ref i, true);

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
        /// Manage DataLabels
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="chart"></param>
        private static void ManageDataLabels(ScatterModel chartModel, ScatterChart chart)
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
            dataLabelColor.CheckColorFormat();

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
            chart.AppendChild(dLbls);
        }

        /// <summary>
        ///  Manage axes titles
        /// </summary>
        /// <param name="title"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        private static Title ManageTitle(string title, string format, string color)
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
                color.CheckColorFormat();

                rpr.AppendChild(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } });
            }

            var paragraph = new A.Paragraph();
            paragraph.AppendChild(new A.Run
            {
                Text = new A.Text(string.IsNullOrWhiteSpace(format) ? title : string.Format(format, title)),
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
                color.CheckColorFormat();

                return new ChartShapeProperties(new A.Outline(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } }));
            }

            return new ChartShapeProperties();
        }

        /// <summary>
        /// Manage legend
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="chart"></param>
        private static void ManageLegend(ScatterModel chartModel, Chart chart)
        {
            // Add the chart Legend.
            if (chartModel.ShowLegend)
            {
                var defaultRunProperties = new A.DefaultRunProperties { Baseline = 0 };
                if (!string.IsNullOrEmpty(chartModel.FontFamilyLegend))
                    defaultRunProperties.AppendChild(new A.LatinFont { CharacterSet = 0, Typeface = chartModel.FontFamilyLegend });

                var textProperty = new TextProperties
                    (
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.ParagraphProperties(defaultRunProperties)));

                chart.AppendChild(
                    new Legend(new LegendPosition() { Val = chartModel.LegendPosition.ToOOxml() },
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
        private static void ManageGraphBorders(ScatterModel chartModel, ChartPart chartPart)
        {
            // Graph borders.
            if (chartModel.HasBorder)
            {
                chartModel.BorderWidth = chartModel.BorderWidth ?? 12700;

                if (!string.IsNullOrEmpty(chartModel.BorderColor))
                {
                    var color = chartModel.BorderColor.Replace("#", "");
                    color.CheckColorFormat();

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
        private static Run SaveChart(ScatterModel chartModel, OpenXmlPart documentPart, ChartPart chartPart)
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
