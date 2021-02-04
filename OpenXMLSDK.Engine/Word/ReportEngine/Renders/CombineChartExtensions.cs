using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.Charts;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using D = OpenXMLSDK.Engine.ReportEngine.DataContext.Charts;
using DC = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    /// <summary>
    /// Combine chart render (line + bar, and more)
    /// </summary>
    public static class CombineChartExtensions
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
        public static Run Render(this CombineChartModel combineChartModel, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(combineChartModel, formatProvider);

            Run runItem = null;

            // We construct categories and series from the context object
            if (!string.IsNullOrWhiteSpace(combineChartModel.DataSourceKey) && context.TryGetItem(combineChartModel.DataSourceKey, out MultipleSeriesChartModel contextModel))
            {
                if (contextModel.ChartContent is null || contextModel.ChartContent.Categories is null
                   || contextModel.ChartContent.Series is null)
                    return runItem;

                if (contextModel.ChartContent.CategoryType != null)
                    combineChartModel.CategoryType = (D.CategoryType)contextModel.ChartContent.CategoryType;

                // Update categories object :
                combineChartModel.Categories = contextModel.ChartContent.Categories.Select(e => new ChartCategory()
                {
                    Name = e.Name,
                    Value = e.Value,
                    Color = e.Color
                }).ToList();

                // We update line series
                combineChartModel.LineSeries = contextModel.ChartContent.Series.Where(s => s.SerieChartType.Equals(D.SerieChartType.Line)).Select(e => new LineSerie()
                {
                    Name = e.Name,
                    Values = e.Values,
                    Color = e.Color,
                    DataLabelColor = e.DataLabelColor,
                    LabelFormatString = e.LabelFormatString,
                    HasBorder = e.HasBorder,
                    UseSecondaryAxis = e.UseSecondaryAxis,
                    SmoothCurve = e.SmoothCurve,
                    PresetLineDashValues = e.PresetLineDashValues
                }).ToList();

                // We update bar series
                combineChartModel.BarSeries = contextModel.ChartContent.Series.Where(s => s.SerieChartType.Equals(D.SerieChartType.Bar)).Select(e => new BarSerie()
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

                // Update Axes
                UpdateAxisFromcontext(combineChartModel.CategoriesAxisModel, contextModel.ChartContent.CategoriesAxisModel);
                UpdateAxisFromcontext(combineChartModel.ValuesAxisModel, contextModel.ChartContent.ValuesAxisModel);
                UpdateAxisFromcontext(combineChartModel.SecondaryValuesAxisModel, contextModel.ChartContent.SecondaryValuesAxisModel);
            }

            runItem = CreateGraph(combineChartModel, documentPart);

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
        private static void UpdateAxisFromcontext(ChartAxisModel template, D.AxisModel context)
        {
            if (!string.IsNullOrWhiteSpace(context.Title))
                template.Title = context.Title;

            if (!string.IsNullOrWhiteSpace(context.Color))
                template.TitleColor = context.Color;

            if (context.CrossesAt != null)
                template.CrossesAt = context.CrossesAt;

            if (!string.IsNullOrWhiteSpace(context.LabelFormat))
                template.LabelFormat = context.LabelFormat;
        }

        /// <summary>
        /// Create the graph
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="documentPart"></param>
        /// <returns></returns>
        private static Run CreateGraph(CombineChartModel chartModel, OpenXmlPart documentPart)
        {
            if (chartModel.Categories == null)
                throw new ArgumentNullException("Categories of chartModel must not be null");
            if (chartModel.LineSeries == null && chartModel.BarSeries == null)
                throw new ArgumentNullException("At least one LineSeries or BarSeries must been set");

            if (chartModel.LineSeries.Any(e => e.Values.Count != chartModel.Categories.Count))
                throw new ChartModelException("Error in LineSeries. Serie values must have same count as categories.", "004-001");
            if (chartModel.BarSeries.Any(e => e.Values.Count != chartModel.Categories.Count))
                throw new ChartModelException("Error in BarSeries. Serie values must have same count as categories.", "004-001");

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
            var addAxes = true;
            // Lines
            if (chartModel.LineSeries.Any())
            {
                LineChartExtensions.ManageLineChart(GenerateLineModelFromCombineModel(chartModel), plotArea, new UInt32Value(48650112U), new UInt32Value(48672768U), ref i);
                if (chartModel.LineSeries.Any(s => s.UseSecondaryAxis))
                    LineChartExtensions.ManageLineChart(GenerateLineModelFromCombineModel(chartModel), plotArea, new UInt32Value(48650112U), new UInt32Value(48672708U), ref i, true);
                addAxes = false;
            }

            // Bars
            if (chartModel.BarSeries.Any())
                BarChartExtensions.ManageBarChart(GenerateBarModelFromCombineModel(chartModel), plotArea, new UInt32Value(48650112U), new UInt32Value(48672768U), ref i, addAxes);

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
        /// Generate LineModel
        /// </summary>
        /// <param name="chartModel"></param>
        /// <returns></returns>
        private static LineModel GenerateLineModelFromCombineModel(CombineChartModel chartModel)
        {
            return new LineModel
            {
                GroupingValues = chartModel.GroupingValues,
                VaryColors = chartModel.VaryColors,
                Series = chartModel.LineSeries,
                CategoryType = chartModel.CategoryType,
                Categories = chartModel.Categories.Select(c => new LineCategory
                {
                    Name = c.Name,
                    Value = c.Value,
                    Color = c.Color
                }).ToList(),
                DataLabel = chartModel.DataLabel,
                DataLabelColor = chartModel.DataLabelColor,
                SpaceBetweenLineCategories = chartModel.SpaceBetweenLineCategories,
                CategoriesAxisModel = chartModel.CategoriesAxisModel,
                ValuesAxisModel = chartModel.ValuesAxisModel,
                SecondaryValuesAxisModel = chartModel.SecondaryValuesAxisModel,
                ValuesAxisScaling = chartModel.ValuesAxisScaling
            };
        }

        /// <summary>
        /// Generate BarModel
        /// </summary>
        /// <param name="chartModel"></param>
        /// <returns></returns>
        private static BarModel GenerateBarModelFromCombineModel(CombineChartModel chartModel)
        {
            return new BarModel
            {
                BarDirectionValues = Models.Charts.BarDirectionValues.Column,
                BarGroupingValues = Models.Charts.BarGroupingValues.Standard,
                Series = chartModel.BarSeries,
                CategoryType = chartModel.CategoryType,
                Categories = chartModel.Categories.Select(c => new BarCategory
                {
                    Name = c.Name,
                    Value = c.Value,
                    Color = c.Color
                }).ToList(),
                DataLabel = chartModel.DataLabel,
                DataLabelColor = chartModel.DataLabelColor,
                SpaceBetweenLineCategories = chartModel.SpaceBetweenLineCategories,
                CategoriesAxisModel = chartModel.CategoriesAxisModel,
                ValuesAxisModel = chartModel.ValuesAxisModel,
                SecondaryValuesAxisModel = chartModel.SecondaryValuesAxisModel,
                ValuesAxisScaling = new BarChartScalingModel(),
                Overlap = chartModel.Overlap
            };
        }

        /// <summary>
        /// Manage legend
        /// </summary>
        /// <param name="chartModel"></param>
        /// <param name="chart"></param>
        private static void ManageLegend(CombineChartModel chartModel, Chart chart)
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
                    new Legend(new LegendPosition() { Val = new EnumValue<DC.LegendPositionValues>((DC.LegendPositionValues)(int)chartModel.LegendPosition) },
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
        private static void ManageGraphBorders(CombineChartModel chartModel, ChartPart chartPart)
        {
            // Graph borders.
            if (chartModel.HasBorder)
            {
                chartModel.BorderWidth ??= 12700;

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
        private static Run SaveChart(CombineChartModel chartModel, OpenXmlPart documentPart, ChartPart chartPart)
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
