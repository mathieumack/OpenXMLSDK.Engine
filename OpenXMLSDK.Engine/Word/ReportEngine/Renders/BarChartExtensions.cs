using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using dc = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OpenXMLSDK.Engine.Word.Charts;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template.Charts;
using ReportEngine.Core.Template.Extensions;
using RDC = ReportEngine.Core.DataContext.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class BarChartExtensions
    {
        /// <summary>
        /// Render a table element
        /// </summary>
        /// <param name="table"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static Run Render(this BarModel barChart, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(barChart, formatProvider);

            Run runItem = null;

            if (!string.IsNullOrWhiteSpace(barChart.DataSourceKey))
            {
                // We construct categories and series from the context object
                if (context.ExistItem<BarChartModel>(barChart.DataSourceKey))
                {
                    var contextModel = context.GetItem<BarChartModel>(barChart.DataSourceKey);
                    if (contextModel.BarChartContent != null && contextModel.BarChartContent.Categories != null
                       && contextModel.BarChartContent.Series != null)
                    {
                        // Update barChart object :
                        barChart.Categories = contextModel.BarChartContent.Categories.Select(e => new BarCategory()
                        {
                            Name = e.Name,
                            Color = e.Color
                        }).ToList();

                        // We update
                        barChart.Series = contextModel.BarChartContent.Series.Select(e => new BarSerie()
                        {
                            LabelFormatString = e.LabelFormatString,
                            Color = e.Color,
                            DataLabelColor = e.DataLabelColor,
                            Values = e.Values,
                            Name = e.Name
                        }).ToList();
                    }
                    else
                        return runItem;
                }
                else if (context.ExistItem<RDC.MultipleSeriesChartModel>(barChart.DataSourceKey)) //MultipleSeriesChartModel
                {
                    var multipleSeriesContextModel = context.GetItem<RDC.MultipleSeriesChartModel>(barChart.DataSourceKey);

                    if (multipleSeriesContextModel.ChartContent != null && multipleSeriesContextModel.ChartContent.Categories != null
                     && multipleSeriesContextModel.ChartContent.Series != null)
                    {
                        // Update barChart object :
                        barChart.Categories = multipleSeriesContextModel.ChartContent.Categories.Select(e => new BarCategory()
                        {
                            Name = e.Name,
                            Color = e.Color
                        }).ToList();

                        // We update
                        barChart.Series = multipleSeriesContextModel.ChartContent.Series.Select(e => new BarSerie()
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
            }

            switch (barChart.BarChartType)
            {
                case BarChartType.BarChart:
                    runItem = CreateBarGraph(barChart, documentPart);
                    break;
            }
           
            if(runItem != null)
                parent.AppendChild(runItem);

            return runItem;
        }

        #region Internal methods

        /// <summary>
        /// Create a bargraph inside a word document
        /// </summary>
        /// <param name="chartModel">Graph model</param>
        /// <param name="showLegend"></param>
        /// <param name="title"></param>
        /// <param name="maxWidth"></param>
        /// <param name="maxHeight"></param>
        /// <exception cref="ChartModelException"></exception>
        /// <returns></returns>
        private static Run CreateBarGraph(BarModel chartModel, OpenXmlPart documentPart)
        {
            if (chartModel == null)
                throw new ArgumentNullException(nameof(chartModel), "categories must not be null");
            if (chartModel.Categories == null)
                throw new ArgumentNullException(nameof(chartModel), "categories of chartModel must not be null");
            if (chartModel.Series == null)
                throw new ArgumentNullException(nameof(chartModel), "series of chartModel must be not null");

            int countCategories = chartModel.Categories.Count;

            // Check that number of categories equals number of items in series
            var ok = chartModel.Series.Count(e => e.Values.Count != countCategories) == 0;

            if (!ok)
                throw new ChartModelException("Error in series. Serie values must have same count as categories.", "004-001");

            // Add a new chart and set the chart language to English-US.
            ChartPart chartPart = documentPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new dc.ChartSpace();
            chartPart.ChartSpace.Append(new dc.EditingLanguage() { Val = new StringValue("en-US") });
            chartPart.ChartSpace.Append(new dc.RoundedCorners { Val = new BooleanValue(chartModel.RoundedCorner) });
            dc.Chart chart = chartPart.ChartSpace.AppendChild
                <DocumentFormat.OpenXml.Drawing.Charts.Chart>
                (new dc.Chart());

            // Ajout du titre au graphique
            if (chartModel.ShowTitle)
            {
                dc.Title titleChart = chart.AppendChild<dc.Title>(new dc.Title());
                titleChart.AppendChild(new dc.ChartText(new dc.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(chartModel.Title))))));
                titleChart.AppendChild(new dc.Overlay() { Val = false });
            }

            // Create a new clustered column chart.
            dc.PlotArea plotArea = chart.AppendChild<dc.PlotArea>(new dc.PlotArea());
            dc.Layout layout = plotArea.AppendChild<dc.Layout>(new dc.Layout());
            dc.BarChart barChart = plotArea.AppendChild<dc.BarChart>(new dc.BarChart(new dc.BarDirection() { Val = new DocumentFormat.OpenXml.EnumValue<dc.BarDirectionValues>((dc.BarDirectionValues)(int)chartModel.BarDirectionValues) },
                new dc.BarGrouping() { Val = new DocumentFormat.OpenXml.EnumValue<dc.BarGroupingValues>((dc.BarGroupingValues)(int)chartModel.BarGroupingValues) }));

            uint i = 0;
            uint p = 0;
            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.
            foreach (var serie in chartModel.Series)
            {
                // Gestion des séries
                dc.BarChartSeries barChartSeries = barChart.AppendChild<dc.BarChartSeries>
                    (new dc.BarChartSeries(new dc.Index() { Val = i },
                    new dc.Order() { Val = i }, new dc.SeriesText(new dc.StringReference(new dc.StringCache(
                    new dc.PointCount() { Val = new UInt32Value(1U) },
                    new dc.StringPoint() { Index = (uint)0, NumericValue = new dc.NumericValue() { Text = serie.Name } })))));

                // Gestion de la couleur de la série
                if (!string.IsNullOrWhiteSpace(serie.Color))
                {
                    string color = serie.Color;
                    color = color.Replace("#", "");
                    if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                        throw new Exception("Error in color of serie.");

                    barChartSeries.AppendChild<A.ShapeProperties>(new A.ShapeProperties(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } }));
                }

                // Gestion des catégories
                dc.StringReference strLit = barChartSeries.AppendChild<dc.CategoryAxisData>
                        (new dc.CategoryAxisData()).AppendChild<dc.StringReference>(new dc.StringReference());
                strLit.AppendChild(new dc.StringCache());
                strLit.StringCache.AppendChild(new dc.PointCount() { Val = (uint)countCategories });
                // Liste catégorie
                foreach (var categorie in chartModel.Categories)
                {
                    strLit.StringCache.AppendChild(new dc.StringPoint() { Index = p, NumericValue = new dc.NumericValue(categorie.Name) }); // chartModel.Categories[k].Name
                    p++;
                }
                p = 0;

                // Gestion des valeurs
                dc.NumberReference numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>
                    (new dc.Values())
                        .AppendChild<dc.NumberReference>(new dc.NumberReference());
                numLit.AppendChild(new dc.NumberingCache());
                numLit.NumberingCache.AppendChild(new dc.FormatCode("General"));
                numLit.NumberingCache.AppendChild(new dc.PointCount() { Val = (uint)serie.Values.Count });
                foreach (var value in serie.Values)
                {
                    numLit.NumberingCache.AppendChild<dc.NumericPoint>(new dc.NumericPoint() { Index = p, NumericValue = new dc.NumericValue(value != null ? value.ToString() : string.Empty) });
                    p++;
                }
                i++;
            }

            dc.DataLabels dLbls = new dc.DataLabels(
                new dc.ShowLegendKey() { Val = false },
                new dc.ShowValue() { Val = chartModel.ShowDataLabel },
                new dc.ShowCategoryName() { Val = false },
                new dc.ShowSeriesName() { Val = false },
                new dc.ShowPercent() { Val = false },
                new dc.ShowBubbleSize() { Val = false });

            // Gestion de la couleur du ShowValue
            if (chartModel.ShowDataLabel && !string.IsNullOrWhiteSpace(chartModel.DataLabelColor))
            {
                string color = chartModel.DataLabelColor;
                color = color.Replace("#", "");
                if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                    throw new Exception("Error in color of serie.");

                dc.TextProperties txtPr = new dc.TextProperties(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.ParagraphProperties(
                    new A.DefaultRunProperties(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } }) { Baseline = 0 })));

                dLbls.Append(txtPr);
            }

            barChart.Append(dLbls);

            if (chartModel.SpaceBetweenLineCategories.HasValue)
                barChart.Append(new dc.GapWidth() { Val = (UInt16)chartModel.SpaceBetweenLineCategories.Value });
            else
                barChart.Append(new dc.GapWidth() { Val = 55 });

            barChart.Append(new dc.Overlap() { Val = 100 });

            barChart.Append(new dc.AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new dc.AxisId() { Val = new UInt32Value(48672768u) });
            
            // Set ShapeProperties
            dc.ShapeProperties dcSP = null;
            if (chartModel.ShowMajorGridlines)
            {
                if (!string.IsNullOrWhiteSpace(chartModel.MajorGridlinesColor))
                {
                    string color = chartModel.MajorGridlinesColor;
                    color = color.Replace("#", "");
                    if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                        throw new Exception("Error in color of grid lines.");
                    dcSP = new dc.ShapeProperties(new A.Outline(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color }}));                    
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
            dc.CategoryAxis catAx = plotArea.AppendChild<dc.CategoryAxis>(new dc.CategoryAxis(new dc.AxisId() { Val = new UInt32Value(48650112u) }, new dc.Scaling(new dc.Orientation()
            {
                Val = new DocumentFormat.OpenXml.EnumValue<dc.OrientationValues>(dc.OrientationValues.MinMax)
            }),
                new dc.Delete() { Val = chartModel.DeleteAxeCategory },
                new dc.AxisPosition() { Val = new DocumentFormat.OpenXml.EnumValue<dc.AxisPositionValues>(dc.AxisPositionValues.Left) },
                new dc.MajorTickMark() { Val = dc.TickMarkValues.None },
                new dc.MinorTickMark() { Val = dc.TickMarkValues.None },
                new dc.TickLabelPosition() { Val = new DocumentFormat.OpenXml.EnumValue<dc.TickLabelPositionValues>(dc.TickLabelPositionValues.NextTo) },
                new dc.CrossingAxis() { Val = new UInt32Value(48672768U) },
                new dc.Crosses() { Val = new DocumentFormat.OpenXml.EnumValue<dc.CrossesValues>(dc.CrossesValues.AutoZero) },
                new dc.AutoLabeled() { Val = new BooleanValue(true) },
                new dc.LabelAlignment() { Val = new DocumentFormat.OpenXml.EnumValue<dc.LabelAlignmentValues>(dc.LabelAlignmentValues.Center) },
                new dc.LabelOffset() { Val = new UInt16Value((ushort)100) },
                new dc.NoMultiLevelLabels() { Val = false },
                dcSP
                ));

            // Add the Value Axis.
            dc.ValueAxis valAx = plotArea.AppendChild<dc.ValueAxis>(new dc.ValueAxis(new dc.AxisId() { Val = new UInt32Value(48672768u) },
                new dc.Scaling(new dc.Orientation()
                {
                    Val = new DocumentFormat.OpenXml.EnumValue<dc.OrientationValues>(
                        DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }),
                new dc.Delete() { Val = chartModel.DeleteAxeValue },
                new dc.AxisPosition() { Val = new DocumentFormat.OpenXml.EnumValue<dc.AxisPositionValues>(dc.AxisPositionValues.Bottom) },
                new dc.NumberingFormat()
                {
                    FormatCode = new StringValue("General"),
                    SourceLinked = new BooleanValue(true)
                },
                new dc.MajorTickMark() { Val = dc.TickMarkValues.None },
                new dc.MinorTickMark() { Val = dc.TickMarkValues.None },
                new dc.TickLabelPosition() { Val = new DocumentFormat.OpenXml.EnumValue<dc.TickLabelPositionValues>(dc.TickLabelPositionValues.NextTo) },
                new dc.CrossingAxis() { Val = new UInt32Value(48650112U) }, new dc.Crosses()
                {
                    Val = new DocumentFormat.OpenXml.EnumValue<dc.CrossesValues>(dc.CrossesValues.AutoZero)
                }, new dc.CrossBetween() { Val = new DocumentFormat.OpenXml.EnumValue<dc.CrossBetweenValues>(dc.CrossBetweenValues.Between) }));

            if (chartModel.ShowMajorGridlines)
            {
                if (!string.IsNullOrWhiteSpace(chartModel.MajorGridlinesColor))
                {
                    string color = chartModel.MajorGridlinesColor;
                    color = color.Replace("#", "");
                    if (!Regex.IsMatch(color, "^[0-9-A-F]{6}$"))
                        throw new Exception("Error in color of grid lines.");

                    valAx.AppendChild(new dc.MajorGridlines(new dc.ShapeProperties(new A.Outline(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } }))));
                }
                else
                {
                    valAx.AppendChild(new dc.MajorGridlines());
                }
            }

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

                dc.Legend legend = chart.AppendChild<dc.Legend>(new dc.Legend(new dc.LegendPosition() { Val = new DocumentFormat.OpenXml.EnumValue<dc.LegendPositionValues>(dc.LegendPositionValues.Right) },
                new dc.Overlay() { Val = false },
                new dc.Layout(),
                textProperty));
            }

            chart.Append(new dc.PlotVisibleOnly() { Val = new BooleanValue(true) },
                new dc.DisplayBlanksAs() { Val = new DocumentFormat.OpenXml.EnumValue<dc.DisplayBlanksAsValues>(dc.DisplayBlanksAsValues.Gap) },
                new dc.ShowDataLabelsOverMaximum() { Val = false });

            // Gestion des bordures du graphique
            if (chartModel.HasBorder)
            {
                chartModel.BorderWidth = chartModel.BorderWidth.HasValue ? chartModel.BorderWidth.Value : 12700;

                if (!string.IsNullOrEmpty(chartModel.BorderColor))
                {
                    chartPart.ChartSpace.Append(new dc.ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = chartModel.BorderColor })) { Width = chartModel.BorderWidth.Value }));
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

            // Id du graphique pour faire le lien dans l'élément Drawing
            string relationshipId = documentPart.GetIdOfPart(chartPart);

            // Gestion du redimensionnement du graphique
            long imageWidth = 5486400;
            long imageHeight = 3200400;

            if (chartModel.MaxWidth.HasValue)
                // Conversion de pixel en EMU (English Metric Unit normalement c'est : EMU = pixel * 914400 / 96) --> 914400 / 96 = 9525
                imageWidth = (long)chartModel.MaxWidth * 9525;
            if (chartModel.MaxHeight.HasValue)
                imageHeight = (long)chartModel.MaxHeight * 9525;
            
            // Gestion de l'élément Drawing
            var element = new Run(
                new DocumentFormat.OpenXml.Wordprocessing.Drawing(
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
