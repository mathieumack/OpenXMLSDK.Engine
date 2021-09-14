﻿using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.ReportEngine.Validations;
using OpenXMLSDK.Engine.Word.Charts;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using dc = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.Extensions;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class PieChartExtensions
    {
        /// <summary>
        /// Render a table element
        /// </summary>
        /// <param name="pieChart"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static Run Render(this PieModel pieChart, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(pieChart, formatProvider);

            Run runItem = null;

            if (context.TryGetItem(pieChart.DataSourceKey, out SingleSerieChartModel contextModel))
            {
                // We construct categories and series from the context object
                if (contextModel.ChartContent != null && contextModel.ChartContent.Categories != null
                   && contextModel.ChartContent.Serie != null)
                {
                    // Update pieChart object :
                    pieChart.Categories = contextModel.ChartContent.Categories.Select(e => new PieCategory()
                    {
                        Name = e.Name,
                        Color = e.Color
                    }).ToList();

                    // We update
                    pieChart.Serie = new PieSerie()
                    {
                        LabelFormatString = contextModel.ChartContent.Serie.LabelFormatString,
                        Color = contextModel.ChartContent.Serie.Color,
                        DataLabelColor = contextModel.ChartContent.Serie.DataLabelColor,
                        Values = contextModel.ChartContent.Serie.Values,
                        Name = contextModel.ChartContent.Serie.Name,
                        HasBorder = contextModel.ChartContent.Serie.HasBorder,
                        BorderColor = contextModel.ChartContent.Serie.BorderColor,
                        BorderWidth = contextModel.ChartContent.Serie.BorderWidth,
                    };
                }
                else
                {
                    return runItem;
                }
            }

            switch (pieChart.PieChartType)
            {
                case PieChartType.PieChart:
                    runItem = CreatePieGraph(pieChart, documentPart);
                    break;
            }

            if (runItem != null)
                parent.AppendChild(runItem);

            return runItem;
        }

        #region Internal methods

        /// <summary>
        /// Create a pieGraph inside a word document
        /// </summary>
        /// <param name="chartModel">Graph model</param>
        /// <param name="documentPart"></param>
        /// <returns></returns>
        private static Run CreatePieGraph(PieModel chartModel, OpenXmlPart documentPart)
        {
            if (chartModel.Categories == null)
                throw new ArgumentNullException("categories of chartModel must not be null");
            if (chartModel.Serie == null)
                throw new ArgumentNullException("Serie of chartModel must be not null");
            if (chartModel.Serie.Values.Count != chartModel.Categories.Count)
                throw new ChartModelException("Error in series. Serie values must have same count as categories.", "004-001");

            // Add a new chart and set the chart language to English-US.
            ChartPart chartPart = documentPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new dc.ChartSpace();
            chartPart.ChartSpace.Append(new dc.EditingLanguage() { Val = new StringValue("en-US") });
            chartPart.ChartSpace.Append(new dc.RoundedCorners { Val = new BooleanValue(chartModel.RoundedCorner) });
            dc.Chart chart = chartPart.ChartSpace.AppendChild(new dc.Chart())
                                        .TryAddTitle(chartModel);

            // Create a new clustered column chart.
            dc.PlotArea plotArea = chart.AppendChild<dc.PlotArea>(new dc.PlotArea());
            //dc.View3D view3D = chart.AppendChild<dc.View3D>(
            //    new dc.View3D()
            //    {
            //        RotateX = new dc.RotateX() { Val = 40 },
            //        RotateY = new dc.RotateY() { Val = 0 }
            //    });

            plotArea.AppendChild<dc.Layout>(new dc.Layout());
            dc.PieChart pieChart = plotArea.AppendChild<dc.PieChart>(new dc.PieChart());

            uint i = 0;
            var serie = chartModel.Serie;
            // Gestion des séries
            dc.PieChartSeries pieChartSeries = pieChart.AppendChild<dc.PieChartSeries>
                (new dc.PieChartSeries(new dc.Index() { Val = i },
                new dc.Order() { Val = i }, new dc.SeriesText(new dc.StringReference(new dc.StringCache(
                new dc.PointCount() { Val = new UInt32Value(1U) },
                new dc.StringPoint() { Index = (uint)0, NumericValue = new dc.NumericValue() { Text = serie.Name } })))));

            // Gestion de la couleur de la série
            var shapeProperties = new A.ShapeProperties()
                                                        .TryAddSolidLine(serie)
                                                        .TryAddBorder(serie);

            if (shapeProperties.HasChildren)
                pieChartSeries.AppendChild(shapeProperties);

            // Gestion des catégories
            dc.StringReference strLit = pieChartSeries.AppendChild<dc.CategoryAxisData>
                    (new dc.CategoryAxisData()).AppendChild<dc.StringReference>(new dc.StringReference());
            strLit.AppendChild(new dc.StringCache());
            strLit.StringCache.AppendChild(new dc.PointCount() { Val = (uint)chartModel.Categories.Count });
            uint p = 0;
            // Liste catégorie
            foreach (var categorie in chartModel.Categories)
            {
                strLit.StringCache.AppendChild(new dc.StringPoint() { Index = p, NumericValue = new dc.NumericValue(categorie.Name) }); // chartModel.Categories[k].Name

                if (!string.IsNullOrWhiteSpace(categorie.Color))
                {
                    string color = categorie.Color;
                    color = color.Replace("#", "");
                    color.CheckColorFormat();

                    A.ShapeProperties shapeProp = new A.ShapeProperties();

                    dc.DataPoint dataPoint = new dc.DataPoint() { Index = new dc.Index() { Val = p }, Bubble3D = new dc.Bubble3D() { Val = false } };

                    shapeProp.AppendChild(new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = color } });

                    if (serie.HasBorder)
                    {
                        shapeProp.AppendChild(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = serie.BorderColor })) { Width = serie.BorderWidth.Value });
                    }

                    dataPoint.AppendChild(shapeProp);
                    pieChartSeries.AppendChild<dc.DataPoint>(dataPoint);
                }
                p++;
            }
            p = 0;

            // Gestion des valeurs
            dc.NumberReference numLit = pieChartSeries.AppendChild(new dc.Values())
                    .AppendChild(new dc.NumberReference());
            numLit.AppendChild(new dc.NumberingCache());
            numLit.NumberingCache.AppendChild(new dc.FormatCode("General"));
            numLit.NumberingCache.AppendChild(new dc.PointCount() { Val = (uint)serie.Values.Count });
            foreach (var value in serie.Values)
            {
                numLit.NumberingCache.AppendChild<dc.NumericPoint>(new dc.NumericPoint() { Index = p, NumericValue = new dc.NumericValue(value != null ? value.ToString() : string.Empty) });
                p++;
            }

            dc.DataLabels dLbls = new dc.DataLabels(
                new dc.ShowLegendKey() { Val = false },
                new dc.ShowValue() { Val = chartModel.DataLabel?.ShowDataLabel },
                new dc.ShowCategoryName() { Val = chartModel.DataLabel?.ShowCatName },
                new dc.ShowSeriesName() { Val = false },
                new dc.ShowPercent() { Val = chartModel.DataLabel?.ShowPercent },
                new dc.ShowBubbleSize() { Val = false },
                new dc.DataLabelPosition() { Val = chartModel.DataLabel?.LabelPosition },
                new dc.Separator() { Text = chartModel.DataLabel?.Separator }
                );

            // Gestion des DataLabel
            string dataLabelColor = "#000000"; //Black by default
            if (!string.IsNullOrWhiteSpace(chartModel.DataLabelColor))
                dataLabelColor = chartModel.DataLabelColor;
            dataLabelColor = dataLabelColor.Replace("#", "");
            dataLabelColor.CheckColorFormat();

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

            pieChart.Append(dLbls);

            if (chartModel.SpaceBetweenLineCategories.HasValue)
                pieChart.Append(new dc.GapWidth() { Val = (UInt16)chartModel.SpaceBetweenLineCategories.Value });
            else
                pieChart.Append(new dc.GapWidth() { Val = 55 });

            pieChart.Append(new dc.Overlap() { Val = 100 });

            pieChart.Append(new dc.AxisId() { Val = new UInt32Value(48650112u) });
            pieChart.Append(new dc.AxisId() { Val = new UInt32Value(48672768u) });

            // Add the chart Legend.
            if (chartModel.ShowLegend)
            {
                var defaultRunProperties = new A.DefaultRunProperties { Baseline = 0 };
                if (!string.IsNullOrEmpty(chartModel.FontFamilyLegend))
                    defaultRunProperties.AppendChild(new A.LatinFont { CharacterSet = 0, Typeface = chartModel.FontFamilyLegend });

                var textProperty = new dc.TextProperties
                    (
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.ParagraphProperties(defaultRunProperties)));

                chart.AppendChild<dc.Legend>(new dc.Legend(new dc.LegendPosition() { Val = new DocumentFormat.OpenXml.EnumValue<dc.LegendPositionValues>(dc.LegendPositionValues.Right) },
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
                var color = !string.IsNullOrEmpty(chartModel.BorderColor) ? chartModel.BorderColor : "000000";

                chartPart.ChartSpace.Append(new dc.ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = color })) { Width = chartModel.BorderWidth.Value }));
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
