using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using D = DocumentFormat.OpenXml.Drawing;

namespace OpenXMLSDK.Engine.Word.Extensions
{
    /// <summary>
    /// Contains all mapping transformations for all enumerations
    /// </summary>
    internal static class EnumerationTransformationExtenions
    {
        #region GroupingValues

        internal static List<string> AvailableGroupingValues { get; } = new List<string>()
        {
            "percentStacked",
            "standard",
            "stacked"
        };

        internal static GroupingValues ToOxml(this OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.GroupingValues value)
        {
            var oXmlValue = AvailableGroupingValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return GroupingValues.Standard;
            return new GroupingValues(oXmlValue);
        }

        #endregion

        #region MarkerStyleValues

        internal static List<string> AvailableMarkerStyleValues { get; } = new List<string>()
        {
            "auto",
            "circle",
            "dash",
            "diamond",
            "dot",
            "none",
            "picture",
            "plus",
            "square",
            "star",
            "triangle",
            "x"
        };

        internal static MarkerStyleValues ToOxml(this OpenXMLSDK.Engine.ReportEngine.DataContext.Charts.MarkerStyleValues value)
        {
            var oXmlValue = AvailableMarkerStyleValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return MarkerStyleValues.None;
            return new MarkerStyleValues(oXmlValue);
        }

        #endregion

        #region PageOrientationValues

        internal static W.PageOrientationValues ToOxml(this OpenXMLSDK.Engine.Word.PageOrientationValues value)
        {
            if (value == OpenXMLSDK.Engine.Word.PageOrientationValues.Landscape)
                return W.PageOrientationValues.Landscape;
            return W.PageOrientationValues.Portrait;
        }

        #endregion

        #region SpaceProcessingModeValues

        internal static O.SpaceProcessingModeValues ToOxml(this OpenXMLSDK.Engine.Word.SpaceProcessingModeValues value)
        {
            if (value == SpaceProcessingModeValues.Preserve)
                return O.SpaceProcessingModeValues.Preserve;

            return O.SpaceProcessingModeValues.Default;
        }

        #endregion

        #region HeaderFooterValues

        internal static W.HeaderFooterValues? ToOxml(this OpenXMLSDK.Engine.Word.HeaderFooterValues value)
        {
            if(value == HeaderFooterValues.Even)
                return W.HeaderFooterValues.Even;
            if(value == HeaderFooterValues.Default) 
                return W.HeaderFooterValues.Default;
            if(value == HeaderFooterValues.First)
                return W.HeaderFooterValues.First;

            return null;
        }


        #endregion

        #region TableVerticalAlignmentValues

        internal static TableVerticalAlignmentValues ToOxml(this OpenXMLSDK.Engine.Word.Tables.TableVerticalAlignmentValues value)
        {
            if (value == OpenXMLSDK.Engine.Word.Tables.TableVerticalAlignmentValues.Top)
                return TableVerticalAlignmentValues.Top;
            if (value == OpenXMLSDK.Engine.Word.Tables.TableVerticalAlignmentValues.Center)
                return TableVerticalAlignmentValues.Center;
            return TableVerticalAlignmentValues.Bottom;
        }

        #endregion

        #region TextDirectionValues?

        /// <summary>
        /// Available values :
        /// 
        /// LefToRightTopToBottom = 0,
        /// LeftToRightTopToBottom2010 = 1,
        /// TopToBottomRightToLeft = 2,
        /// TopToBottomRightToLeft2010 = 3,
        ///  BottomToTopLeftToRight = 4,
        /// BottomToTopLeftToRight2010 = 5,
        /// LefttoRightTopToBottomRotated = 6,
        /// LeftToRightTopToBottomRotated2010 = 7,
        ///  TopToBottomRightToLeftRotated = 8,
        /// TopToBottomRightToLeftRotated2010 = 9,
        /// TopToBottomLeftToRightRotated = 10,
        /// TopToBottomLeftToRightRotated2010 = 11
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal static W.TextDirectionValues ToOOxml(this OpenXMLSDK.Engine.Word.TextDirectionValues? value)
        {
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.LefToRightTopToBottom)
                return W.TextDirectionValues.LefToRightTopToBottom;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.LeftToRightTopToBottom2010)
                return W.TextDirectionValues.LeftToRightTopToBottom2010;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.TopToBottomRightToLeft)
                return W.TextDirectionValues.TopToBottomRightToLeft;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.TopToBottomRightToLeft2010)
                return W.TextDirectionValues.TopToBottomRightToLeft2010;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.BottomToTopLeftToRight)
                return W.TextDirectionValues.BottomToTopLeftToRight;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.BottomToTopLeftToRight2010)
                return W.TextDirectionValues.BottomToTopLeftToRight2010;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.LefttoRightTopToBottomRotated)
                return W.TextDirectionValues.LefttoRightTopToBottomRotated;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.LeftToRightTopToBottomRotated2010)
                return W.TextDirectionValues.LeftToRightTopToBottomRotated2010;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.TopToBottomRightToLeftRotated)
                return W.TextDirectionValues.TopToBottomRightToLeftRotated;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.TopToBottomRightToLeftRotated2010)
                return W.TextDirectionValues.TopToBottomRightToLeftRotated2010;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.TopToBottomLeftToRightRotated)
                return W.TextDirectionValues.TopToBottomLeftToRightRotated;
            if (value == OpenXMLSDK.Engine.Word.TextDirectionValues.TopToBottomLeftToRightRotated2010)
                return W.TextDirectionValues.TopToBottomLeftToRightRotated2010;
            return W.TextDirectionValues.LefToRightTopToBottom;
        }

        #endregion

        #region StyleValues

        internal static List<string> AvailableStyleValues { get; } = new List<string>()
        {
            "paragraph",
            "character",
            "table",
            "numbering"
        };

        internal static W.StyleValues ToOOxml(this OpenXMLSDK.Engine.Word.StyleValues value)
        {
            var oXmlValue = AvailableStyleValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return W.StyleValues.Paragraph;
            return new W.StyleValues();
        }

        #endregion

        #region TableWidthUnitValues

        internal static List<string> AvailableTableWidthUnitValues { get; } = new List<string>()
    {
        "nil",
        "auto",
        "pct",
        "dxa"
    };

        internal static TableWidthUnitValues ToOOxml(this OpenXMLSDK.Engine.Word.Tables.TableWidthUnitValues value)
        {
            var oXmlValue = AvailableTableWidthUnitValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return TableWidthUnitValues.Nil;
            return new TableWidthUnitValues(oXmlValue);
        }

        #endregion

        #region ShadingPatterns

        internal static List<string> AvailableShadingPatterns { get; } = new List<string>()
    {
        "nil",
        "clear",
        "solid",
        "horzStripe",
        "vertStripe",
        "reverseDiagStripe",
        "diagStripe",
        "horzCross",
        "diagCross",
        "thinHorzStripe",
        "thinVertStripe",
        "thinReverseDiagStripe",
        "thinDiagStripe",
        "thinHorzCross",
        "thinDiagCross",
        "pct5",
        "pct10",
        "pct12",
        "pct15",
        "pct20",
        "pct25",
        "pct30",
        "pct35",
        "pct37",
        "pct40",
        "pct45",
        "pct50",
        "pct55",
        "pct60",
        "pct62",
        "pct65",
        "pct70",
        "pct75",
        "pct80",
        "pct85",
        "pct87",
        "pct90",
        "pct95"
    };

        internal static W.ShadingPatternValues ToOOxml(this OpenXMLSDK.Engine.Word.ShadingPatternValues value)
        {
            var oXmlValue = AvailableShadingPatterns.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return W.ShadingPatternValues.Clear;
            return new W.ShadingPatternValues(oXmlValue);
        }

        #endregion

        #region BarDirectionValues

        internal static BarDirectionValues ToOOxml(this OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.BarDirectionValues value)
        {
            if (value == OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.BarDirectionValues.Column)
                return BarDirectionValues.Column;
            return BarDirectionValues.Bar;
        }

        #endregion

        internal static OrientationValues ToOOxml(this OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.OrientationType orientation)
        {
            if (orientation == OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.OrientationType.MaxMin)
                return OrientationValues.MaxMin;
            return OrientationValues.MinMax;
        }

        internal static List<string> AvailableBarGroupingValue { get; } = new List<string>()
    {
        "percentStacked",
        "clustered",
        "standard",
        "stacked"
    };

        internal static BarGroupingValues ToOOxml(this OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.BarGroupingValues value)
        {
            var oXmlValue = AvailableBarGroupingValue.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return BarGroupingValues.Standard;
            return new BarGroupingValues(oXmlValue);
        }

        internal static OrientationValues ToOOxml(this OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.BarChartOrientationType orientation)
        {
            if (orientation == OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.BarChartOrientationType.MaxMin)
                return OrientationValues.MaxMin;
            return OrientationValues.MinMax;
        }

        #region TabStopValues

        internal static List<string> AvailableTabStopValues { get; } = new List<string>()
    {
        "clear",
        "left",
        "start",
        "center",
        "right",
        "end",
        "decimal",
        "bar",
        "num"
    };

        internal static TabStopValues ToOOxml(this OpenXMLSDK.Engine.Word.TabAlignmentValues value)
        {
            var oXmlValue = AvailableTabStopValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return TabStopValues.Clear;
            return new TabStopValues(oXmlValue);
        }

        #endregion

        #region LegendPositionValues

        internal static LegendPositionValues ToOOxml(this OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.LegendPositionValues value)
        {
            if (value == OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.LegendPositionValues.Bottom)
                return LegendPositionValues.Bottom;
            if (value == OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.LegendPositionValues.Left)
                return LegendPositionValues.Left;
            if (value == OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.LegendPositionValues.Right)
                return LegendPositionValues.Right;
            if (value == OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.LegendPositionValues.Top)
                return LegendPositionValues.Top;
            return LegendPositionValues.TopRight;
        }

        #endregion

        #region PresetLineDashValues

        internal static List<string> AvailablePresetLineDashValues { get; } = new List<string>()
    {
        "solid",
        "dot",
        "dash",
        "lgDash",
        "dashDot",
        "lgDashDot",
        "lgDashDotDot",
        "sysDash",
        "sysDot",
        "sysDashDot",
        "sysDashDotDot"
    };

        internal static D.PresetLineDashValues ToOOxml(this OpenXMLSDK.Engine.ReportEngine.DataContext.Charts.PresetLineDashValues value)
        {
            var oXmlValue = AvailablePresetLineDashValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return D.PresetLineDashValues.Solid;
            return new D.PresetLineDashValues(oXmlValue);
        }

        #endregion

        #region TickLabelPositionValues

        internal static List<string> AvailableTickLabelPositionValues { get; } = new List<string>()
    {
        "high",
        "low",
        "nextTo",
        "none"
    };

        internal static TickLabelPositionValues ToOOxml(this OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts.TickLabelPositionValues value)
        {
            var oXmlValue = AvailableTickLabelPositionValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return TickLabelPositionValues.NextTo;
            return new TickLabelPositionValues(oXmlValue);
        }

        #endregion

        #region TabStopLeaderCharValues

        internal static List<string> AvailableTabStopLeaderCharValues { get; } = new List<string>()
    {
        "none",
        "dot",
        "hyphen",
        "underscore",
        "heavy",
        "middleDot"
    };

        internal static TabStopLeaderCharValues ToOOxml(this OpenXMLSDK.Engine.Word.ReportEngine.Models.TabStopLeaderCharValues value)
        {
            var oXmlValue = AvailableTabStopLeaderCharValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return TabStopLeaderCharValues.None;
            return new TabStopLeaderCharValues(oXmlValue);
        }

        #endregion

        #region BorderValues

        internal static List<string> AvailableBorderValues { get; } = new List<string>()
{
    "nil",
    "none",
    "single",
    "thick",
    "double",
    "dotted",
    "dashed",
    "dotDash",
    "dotDotDash",
    "triple",
    "thinThickSmallGap",
    "thickThinSmallGap",
    "thinThickThinSmallGap",
    "thinThickMediumGap",
    "thickThinMediumGap",
    "thinThickThinMediumGap",
    "thinThickLargeGap",
    "thickThinLargeGap",
    "thinThickThinLargeGap",
    "wave",
    "doubleWave",
    "dashSmallGap",
    "dashDotStroked",
    "threeDEmboss",
    "threeDEngrave",
    "outset",
    "inset",
    "apples",
    "archedScallops",
    "babyPacifier",
    "babyRattle",
    "balloons3Colors",
    "balloonsHotAir",
    "basicBlackDashes",
    "basicBlackDots",
    "basicBlackSquares",
    "basicThinLines",
    "basicWhiteDashes",
    "basicWhiteDots",
    "basicWhiteSquares",
    "basicWideInline",
    "basicWideMidline",
    "basicWideOutline",
    "bats",
    "birds",
    "birdsFlight",
    "cabins",
    "cakeSlice",
    "candyCorn",
    "celticKnotwork",
    "certificateBanner",
    "chainLink",
    "champagneBottle",
    "checkedBarBlack",
    "checkedBarColor",
    "checkered",
    "christmasTree",
    "circlesLines",
    "circlesRectangles",
    "classicalWave",
    "clocks",
    "compass",
    "confetti",
    "confettiGrays",
    "confettiOutline",
    "confettiStreamers",
    "confettiWhite",
    "cornerTriangles",
    "couponCutoutDashes",
    "couponCutoutDots",
    "crazyMaze",
    "creaturesButterfly",
    "creaturesFish",
    "creaturesInsects",
    "creaturesLadyBug",
    "crossStitch",
    "cup",
    "decoArch",
    "decoArchColor",
    "decoBlocks",
    "diamondsGray",
    "doubleD",
    "doubleDiamonds",
    "earth1",
    "earth2",
    "eclipsingSquares1",
    "eclipsingSquares2",
    "eggsBlack",
    "fans",
    "film",
    "firecrackers",
    "flowersBlockPrint",
    "flowersDaisies",
    "flowersModern1",
    "flowersModern2",
    "flowersPansy",
    "flowersRedRose",
    "flowersRoses",
    "flowersTeacup",
    "flowersTiny",
    "gems",
    "gingerbreadMan",
    "gradient",
    "handmade1",
    "handmade2",
    "heartBalloon",
    "heartGray",
    "hearts",
    "heebieJeebies",
    "holly",
    "houseFunky",
    "hypnotic",
    "iceCreamCones",
    "lightBulb",
    "lightning1",
    "lightning2",
    "mapPins",
    "mapleLeaf",
    "mapleMuffins",
    "marquee",
    "marqueeToothed",
    "moons",
    "mosaic",
    "musicNotes",
    "northwest",
    "ovals",
    "packages",
    "palmsBlack",
    "palmsColor",
    "paperClips",
    "papyrus",
    "partyFavor",
    "partyGlass",
    "pencils",
    "people",
    "peopleWaving",
    "peopleHats",
    "poinsettias",
    "postageStamp",
    "pumpkin1",
    "pushPinNote2",
    "pushPinNote1",
    "pyramids",
    "pyramidsAbove",
    "quadrants",
    "rings",
    "safari",
    "sawtooth",
    "sawtoothGray",
    "scaredCat",
    "seattle",
    "shadowedSquares",
    "sharksTeeth",
    "shorebirdTracks",
    "skyrocket",
    "snowflakeFancy",
    "snowflakes",
    "sombrero",
    "southwest",
    "stars",
    "starsTop",
    "stars3d",
    "starsBlack",
    "starsShadowed",
    "sun",
    "swirligig",
    "tornPaper",
    "tornPaperBlack",
    "trees",
    "triangleParty",
    "triangles",
    "tribal1",
    "tribal2",
    "tribal3",
    "tribal4",
    "tribal5",
    "tribal6",
    "triangle1",
    "triangle2",
    "triangleCircle1",
    "triangleCircle2",
    "shapes1",
    "shapes2",
    "twistedLines1",
    "twistedLines2",
    "vine",
    "waveline",
    "weavingAngles",
    "weavingBraid",
    "weavingRibbon",
    "weavingStrips",
    "whiteFlowers",
    "woodwork",
    "xIllusions",
    "zanyTriangles",
    "zigZag",
    "zigZagStitch"
};

        internal static W.BorderValues ToOOxml(this OpenXMLSDK.Engine.Word.BorderValues? value)
        {
            var oXmlValue = AvailableBorderValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return W.BorderValues.None;
            return new W.BorderValues(oXmlValue);
        }

        #endregion

        #region JustificationValues

        internal static List<string> AvailableJustificationValues { get; } = new List<string>()
{
    "left",
    "start",
    "center",
    "right",
    "end",
    "both",
    "mediumKashida",
    "distribute",
    "numTab",
    "highKashida",
    "lowKashida",
    "thaiDistribute"
};

        internal static W.JustificationValues ToOOxml(this OpenXMLSDK.Engine.Word.JustificationValues value)
        {
            var oXmlValue = AvailableJustificationValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return W.JustificationValues.Left;
            return new W.JustificationValues(oXmlValue);
        }

        #endregion

        #region JustificationValues

        internal static List<string> AvailableUnderlineValues { get; } = new List<string>()
    {
        "single",
        "words",
        "double",
        "thick",
        "dotted",
        "dottedHeavy",
        "dash",
        "dashedHeavy",
        "dashLong",
        "dashLongHeavy",
        "dotDash",
        "dashDotHeavy",
        "dotDotDash",
        "dashDotDotHeavy",
        "wave",
        "wavyHeavy",
        "wavyDouble",
        "none"
    };

        internal static UnderlineValues ToUnderlineValue(this OpenXMLSDK.Engine.Word.ReportEngine.Models.ExtendedModels.UnderlineValues value)
        {
            var oXmlValue = AvailableUnderlineValues.FirstOrDefault(e => value.ToString().ToLower().Equals(e.ToLower()));
            if (string.IsNullOrWhiteSpace(oXmlValue))
                return UnderlineValues.None;
            return new UnderlineValues(oXmlValue);
        }

        #endregion

    }
}