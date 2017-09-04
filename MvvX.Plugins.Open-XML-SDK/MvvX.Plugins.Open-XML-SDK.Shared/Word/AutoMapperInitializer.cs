using AutoMapper;
using MvvX.Plugins.OpenXMLSDK.Word;
using MvvX.Plugins.OpenXMLSDK.Word.Models;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs.Models;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;
using MvvX.Plugins.OpenXMLSDK.Word.Tables.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
{
    public static class AutoMapperInitializer
    {
        /// <summary>
        /// Init auto mapper to link Models with interfaces
        /// </summary>
        public static void Init()
        {
            Mapper.Initialize(cfg =>
            {
                // Borders :
                cfg.CreateMap<TableCellWidthModel, ITableCellWidth>();
                cfg.CreateMap<TableWidthModel, ITableWidth>();
                cfg.CreateMap<TableBorderModel, IBorderType>();
                cfg.CreateMap<TableBordersModel, ITableBorders>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.BottomBorder, dest.BottomBorder);
                    AutoMapper.Mapper.Map(source.TopBorder, dest.TopBorder);
                    AutoMapper.Mapper.Map(source.LeftBorder, dest.LeftBorder);
                    AutoMapper.Mapper.Map(source.RightBorder, dest.RightBorder);

                    AutoMapper.Mapper.Map(source.InsideHorizontalBorder, dest.InsideHorizontalBorder);
                    AutoMapper.Mapper.Map(source.InsideVerticalBorder, dest.InsideVerticalBorder);
                });

                cfg.CreateMap<TableCellBordersModel, ITableCellBorders>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.BottomBorder, dest.BottomBorder);
                    AutoMapper.Mapper.Map(source.TopBorder, dest.TopBorder);
                    AutoMapper.Mapper.Map(source.LeftBorder, dest.LeftBorder);
                    AutoMapper.Mapper.Map(source.RightBorder, dest.RightBorder);

                    AutoMapper.Mapper.Map(source.TopLeftToBottomRightCellBorder, dest.TopLeftToBottomRightCellBorder);
                    AutoMapper.Mapper.Map(source.TopRightToBottomLeftCellBorder, dest.TopRightToBottomLeftCellBorder);
                });


                cfg.CreateMap<TableRowHeightModel, ITableRowHeight>();
                cfg.CreateMap<TableRowPropertiesModel, ITableRowProperties>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.TableRowHeight, dest.TableRowHeight);
                });

                cfg.CreateMap<TableStyleModel, ITableStyle>();
                cfg.CreateMap<GridSpanModel, IGridSpan>();
                cfg.CreateMap<ShadingModel, IShading>();
                cfg.CreateMap<TableWidthModel, ITableWidth>();
                cfg.CreateMap<TableCellMarginModel, ITableCellMargin>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.BottomMargin, dest.BottomMargin);
                    AutoMapper.Mapper.Map(source.TopMargin, dest.TopMargin);
                    AutoMapper.Mapper.Map(source.LeftMargin, dest.LeftMargin);
                    AutoMapper.Mapper.Map(source.RightMargin, dest.RightMargin);
                });

                cfg.CreateMap<TableCellPropertiesModel, ITableCellProperties>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.TableCellBorders, dest.TableCellBorders);
                    AutoMapper.Mapper.Map(source.Gridspan, dest.GridSpan);
                    AutoMapper.Mapper.Map(source.TableCellWidth, dest.TableCellWidth);
                    AutoMapper.Mapper.Map(source.Shading, dest.Shading);
                    AutoMapper.Mapper.Map(source.TableCellMargin, dest.TableCellMargin);
                });

                cfg.CreateMap<TablePropertiesModel, ITableProperties>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.TableBorders, dest.TableBorders);
                    AutoMapper.Mapper.Map(source.TableWidth, dest.TableWidth);
                    AutoMapper.Mapper.Map(source.Shading, dest.Shading);
                    AutoMapper.Mapper.Map(source.TableStyle, dest.TableStyle);
                    AutoMapper.Mapper.Map(source.Layout, dest.Layout);
                });

                cfg.CreateMap<RunFontsModel, IRunFonts>();
                cfg.CreateMap<RunPropertiesModel, IRunProperties>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.RunFonts, dest.RunFonts);
                });

                cfg.CreateMap<NumberingPropertiesModel, INumberingProperties>();
                cfg.CreateMap<ParagraphStyleIdModel, IParagraphStyleId>();
                cfg.CreateMap<SpacingBetweenLinesModel, ISpacingBetweenLines>();
                cfg.CreateMap<ParagraphBorderModel, IBorderType>();
                cfg.CreateMap<ParagraphBordersModel, IParagraphBorders>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.BottomBorder, dest.BottomBorder);
                    AutoMapper.Mapper.Map(source.TopBorder, dest.TopBorder);
                    AutoMapper.Mapper.Map(source.LeftBorder, dest.LeftBorder);
                    AutoMapper.Mapper.Map(source.RightBorder, dest.RightBorder);
                });
                cfg.CreateMap<ParagraphPropertiesModel, IParagraphProperties>()
                    .AfterMap((source, dest) =>
                    {
                        AutoMapper.Mapper.Map(source.NumberingProperties, dest.NumberingProperties);
                        AutoMapper.Mapper.Map(source.ParagraphStyleId, dest.ParagraphStyleId);
                        AutoMapper.Mapper.Map(source.SpacingBetweenLines, dest.SpacingBetweenLines);
                        AutoMapper.Mapper.Map(source.ParagraphBorders, dest.ParagraphBorders);
                    });

                cfg.CreateMap<TableLayoutModel, ITableLayout>();
            });
        }
    }
}
