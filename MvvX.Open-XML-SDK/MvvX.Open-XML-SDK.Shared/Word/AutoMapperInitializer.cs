using AutoMapper;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Models;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Core.Word.Tables.Models;

namespace MvvX.Open_XML_SDK.Shared.Word
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


                cfg.CreateMap<TableRowHeightModel, ITableRowHeight > ();
                cfg.CreateMap<TableRowPropertiesModel, ITableRowProperties>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.TableRowHeight, dest.TableRowHeight);
                });

                cfg.CreateMap<TableStyleModel, ITableStyle>();
                cfg.CreateMap<GridSpanModel, IGridSpan>();
                cfg.CreateMap<ShadingModel, IShading>();

                cfg.CreateMap<TableCellPropertiesModel, ITableCellProperties>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.TableCellBorders, dest.TableCellBorders);
                    AutoMapper.Mapper.Map(source.Gridspan, dest.GridSpan);
                    AutoMapper.Mapper.Map(source.TableCellWidth, dest.TableCellWidth);
                    AutoMapper.Mapper.Map(source.Shading, dest.Shading);
                });

                cfg.CreateMap<TablePropertiesModel, ITableProperties>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.TableBorders, dest.TableBorders);
                    AutoMapper.Mapper.Map(source.TableWidth, dest.TableWidth);
                    AutoMapper.Mapper.Map(source.Shading, dest.Shading);
                    AutoMapper.Mapper.Map(source.TableStyle, dest.TableStyle);
                });

                cfg.CreateMap<RunFontsModel, IRunFonts>();
                cfg.CreateMap<RunPropertiesModel, IRunProperties>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.RunFonts, dest.RunFonts);
                });
            });
        }
    }
}
