//using AutoMapper;
//using DocumentFormat.OpenXml.Wordprocessing;

//namespace OpenXMLSDK.Engine.Platform.Word
//{
//    public static class AutoMapperInitializer
//    {
//        /// <summary>
//        /// Init auto mapper to link Models with interfaces
//        /// </summary>
//        public static void Init()
//        {
//            Mapper.Initialize(cfg =>
//            {
//                // Borders :
//                cfg.CreateMap<TableCellWidthModel, TableCellWidth>();
//                cfg.CreateMap<TableWidthModel, TableWidth>();
//                cfg.CreateMap<TableBorderModel, BorderType>();
//                cfg.CreateMap<TableBordersModel, TableBorders>()
//                .AfterMap((source, dest) =>
//                {
//                    AutoMapper.Mapper.Map(source.BottomBorder, dest.BottomBorder);
//                    AutoMapper.Mapper.Map(source.TopBorder, dest.TopBorder);
//                    AutoMapper.Mapper.Map(source.LeftBorder, dest.LeftBorder);
//                    AutoMapper.Mapper.Map(source.RightBorder, dest.RightBorder);

//                    AutoMapper.Mapper.Map(source.InsideHorizontalBorder, dest.InsideHorizontalBorder);
//                    AutoMapper.Mapper.Map(source.InsideVerticalBorder, dest.InsideVerticalBorder);
//                });

//                cfg.CreateMap<TableCellBordersModel, TableCellBorders>()
//                .AfterMap((source, dest) =>
//                {
//                    AutoMapper.Mapper.Map(source.BottomBorder, dest.BottomBorder);
//                    AutoMapper.Mapper.Map(source.TopBorder, dest.TopBorder);
//                    AutoMapper.Mapper.Map(source.LeftBorder, dest.LeftBorder);
//                    AutoMapper.Mapper.Map(source.RightBorder, dest.RightBorder);

//                    AutoMapper.Mapper.Map(source.TopLeftToBottomRightCellBorder, dest.TopLeftToBottomRightCellBorder);
//                    AutoMapper.Mapper.Map(source.TopRightToBottomLeftCellBorder, dest.TopRightToBottomLeftCellBorder);
//                });


//                cfg.CreateMap<TableRowHeightModel, TableRowHeight>();
//                cfg.CreateMap<TableRowPropertiesModel, TableRowProperties>()
//                .AfterMap((source, dest) =>
//                {
//                    AutoMapper.Mapper.Map(source.TableRowHeight, dest.TableRowHeight);
//                });

//                cfg.CreateMap<TableStyleModel, TableStyle>();
//                cfg.CreateMap<GridSpanModel, GridSpan>();
//                cfg.CreateMap<ShadingModel, Shading>();
//                cfg.CreateMap<TableWidthModel, TableWidth>();
//                cfg.CreateMap<TableCellMarginModel, TableCellMargin>()
//                .AfterMap((source, dest) =>
//                {
//                    AutoMapper.Mapper.Map(source.BottomMargin, dest.BottomMargin);
//                    AutoMapper.Mapper.Map(source.TopMargin, dest.TopMargin);
//                    AutoMapper.Mapper.Map(source.LeftMargin, dest.LeftMargin);
//                    AutoMapper.Mapper.Map(source.RightMargin, dest.RightMargin);
//                });

//                cfg.CreateMap<TableCellPropertiesModel, TableCellProperties>()
//                .AfterMap((source, dest) =>
//                {
//                    AutoMapper.Mapper.Map(source.TableCellBorders, dest.TableCellBorders);
//                    AutoMapper.Mapper.Map(source.Gridspan, dest.GridSpan);
//                    AutoMapper.Mapper.Map(source.TableCellWidth, dest.TableCellWidth);
//                    AutoMapper.Mapper.Map(source.Shading, dest.Shading);
//                    AutoMapper.Mapper.Map(source.TableCellMargin, dest.TableCellMargin);
//                });

//                cfg.CreateMap<TablePropertiesModel, TableProperties>()
//                .AfterMap((source, dest) =>
//                {
//                    AutoMapper.Mapper.Map(source.TableBorders, dest.TableBorders);
//                    AutoMapper.Mapper.Map(source.TableWidth, dest.TableWidth);
//                    AutoMapper.Mapper.Map(source.Shading, dest.Shading);
//                    AutoMapper.Mapper.Map(source.TableStyle, dest.TableStyle);
//                    AutoMapper.Mapper.Map(source.Layout, dest.Layout);
//                });

//                cfg.CreateMap<RunFontsModel, RunFonts>();
//                cfg.CreateMap<RunPropertiesModel, RunProperties>()
//                .AfterMap((source, dest) =>
//                {
//                    AutoMapper.Mapper.Map(source.RunFonts, dest.RunFonts);
//                });

//                cfg.CreateMap<NumberingPropertiesModel, NumberingProperties>();
//                cfg.CreateMap<ParagraphStyleIdModel, ParagraphStyleId>();
//                cfg.CreateMap<SpacingBetweenLinesModel, SpacingBetweenLines>();
//                cfg.CreateMap<ParagraphBorderModel, BorderType>();
//                cfg.CreateMap<ParagraphBordersModel, ParagraphBorders>()
//                .AfterMap((source, dest) =>
//                {
//                    AutoMapper.Mapper.Map(source.BottomBorder, dest.BottomBorder);
//                    AutoMapper.Mapper.Map(source.TopBorder, dest.TopBorder);
//                    AutoMapper.Mapper.Map(source.LeftBorder, dest.LeftBorder);
//                    AutoMapper.Mapper.Map(source.RightBorder, dest.RightBorder);
//                });
//                cfg.CreateMap<ParagraphPropertiesModel, ParagraphProperties>()
//                    .AfterMap((source, dest) =>
//                    {
//                        AutoMapper.Mapper.Map(source.NumberingProperties, dest.NumberingProperties);
//                        AutoMapper.Mapper.Map(source.ParagraphStyleId, dest.ParagraphStyleId);
//                        AutoMapper.Mapper.Map(source.SpacingBetweenLines, dest.SpacingBetweenLines);
//                        AutoMapper.Mapper.Map(source.ParagraphBorders, dest.ParagraphBorders);
//                    });

//                cfg.CreateMap<TableLayoutModel, TableLayout>();
//            });
//        }
//    }
//}
