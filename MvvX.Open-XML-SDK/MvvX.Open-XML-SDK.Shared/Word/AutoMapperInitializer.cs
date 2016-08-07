using AutoMapper;
using MvvX.Open_XML_SDK.Core.Word;
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
                cfg.CreateMap<TableBorderModel, IBorderType>();
                cfg.CreateMap<TableBordersModel, ITableBorders>();
                cfg.CreateMap<TablePropertiesModel, ITableProperties>()
                .AfterMap((source, dest) =>
                {
                    AutoMapper.Mapper.Map(source.TableBorders, dest.TableBorders);
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
