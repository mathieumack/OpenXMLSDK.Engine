using System;
using System.Collections.Generic;
using System.Text;
using ReportEngine.Core.Template;
using it = iTextSharp.text;

namespace Pdf.Engine.ReportEngine.Extensions
{
    public static class PageSizeExtensions
    {
        public static it.Rectangle ToRectangle(this PageSize PageSize)
        {
            switch (PageSize)
            {
                case PageSize.A0: return it.PageSize.A0;
                case PageSize.A1: return it.PageSize.A1;
                case PageSize.A2: return it.PageSize.A2;
                case PageSize.A3: return it.PageSize.A3;
                case PageSize.A4: return it.PageSize.A4;
                case PageSize.A5: return it.PageSize.A5;
                case PageSize.B0: return it.PageSize.B0;
                case PageSize.B1: return it.PageSize.B1;
                case PageSize.B2: return it.PageSize.B2;
                case PageSize.B3: return it.PageSize.B3;
                case PageSize.B4: return it.PageSize.B4;
                case PageSize.B5: return it.PageSize.B5;
                //case PageSize.A6: return it.PageSize.A6;
                //case PageSize.A7: return it.PageSize.A7;
                //case PageSize.A8: return it.PageSize.A8;
                //case PageSize.A9: return it.PageSize.A9;
                //case PageSize.A10: return it.PageSize.A10;
                //case PageSize.ARCH_A: return it.PageSize.ARCH_A;
                //case PageSize.ARCH_B: return it.PageSize.ARCH_B;
                //case PageSize.ARCH_C: return it.PageSize.ARCH_C;
                //case PageSize.ARCH_D: return it.PageSize.ARCH_D;
                //case PageSize.ARCH_E: return it.PageSize.ARCH_E;
                //case PageSize.B6: return it.PageSize.B6;
                //case PageSize.B7: return it.PageSize.B7;
                //case PageSize.B8: return it.PageSize.B8;
                //case PageSize.B9: return it.PageSize.B9;
                //case PageSize.CROWN_OCTAVO: return it.PageSize.CROWN_OCTAVO;
                //case PageSize.CROWN_QUARTO: return it.PageSize.CROWN_QUARTO;
                //case PageSize.DEMY_OCTAVO: return it.PageSize.DEMY_OCTAVO;
                //case PageSize.DEMY_QUARTO: return it.PageSize.DEMY_QUARTO;
                //case PageSize.EXECUTIVE: return it.PageSize.EXECUTIVE;
                //case PageSize.FLSA: return it.PageSize.FLSA;
                //case PageSize.FLSE: return it.PageSize.FLSE;
                //case PageSize.HALFLETTER: return it.PageSize.HALFLETTER;
                //case PageSize.ID_1: return it.PageSize.ID_1;
                //case PageSize.ID_2: return it.PageSize.ID_2;
                //case PageSize.ID_3: return it.PageSize.ID_3;
                //case PageSize.LARGE_CROWN_OCTAVO: return it.PageSize.LARGE_CROWN_OCTAVO;
                //case PageSize.LARGE_CROWN_QUARTO: return it.PageSize.LARGE_CROWN_QUARTO;
                //case PageSize.LEDGER: return it.PageSize.LEDGER;
                //case PageSize.LEGAL: return it.PageSize.LEGAL;
                //case PageSize.LETTER: return it.PageSize.LETTER;
                //case PageSize.NOTE: return it.PageSize.NOTE;
                //case PageSize.PENGUIN_LARGE_PAPERBACK: return it.PageSize.PENGUIN_LARGE_PAPERBACK;
                //case PageSize.PENGUIN_SMALL_PAPERBACK: return it.PageSize.PENGUIN_SMALL_PAPERBACK;
                //case PageSize.POSTCARD: return it.PageSize.POSTCARD;
                //case PageSize.ROYAL_OCTAVO: return it.PageSize.ROYAL_OCTAVO;
                //case PageSize.ROYAL_QUARTO: return it.PageSize.ROYAL_QUARTO;
                //case PageSize.SMALL_PAPERBACK: return it.PageSize.SMALL_PAPERBACK;
                //case PageSize.TABLOID: return it.PageSize.TABLOID;
                //case PageSize._11X17: return it.PageSize._11X17;
                //case PageSize.B10: return it.PageSize.B10;
                default: return it.PageSize.A4;
            }
        }
    }
}
