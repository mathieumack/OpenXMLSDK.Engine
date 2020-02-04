using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportEngine.Core.Template.ExtendedModels
{
    /// <summary>
    /// Margins for a single table cell 
    /// </summary>
    public class MarginModel
    {
        /// <summary>
        /// Distance (in dxa: twentieths of a point) between the left edge of the cell and the left edge of the content of this cell.
        /// </summary>
        public int Left { get; set; }

        /// <summary>
        /// Distance (in dxa: twentieths of a point) between the top edge of the cell and the top edge of the content of this cell.
        /// </summary>
        public int Top { get; set; }

        /// <summary>
        /// Distance (in dxa: twentieths of a point) between the right edge of the cell and the right edge of the content of this cell.
        /// </summary>
        public int Right { get; set; }

        /// <summary>
        /// Distance (in dxa: twentieths of a point) between the bottom edge of the cell and the bottom edge of the content of this cell.
        /// </summary>
        public int Bottom { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public MarginModel(){}

        /// <summary>
        /// Constructor
        /// </summary>
        public MarginModel(int value)
        {
            Left = value;
            Top = value;
            Right = value;
            Bottom = value;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public MarginModel(int h,int v)
        {
            Left = h;
            Top = v;
            Right = h;
            Bottom = v;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public MarginModel(int left, int top, int right, int bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }
    }
}
