namespace MvvX.Plugins.OpenXMLSDK.Word
{
    public interface IBorderType
    {
        /// <summary>
        /// Type of borders.
        /// </summary>
        BorderValues? BorderValue { get; set; }

        /// <summary>
        /// Size of the order.
        /// </summary>
        int? Size { get; set; }

        /// <summary>
        /// Color of the border
        /// </summary>
        string Color { get; set; }
    }
}
