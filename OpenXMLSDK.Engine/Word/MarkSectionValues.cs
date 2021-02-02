namespace OpenXMLSDK.Engine.Word
{
    /// <summary>
    /// SectionMarkValues enumeration values.
    /// </summary>
    public enum MarkSectionValues
    {
        /// <summary>
        /// Next Page Section Break. When the item is serialized out as xml, its value is "nextPage".
        /// </summary>
        NextPage = 0,

        /// <summary>
        /// Column Section Break. When the item is serialized out as xml, its value is "nextColumn".
        /// </summary>
        NextColumn = 1,

        /// <summary>
        /// Continuous Section Break. When the item is serialized out as xml, its value is "continuous".
        /// </summary>
        Continuous = 2,

        /// <summary>
        /// Even Page Section Break. When the item is serialized out as xml, its value is "evenPage".
        /// </summary>
        EvenPage = 3,

        /// <summary>
        /// Odd Page Section Break. When the item is serialized out as xml, its value is "oddPage".
        /// </summary>
        OddPage = 4,
    }
}
