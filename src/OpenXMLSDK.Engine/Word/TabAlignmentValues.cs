namespace OpenXMLSDK.Engine.Word
{
    /// <summary>
    /// Defines the TabAlignmentValues enumeration.
    /// </summary>
    public enum TabAlignmentValues
    {
        /// <summary>
        /// No Tab Stop.
        /// <para>When the item is serialized out as xml, its value is "clear".</para>
        /// </summary>
        Clear = 0,

        /// <summary>
        /// Left Tab.
        /// <para>When the item is serialized out as xml, its value is "left".</para>
        /// </summary>
        Left = 1,

        /// <summary>
        /// Start
        /// <para>When the item is serialized out as xml, its value is "start".</para>
        /// </summary>
        Start = 2,

        /// <summary>
        /// Centered Tab.
        /// <para>When the item is serialized out as xml, its value is "center".</para>
        /// </summary>
        Center = 3,

        /// <summary>
        /// Right Tab.
        /// <para>When the item is serialized out as xml, its value is "right".</para>
        /// </summary>
        Right = 4,

        /// <summary>
        /// end.
        /// <para>When the item is serialized out as xml, its value is "end".</para>
        /// </summary>
        End = 5,

        /// <summary>
        /// Decimal Tab.
        /// <para>When the item is serialized out as xml, its value is "decimal".</para>
        /// </summary>
        Decimal = 6,

        /// <summary>
        /// Bar Tab.
        /// <para>When the item is serialized out as xml, its value is "bar".</para>
        /// </summary>
        Bar = 7,

        /// <summary>
        /// List Tab.
        /// <para>When the item is serialized out as xml, its value is "num".</para>
        /// </summary>
        Number = 8
    }
}
