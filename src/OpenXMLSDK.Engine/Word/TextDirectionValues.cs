namespace OpenXMLSDK.Engine.Word
{
    /// <summary>
    /// Values for Text directions :
    /// 
    ///    "lrTb" => true,
    ///    "tb" => true,
    ///    "tbRl" => true,
    ///    "rl" => true,
    ///    "btLr" => true,
    ///    "lr" => true,
    ///    "lrTbV" => true,
    ///    "tbV" => true,
    ///    "tbRlV" => true,
    ///    "rlV" => true,
    ///    "tbLrV" => true,
    ///    "lrV" => true,
    /// </summary>
    public enum TextDirectionValues
    {
        LefToRightTopToBottom = 0,
        LeftToRightTopToBottom2010 = 1,
        TopToBottomRightToLeft = 2,
        TopToBottomRightToLeft2010 = 3,
        BottomToTopLeftToRight = 4,
        BottomToTopLeftToRight2010 = 5,
        LefttoRightTopToBottomRotated = 6,
        LeftToRightTopToBottomRotated2010 = 7,
        TopToBottomRightToLeftRotated = 8,
        TopToBottomRightToLeftRotated2010 = 9,
        TopToBottomLeftToRightRotated = 10,
        TopToBottomLeftToRightRotated2010 = 11
    }
}
