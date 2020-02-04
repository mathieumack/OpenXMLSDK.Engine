namespace ReportEngine.Core.Template.Text
{
    /// <summary>
    /// Define a transformation on a label before rendering
    /// </summary>
    public enum LabelTransformOperationType
    {
        Nothing = 0,
        ToUpper = 1,
        ToLower = 2,
        ToUpperInvariant = 3,
        ToLowerInvariant = 4,
        Trim = 10,
        TrimStart = 11,
        TrimEnd = 12
    }
}
