namespace McCoin.Interfaces
{
    /// <summary>
    /// Represents border details
    /// </summary>
    public interface IBorderInfo
    {
        string Top { get; set; }
        string Right { get; set; }
        string Bottom { get; set; }
        string Left { get; set; }
        string Width { get; set; }
    }
}
