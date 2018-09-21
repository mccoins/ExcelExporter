namespace McCoin.Interfaces
{
    /// <summary>
    /// Interface for describing a column
    /// </summary>
    public interface IColumnInfo
    {
        int DisplayOrder { get; set; }
        double Width { get; set; }
        IBorderInfo Border { get; set; }       
        IDisplayElement Header { get; set; }
        IDisplayElement Cell { get; set; }
    }
}
