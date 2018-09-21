namespace McCoin.Interfaces
{
    /// <summary>
    /// Interface used for display elements for excel documents
    /// </summary>
    public interface IDisplayElement
    {
        string BGColor { get; set; }
        string Font { get; set; }
        string FontColor { get; set; }
        double FontSize { get; set; }
        string HorizontalAlignment { get; set; }
        string VerticalAlignment { get; set; }
        IBorderInfo Border { get; set; }
    }
}
