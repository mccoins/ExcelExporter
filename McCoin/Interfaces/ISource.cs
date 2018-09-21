namespace McCoin.Interfaces
{
    /// <summary>
    /// Generic interace usedd in order for the processor to identify export data
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface ISource<T>
    {        
        T ExportData { get; }                
    }
}
