namespace OfficeFlow.Formats.Interfaces
{
    public interface IOfficeFormat
    {
        string[] Extensions { get; }
        
        byte[][] Hexdumps { get; }
    }
}