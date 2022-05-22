namespace VbaServices
{
    public enum eSyskind
    {
        Win16 = 0,
        Win32 = 1,
        Macintosh = 2,
        Win64 = 3
    }
    public interface IVbaService
    {
        void GetVbaProject(string file);
    }
}
