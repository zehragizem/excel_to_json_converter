using exceltojson;
using OfficeOpenXml;

static class Program
{
    [STAThread]
    static void Main()
    {
        // Lisans baðlamýný ayarla
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new Form1());
    }
}
