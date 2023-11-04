using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;


internal class Program
{
    private static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Excel.Application xlApp = new
        Excel.Application();
        FileInfo exc_file = new FileInfo(@"D:\excelOlustur.xlsx");
        ExcelPackage pck = new ExcelPackage();
        ExcelWorkbook wb = pck.Workbook;
        ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add("1.Sayfa");

        Console.WriteLine("Excel sayfanız D klasorünün altında başarıyla oluşturuldu.");

        bool continueSelect = true;

        while (continueSelect)
        {
            Console.WriteLine("İşlem yapmak ister misiniz?..=>...Y/N");
            if (Console.ReadLine().ToLower() == "y")
            {
                int operation = Select();

                if (operation == 1)
                {
                    WriteDataToCell(worksheet);
                }
                else if (operation == 2)
                {
                    ReadDataInCell(worksheet);
                }
                else if (operation == 3)
                {
                    CreateFormula(worksheet);
                }
                else
                {
                    Console.WriteLine("Hatalı seçim yaptınız.");
                }


            }
            else if(Console.ReadLine().ToLower() == "n")
            {
                pck.SaveAs(exc_file);
                pck.Dispose();
                Environment.Exit(2);
            }
        }


    }

    private static void WriteDataToCell(ExcelWorksheet worksheet)
    {
        try
        {
            Console.WriteLine("İşlem yapmak istediğiniz hücreyi giriniz.");
            string cell = Console.ReadLine();
            Console.WriteLine($"{cell} hücresine yazmak istediğiniz değeri giriniz.");
            worksheet.Cells[$"{cell}"].Value = Console.ReadLine();
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
        
    }
    private static void ReadDataInCell(ExcelWorksheet worksheet)
    {
        try
        {
            Console.WriteLine("İşlem yapmak istediğiniz hücreyi giriniz.");
            string cell = Console.ReadLine();

            object value = worksheet.Cells[cell].Text;
            Console.WriteLine($"{cell} hücresindeki değer => {value + worksheet.Cells[cell].Formula}");
        }
        catch (Exception e)
        {

            Console.WriteLine(e.Message);
        }
        
    }

    private static void CreateFormula(ExcelWorksheet worksheet)
    {
        try
        {
            Console.WriteLine("İşlem yapmak istediğiniz hücreyi giriniz.");
            string cell = Console.ReadLine();
            Console.WriteLine($"{cell} hücresine yapmak istediğiniz işlemi giriniz.");
            string formula = Console.ReadLine();
            worksheet.Cells[cell].Formula = $"{formula}";
        }
        catch (Exception e)
        {

            Console.WriteLine(e.Message);
        }
        

    }
    private static int Select()
    {
        Console.WriteLine("Excel Sayfasında yapmak istediğiniz işlemi seçer misiniz?");
        Console.WriteLine("1-Hücreye veri yazma");
        Console.WriteLine("2-Hücredeki veriyi okuma");
        Console.WriteLine("3-Formül ekleme");

        int operation = int.Parse(Console.ReadLine());

        return operation;
    }
}