using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace XLSX_to_XLS
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string klasorYolu = Path.Combine(Environment.CurrentDirectory, "listeler");
            string hedefKlasor = Path.Combine(Environment.CurrentDirectory, "bitenler");

            // Listeler klasör yoksa oluştur
            if (!Directory.Exists(klasorYolu))
            {
                Directory.CreateDirectory(klasorYolu);
            }

            // Hedef klasör yoksa oluştur
            if (!Directory.Exists(hedefKlasor))
            {
                Directory.CreateDirectory(hedefKlasor);
            }

            // Excel uygulamasını başlat
            Application excel = new Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;

            try
            {
                // Tüm .xlsx dosyalarını bul
                string[] dosyalar = Directory.GetFiles(klasorYolu, "*.xlsx");

                foreach (string dosya in dosyalar)
                {
                    Console.WriteLine($"Dönüştürülüyor: {Path.GetFileName(dosya)}");

                    // Dosyayı aç
                    Workbook workbook = excel.Workbooks.Open(dosya);

                    // Tüm sayfalardaki hücreleri metin formatına çevir
                    foreach (Worksheet sheet in workbook.Sheets)
                    {
                        Range kullanılanAlan = sheet.UsedRange;
                        kullanılanAlan.NumberFormat = "@"; // @ işareti metin formatını temsil eder
                    }

                    // Yeni dosya adını oluştur
                    string yeniDosyaAdi = Path.Combine(hedefKlasor, 
                        Path.GetFileNameWithoutExtension(dosya) + ".xls");

                    // XLS formatında kaydet
                    workbook.SaveAs(yeniDosyaAdi, XlFileFormat.xlExcel8);
                    workbook.Close();

                    Console.WriteLine($"Tamamlandı: {Path.GetFileName(yeniDosyaAdi)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hata oluştu: {ex.Message}");
            }
            finally
            {
                // Excel uygulamasını temiz bir şekilde kapat
                excel.Quit();
                Marshal.ReleaseComObject(excel);
            }

            Console.WriteLine("\nTüm işlemler tamamlandı. Çıkmak için bir tuşa basın...");
            Console.ReadKey();
        }
    }
}
