using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Entity;
using System.IO;
using System.ComponentModel;
using System.Reflection;

namespace Template4338
{
    /// <summary>
    /// Логика взаимодействия для WindowInfo.xaml
    /// </summary>
    public partial class _4338_Ilyasov : Window
    {
        public _4338_Ilyasov()
        {
            InitializeComponent();
        }

        static void Import()
        {
            // Создание нового приложения Excel
            string path = AppDomain.CurrentDomain.BaseDirectory;
            path += "3.xlsx";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // Подключение к базе данных с использованием EF
            using (var db = new MyDbContext())
            {
                int numberOfRowDeleted = db.Database.ExecuteSqlCommand("Truncate table Tables");

                for (int i = 2; i <= rowCount; i++)
                {
                    // Создание нового объекта для каждой строки в Excel
                    var myTable = new Table();
                    var descr = TypeDescriptor.GetProperties(myTable);
                    for (int j = 1; j <= colCount; j++)
                    {
                        // Новые строки Excel начинаются с 1, а не с 0
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value != null)
                        {

                            PropertyDescriptor property = descr[j-1];
                            switch (property.Name) {
                                case "Id": continue;
                                default:
                                    property.SetValue(myTable, System.Convert.ToString(xlRange.Cells[i, j].Value.ToString())); 
                                    break;
                            }
                            
                        }
                    }
                    // Добавление объекта в контекст EF
                    db.Tables.Add(myTable);
                }
                // Сохранение изменений в базе данных
                db.SaveChanges();
            }

            // Очистка
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Закрытие и освобождение
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }
        private void ImportClick(object sender, RoutedEventArgs e)
        {
            Import();
        }

        private void ExportClick(object sender, RoutedEventArgs e)
        {

        }
    }
}
