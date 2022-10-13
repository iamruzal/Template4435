using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;


namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BnTask_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Сабиров Зульфат Зуфарович","4435_Сабиров_Зульфат");
        }

        private void BnnTask_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Назмутдинов Рузаль Ильгизович", "4435_Назмутдинов_Рузаль");
        }

        private void Read_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new
            Excel.Application();
            Excel.Workbook ObjWorkBook =
            ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                 for (int i = 0; i < _rows; i++)
                 list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
             ObjWorkExcel.Quit();
             GC.Collect();
            using (LR2Entities usersEntities = new LR2Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.User.Add(new User()
                    {
                        Должность = list[i, 1],
                        ФИО = list[i, 2],
                        Логин = list[i, 3],
                        Пароль = list[i, 4],
                        Последний_вход= list[i, 5],
                        Тип_входа= list[i, 6]

                    });
                }
                usersEntities.SaveChanges();
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            Dictionary <string, List <User>> keyValues = new Dictionary<string, List<User>>();
            using (LR2Entities usersEntities = new LR2Entities())
            {
                if (usersEntities.User.FirstOrDefault() == null)
                {
                    MessageBox.Show("База данных пуста!");
                    return;
                }
                foreach (User em in usersEntities.User)
                {
                    if (!keyValues.ContainsKey(em.Должность))
                    {
                        keyValues.Add(em.Должность, new List<User>() { em });
                    }
                    else
                    {
                        keyValues[em.Должность].Add(em);
                    }
                }
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == false)
                return;

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = keyValues.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < keyValues.Count(); i++)
            {
                string key = keyValues.Keys.ToArray()[i];
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Cells[1][1] = "Id";
                worksheet.Cells[2][1] = "Фио";
                worksheet.Cells[3][1] = "Логин";
                int j = 2;
                foreach (User emp in keyValues[key])
                {
                    worksheet.Cells[1][j] = emp.Код_сотрудника.ToString();
                    worksheet.Cells[2][j] = emp.ФИО;
                    worksheet.Cells[3][j] = emp.Логин;
                    j++;
                }
                worksheet.Columns.AutoFit();
                worksheet.Name = key;

            }

            if (saveFileDialog.FileName != "")
            {
                workbook.SaveAs(saveFileDialog.FileName);
                workbook.Close();
                Process.Start(saveFileDialog.FileName);
            }

            app.Quit();
        }
    
    }
}
