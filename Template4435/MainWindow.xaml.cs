using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


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
            using (Entities usersEntities = new Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.User.Add(new User()
                    {
                        Position = list[i, 1],
                        FullName = list[i, 2],
                        Log = list[i, 3],
                        Password = list[i, 4],
                        LastEnter= list[i, 5],
                        TypeEnter= list[i, 6]

                    });
                }
                usersEntities.SaveChanges();
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            Dictionary <string, List <User>> keyValues = new Dictionary<string, List<User>>();
            using (Entities usersEntities = new Entities())
            {
                if (usersEntities.User.FirstOrDefault() == null)
                {
                    MessageBox.Show("База данных пуста!");
                    return;
                }
                foreach (User em in usersEntities.User)
                {
                    if (!keyValues.ContainsKey(em.Position))
                    {
                        keyValues.Add(em.Position, new List<User>() { em });
                    }
                    else
                    {
                        keyValues[em.Position].Add(em);
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
                worksheet.Cells[2][1] = "FullName";
                worksheet.Cells[3][1] = "Log";
                int j = 2;
                foreach (User emp in keyValues[key])
                {
                    worksheet.Cells[1][j] = emp.CodeStaff.ToString();
                    worksheet.Cells[2][j] = emp.FullName;
                    worksheet.Cells[3][j] = emp.Log;
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

        private void ExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<User> allServices;
            using (Entities usersEntities = new Entities())
            {
                allServices = usersEntities.User.ToList().OrderBy(s => s.Position).ToList();
            }
            var users = allServices.OrderBy(o => o.Position).GroupBy(s => s.CodeStaff)
                    .ToDictionary(g => g.Key, g => g.Select(s => new { s.CodeStaff, s.Position, s.FullName, s.Log, s.Password, s.LastEnter, s.TypeEnter }).ToArray());
            var app = new Word.Application();
            Word.Document document = app.Documents.Add();
            for (int i = 0; i < 3; i++)
            {
                var data = i == 0 ? users.Where(w => w.Value.All(p => p.Position.Equals("Администратор")))
                         : i == 1 ? users.Where(w => w.Value.All(p => p.Position.Equals("Старший смены")))
                         : i == 2 ? users.Where(w => w.Value.All(p => p.Position.Equals("Продавец"))) : users;
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = $"Категория {i + 1}";
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                var studentsTable = document.Tables.Add(tableRange, data.Select(s => s.Value.Length).Sum() + 1, 3);
                studentsTable.Borders.InsideLineStyle = studentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                studentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Word.Range cellRange;
                cellRange = studentsTable.Cell(1, 1).Range;
                cellRange.Text = "Код сотрудника";
                cellRange = studentsTable.Cell(1, 2).Range;
                cellRange.Text = "Фио";
                cellRange = studentsTable.Cell(1, 3).Range;
                cellRange.Text = "Логин";
                studentsTable.Rows[1].Range.Bold = 1;
                studentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                int row = 1;
                var stepSize = 1;
                foreach (var group in data)
                {
                    foreach (var currentCost in group.Value)
                    {
                        cellRange = studentsTable.Cell(row + stepSize, 1).Range;
                        cellRange.Text = currentCost.CodeStaff.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = studentsTable.Cell(row + stepSize, 2).Range;
                        cellRange.Text = currentCost.FullName;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = studentsTable.Cell(row + stepSize, 3).Range;
                        cellRange.Text = currentCost.Log.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        row++;
                    }
                }
                Word.Paragraph countCostsParagraph = document.Paragraphs.Add();
                Word.Range countCostsRange = countCostsParagraph.Range;
                countCostsRange.Text = $"Количество сотрудников - {data.Select(s => s.Value.Length).Sum()} ";
                countCostsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                countCostsRange.InsertParagraphAfter();
                document.Words.Last.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            }
            app.Visible = true;
            document.SaveAs(@"E:\outputFileWord.docx");
            document.SaveAs(@"E:\outputFilePdf.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }

        private async void JSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Spisok.json)|*.json",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate);
            List<User> users = await JsonSerializer.DeserializeAsync<List<User>>(fs);

            using (Entities db = new Entities())
            {
                foreach (User user in users)
                {
                    User user1 = new User();
                    user1.CodeStaff = user.CodeStaff;
                    user1.Position = user.Position;
                    user1.FullName = user.FullName;
                    user1.Log = user.Log;
                    user1.Password = user.Password;
                    user1.LastEnter = user.LastEnter;
                    user1.TypeEnter = user.TypeEnter;
                    db.User.Add(user1);
                }
                db.SaveChanges();
            }
        }
    }
}
