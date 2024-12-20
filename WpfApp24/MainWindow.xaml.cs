using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.DataVisualization.Charting;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp24
{
    public partial class MainWindow : Window
    {
        private Aparxof_DB_PaymentEntities _context = new Aparxof_DB_PaymentEntities();
        private Chart ChartPayments; // Объект диаграммы

        public MainWindow()
        {
            InitializeComponent();

            // Создаем объект диаграммы
            ChartPayments = new Chart();
            ChartPayments.ChartAreas.Add(new ChartArea("Main"));

            // Добавляем диаграмму в контейнер WindowsFormsHost
            HostChart.Child = ChartPayments;

            // Заполняем список пользователей
            CmbUser.ItemsSource = _context.User.ToList();
            //CmbUser.DisplayMemberPath = "FIO";

            // Заполняем типы диаграмм
            CmbDiagram.ItemsSource = Enum.GetValues(typeof(SeriesChartType));
            //CmbDiagram.SelectedItem = SeriesChartType.Column;
        }

        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            if (CmbUser.SelectedItem is User currentUser && CmbDiagram.SelectedItem is SeriesChartType currentType)
            {
                var currentSeries = ChartPayments.Series.FirstOrDefault();
                if (currentSeries == null)
                {
                    currentSeries = new Series("Платежи") { IsValueShownAsLabel = true };
                    ChartPayments.Series.Add(currentSeries);
                }

                currentSeries.ChartType = currentType;
                currentSeries.Points.Clear();

                var categories = _context.Category.ToList();
                MessageBox.Show(currentUser.FIO);
                foreach (var category in categories)
                {
                    currentSeries.Points.AddXY(category.Name, _context.Payment.ToList()
                        .Where(p => p.User == currentUser && p.Category == category)
                        .Sum(p => p.Price * p.Num));

                }
            }
        }


        private void ExportToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                var excelApp = new Excel.Application();
                var workbook = excelApp.Workbooks.Add();
                var worksheet = (Excel.Worksheet)workbook.Sheets[1];

                worksheet.Cells[1, 1] = "Категория";
                worksheet.Cells[1, 2] = "Сумма";

                int row = 2;
                foreach (var point in ChartPayments.Series[0].Points)
                {
                    worksheet.Cells[row, 1] = point.AxisLabel;
                    worksheet.Cells[row, 2] = point.YValues[0];
                    row++;
                }

                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта в Excel: {ex.Message}");
            }
        }

        private void ExportToWord(object sender, RoutedEventArgs e)
        {
            try
            {
                var wordApp = new Word.Application();
                var document = wordApp.Documents.Add();
                var paragraph = document.Content.Paragraphs.Add();
                paragraph.Range.Text = "Отчет по платежам";

                var table = document.Tables.Add(paragraph.Range, ChartPayments.Series[0].Points.Count + 1, 2);
                table.Cell(1, 1).Range.Text = "Категория";
                table.Cell(1, 2).Range.Text = "Сумма";

                int row = 2;
                foreach (var point in ChartPayments.Series[0].Points)
                {
                    table.Cell(row, 1).Range.Text = point.AxisLabel;
                    table.Cell(row, 2).Range.Text = point.YValues[0].ToString();
                    row++;
                }

                wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта в Word: {ex.Message}");
            }
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            var result = MessageBox.Show("Вы действительно хотите выйти?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
            {
                e.Cancel = true;
            }
        }

        // Экспорт данных диаграммы в Excel
        private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем список всех пользователей, отсортированных по ФИО
                var allUsers = _context.User.ToList().OrderBy(u => u.FIO).ToList();

                // Создаем объект Excel
                var application = new Excel.Application
                {
                    SheetsInNewWorkbook = allUsers.Count(),
                    Visible = true // Чтобы приложение Excel было видно
                };

                // Добавляем новую книгу
                Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

                for (int i = 0; i < allUsers.Count(); i++)
                {
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                    worksheet.Name = allUsers[i].FIO; // Название листа - ФИО пользователя

                    // Заголовки столбцов
                    worksheet.Cells[1][startRowIndex] = "Дата платежа";
                    worksheet.Cells[2][startRowIndex] = "Название";
                    worksheet.Cells[3][startRowIndex] = "Стоимость";
                    worksheet.Cells[4][startRowIndex] = "Количество";
                    worksheet.Cells[5][startRowIndex] = "Сумма";

                    Excel.Range columnHeaderRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][1]];
                    columnHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    columnHeaderRange.Font.Bold = true;
                    startRowIndex++;

                    // Группируем платежи по категориям
                    var userCategories = allUsers[i].Payment
                        .OrderBy(u => u.Date)
                        .GroupBy(u => u.Category)
                        .OrderBy(u => u.Key.Name);
                    
                    // Цикл по категориям платежей
                    foreach (var groupCategory in userCategories)
                    {
                        // Отображаем название категории
                        Excel.Range categoryHeaderRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[5][startRowIndex]];
                        categoryHeaderRange.Merge();
                        categoryHeaderRange.Value = groupCategory.Key.Name;
                        categoryHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        categoryHeaderRange.Font.Italic = true;
                        startRowIndex++;

                        // Цикл по платежам в категории
                        foreach (var payment in groupCategory)
                        {
                            worksheet.Cells[1][startRowIndex] = payment.Date.ToString("dd.MM.yyyy");
                            worksheet.Cells[2][startRowIndex] = payment.Name;
                            worksheet.Cells[3][startRowIndex] = payment.Price;
                            (worksheet.Cells[3][startRowIndex] as Excel.Range).NumberFormat = "0.00";
                            worksheet.Cells[4][startRowIndex] = payment.Num;
                            worksheet.Cells[5][startRowIndex].Formula = $"=C{startRowIndex}*D{startRowIndex}";
                            (worksheet.Cells[5][startRowIndex] as Excel.Range).NumberFormat = "0.00";
                            startRowIndex++;
                        }

                        // Добавляем итог по категории
                        Excel.Range sumRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[4][startRowIndex]];
                        sumRange.Merge();
                        sumRange.Value = "ИТОГО:";
                        sumRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        worksheet.Cells[5][startRowIndex].Formula = $"=SUM(E{startRowIndex - groupCategory.Count()}:" +
                                                                  $"E{startRowIndex - 1})";
                        sumRange.Font.Bold = worksheet.Cells[5][startRowIndex].Font.Bold = true;
                        startRowIndex++;
                    }

                    // Добавляем границы для таблицы
                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]];
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                        Excel.XlLineStyle.xlContinuous;

                    // Автоматическая подгонка ширины столбцов
                    worksheet.Columns.AutoFit();
                }

                // Добавление общего итога для всех пользователей на отдельном листе
                Excel.Worksheet totalWorksheet = application.Worksheets.Add();
                totalWorksheet.Name = "Общий итог";
                int rowIndex = 1;

                totalWorksheet.Cells[1][rowIndex] = "Общий итог:";
                totalWorksheet.Cells[2][rowIndex] = "Сумма всех платежей:";

                // Подсчитываем общую сумму по всем пользователям
                totalWorksheet.Cells[3][rowIndex].Formula = "=SUM(" + string.Join(",", allUsers.Select((user, index) => $"'{user.FIO}'!E2:E{index + 2}")) + ")";
                (totalWorksheet.Cells[3][rowIndex] as Excel.Range).NumberFormat = "0.00";

                // Форматируем строку общего итога красным цветом (RGB)
                Excel.Range totalRange = totalWorksheet.Range[totalWorksheet.Cells[1][rowIndex], totalWorksheet.Cells[3][rowIndex]];
                totalRange.Font.Color = 255; // Красный цвет

                // Отображаем Excel
                application.Visible = true;
            }
            catch (Exception ex)
            {
                // Обработка ошибок
                System.Windows.MessageBox.Show($"Произошла ошибка: {ex.Message}");
            }
        }



        // Экспорт данных диаграммы в Word
        private void ExportToWordButton_Click(object sender, EventArgs e)
        {
            var allUsers = _context.User.ToList();  // Получаем список пользователей
            var allCategories = _context.Category.ToList();  // Получаем список категорий

            // Создаем новый документ Word
            var application = new Word.Application();
            Word.Document document = application.Documents.Add();

            // Добавляем верхний колонтитул с текущей датой (Проверка наличия колонтитула)
            if (document.Sections.Count > 0)
            {
                var headerRange = document.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = $"Отчет о платежах на {DateTime.Now.ToString("dd.MM.yyyy")}";
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

            // Перебираем пользователей
            foreach (var user in allUsers)
            {
                // Проверяем, выбран ли пользователь в ComboBox
                if (CmbUser.SelectedItem != null && user == CmbUser.SelectedItem as User)
                {
                    // Добавляем абзац с именем пользователя
                    Word.Paragraph userParagraph = document.Paragraphs.Add();
                    Word.Range userRange = userParagraph.Range;
                    userRange.Text = user.FIO;
                    userParagraph.set_Style("Заголовок");
                    userRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    userRange.InsertParagraphAfter();
                    document.Paragraphs.Add(); //Пустая строка


                    // Добавляем пустую строку
                    document.Paragraphs.Add();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 2);
                    paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle =
                             Word.WdLineStyle.wdLineStyleSingle;
                    paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;


                    // Названия колонок
                    Word.Range cellRange;

                    cellRange = paymentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Категория";
                    cellRange = paymentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Сумма расходов";

                    paymentsTable.Rows[1].Range.Font.Name = "Times New Roman";
                    paymentsTable.Rows[1].Range.Font.Size = 14;
                    paymentsTable.Rows[1].Range.Bold = 1;
                    paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;


                    // Заполнение таблицы данными для выбранного пользователя
                    for (int i = 0; i < allCategories.Count(); i++)
                    {
                        var currentCategory = allCategories[i];
                        cellRange = paymentsTable.Cell(i + 2, 1).Range;
                        cellRange.Text = currentCategory.Name;
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;

                        cellRange = paymentsTable.Cell(i + 2, 2).Range;
                        cellRange.Text = user.Payment.ToList().
                   Where(u => u.Category == currentCategory).Sum(u => u.Num * u.Price).ToString("N2") + " руб.";
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;
                    } //завершение цикла по строкам таблицы
                    document.Paragraphs.Add(); //пустая строка


                    // Добавляем информацию о самом дорогом платеже
                    Payment maxPayment = user.Payment.OrderByDescending(u => u.Price * u.Num).FirstOrDefault();
                    if (maxPayment != null)
                    {
                        Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                        Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                        maxPaymentRange.Text = $"Самый дорогостоящий платеж - {maxPayment.Name} за {(maxPayment.Price * maxPayment.Num).ToString("N2")} " + $"руб. от {maxPayment.Date.ToString("dd.MM.yyyy")}";
                        maxPaymentParagraph.set_Style("Подзаголовок");
                        maxPaymentRange.Font.Color = Word.WdColor.wdColorDarkRed;
                        maxPaymentRange.InsertParagraphAfter();
                    }
                    document.Paragraphs.Add(); //пустая строка


                    // Добавляем информацию о самом дешевом платеже
                    Payment minPayment = user.Payment.OrderBy(u => u.Price * u.Num).FirstOrDefault();
                    if (maxPayment != null)
                    {
                        Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                        Word.Range minPaymentRange = minPaymentParagraph.Range;
                        minPaymentRange.Text = $"Самый дешевый платеж - {minPayment.Name} за {(minPayment.Price * minPayment.Num).ToString("N2")} " + $"руб. от {minPayment.Date.ToString("dd.MM.yyyy")}";
                        minPaymentParagraph.set_Style("Подзаголовок");
                        minPaymentRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                        minPaymentRange.InsertParagraphAfter();
                    }


                    // Добавляем разрыв страницы, если это не последний пользователь
                    if (user != allUsers.LastOrDefault()) document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
            }

            // Добавляем нижний колонтитул с номером страницы
            var footerRange = document.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
            footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // Открываем документ
            application.Visible = true;

            // Сохраняем документ в формате .docx и .pdf
            document.SaveAs2(@"C:\Users\user\Documents\Payments.docx");
            document.SaveAs2(@"C:\Users\user\Documents\Payments.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }
    }
}
