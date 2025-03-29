using DocumentsLibrary;
using System.Windows;



namespace WordTask
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void AddGameButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string fileName = "C:\\Users\\slavv\\Downloads\\Docs\\Приказ.pdf";

                string subject = "Назначение разработчиком по разработке игры\n";
                string body = $"Добрый день, коллеги.\nВы назначены разработчиком новой игры \"Minecraft\"! Просим ознакомиться с приказом и приступить к работе.\nС уважением, HR-специалист.";

                if (!string.IsNullOrWhiteSpace(CustomBodyTextBox.Text))
                    body = CustomBodyTextBox.Text;


                using (WordService wordService = new WordService())
                    wordService.CreatePdf(fileName, subject, body);

                MessageBox.Show("Приказ успешно создан");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }





        private async void RenerateReportButton_Click(object sender, RoutedEventArgs e)
        {
            //string fileName = "C:\\Users\\slavv\\Downloads\\Docs\\Приказ.pdf";

            //List<Game> games = await _gameService.GetAllAsync();

            //GenerateExcel(fileName, games);
        }

        //private void GenerateExcel(string fileName, List<Game> games)
        //{
        //    try
        //    {
        //        var excelApp = new Excel.Application();
        //        excelApp.Visible = true;
        //        var workbook = excelApp.Workbooks.Add();

        //        var worksheet = workbook.Worksheets[1];
        //        worksheet.Name = "Отчет по количетсву игр";
        //        for (int i = 0; i < games.Count; i++)
        //            for (int j = 0;


        //        //workbook.SaveAs(fileName, Word.WdSaveFormat.wdFormatPDF);
        //        workbook.Close(false);
        //        excelApp.Quit();

        //        MessageBox.Show("Файл Excel отчета создан успешно!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Ошибка при создании Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        //    }
        //}
    }
}