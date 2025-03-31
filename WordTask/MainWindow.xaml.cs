using DocumentsLibrary;
using DocumentsLibrary.Models;
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
                string fileName = "C:\\Users\\WSR\\Desktop\\Slava\\Docs\\Приказ.pdf";

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
            try
            {
                string fileName = "C:\\Users\\WSR\\Desktop\\Slava\\Docs\\Отчет.xlsx";

                using (ExcelService excelService = new ExcelService())
                    excelService.CreatReport(fileName, Model.Games);

                MessageBox.Show("Отчет успешно создан");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}