using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Project_for_san_com
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SanComRaportWindow sanComRaportWindow = new SanComRaportWindow();
            sanComRaportWindow.Show();
            this.Close();
            //WordDocumentCreator creator = new WordDocumentCreator();
            //creator.CreateAndFillWord("1",10);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            RaportN raportN = new RaportN();
            raportN.Show();
            this.Close();
        }
    }
}