using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Prism.Commands;
using System.Collections.ObjectModel;
using System.Windows.Forms;
using System.Windows.Controls;

namespace Parser
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ChangeInfoWindow changeInfoWindow;
        private XlFile file { get; }
        

        public MainWindow()
        {
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            file = new XlFile();

            InitializeComponent();
            lvSheet.ItemsSource = XlFile.Sheet;
            ShortInfoList.ItemsSource = XlFile.Sheet;
        }

        private void ParseButton_Click(object sender, RoutedEventArgs e)
        {
            
            if (changeInfoWindow != null)
            {
                changeInfoWindow.Close();
            }
            changeInfoWindow = new ChangeInfoWindow(file.UpdateTable());
            changeInfoWindow.Show();

            
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ShortInfoButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            file.SaveTable();
        }


    }
}