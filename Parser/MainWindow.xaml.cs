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
        private PagingCollectionView pagingCollectionView;
        private string messageUpdate = "";


        public MainWindow()
        {
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            file = new XlFile();

            InitializeComponent();

            this.pagingCollectionView = new PagingCollectionView(XlFile.Sheet);
            this.DataContext = this.pagingCollectionView;
        }

        private void ParseButton_Click(object sender, RoutedEventArgs e)
        {
            if (changeInfoWindow != null)
            {
                changeInfoWindow.Close();
            }
            messageUpdate = file.UpdateTable();
            changeInfoWindow = new ChangeInfoWindow(messageUpdate);
            changeInfoWindow.Show();

            pagingCollectionView.Refresh();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            file.SaveTable();
        }

        private void ShortInfoButton_Click(object sender, RoutedEventArgs e)
        {
            pagingCollectionView.CurrentNote = DataGrid1.SelectedItem.ToString();
        }

        private void HistoryButton_Click(object sender, RoutedEventArgs e)
        {
            if(changeInfoWindow != null)
            {
                changeInfoWindow.Close();
            }
            changeInfoWindow = new ChangeInfoWindow(messageUpdate);
            changeInfoWindow.Show();
        }

        private void OnNextClicked(object sender, RoutedEventArgs e)
        {
            this.pagingCollectionView.MoveToNextPage();
        }

        private void OnPreviousClicked(object sender, RoutedEventArgs e)
        {
            this.pagingCollectionView.MoveToPreviousPage();
        }

        private void DataGrid1_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            pagingCollectionView.CurrentNote = DataGrid1.SelectedItem.ToString();
        }
    }
}