using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Windows.Shapes;

namespace Parser
{
    /// <summary>
    /// Interaction logic for ChangeInfoWindow.xaml
    /// </summary>
    public partial class ChangeInfoWindow : Window
    {
        private string message;
        public static ObservableCollection<Note> notesOld { get; set; } = new ObservableCollection<Note>();
        public static ObservableCollection<Note> notesNew { get; set; } = new ObservableCollection<Note>();
        //public static bool isClosed { get; set; } = false;

        public ChangeInfoWindow(string message)
        {
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.message = message;
            InitializeComponent();

            lvSheetOld.ItemsSource = notesOld;
            lvSheetNew.ItemsSource = notesNew;
            changedCount.Text = notesOld.Count.ToString();
            statusTextBox.Text = message;

            if (message == "Успешно")
            {
                statusTextBox.Foreground = Brushes.DarkOliveGreen;
            }
            else
            {
                statusTextBox.Foreground = Brushes.Red;
            }
            //this.Closed += MainWindow_Closed;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
