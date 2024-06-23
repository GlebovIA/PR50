using Microsoft.Win32;
using PR50.Contexts;
using System.Windows;

namespace PR50
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

        private void Report(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word Files (*.docx)|*.docx";
            sfd.ShowDialog();
            if (sfd.FileName != "") OwnerContext.Report(sfd.FileName);
        }
        public void LoadRooms()
        {
            for (int i = 1; i < 20; i++)
                Parent.Children.Add(new Elements.Room(i));
        }
    }
}
