using System.Windows;
using System.Windows.Controls;

namespace GenerateMasterDataForExcel
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void GetChangedText(object sender, TextChangedEventArgs e)
        {
            this.TextField.Text = EnemyNameTextField.Text;
        }
    }
}