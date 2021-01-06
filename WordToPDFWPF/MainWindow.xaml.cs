using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace WordToPDFWPF
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private HashSet<string> pathList = new HashSet<string>();
        public MainWindow()
        {
            InitializeComponent();
            this.fileList.ItemsSource = pathList;
            this.progressBarStack.Visibility = Visibility.Hidden;
            
        }

        private void Transform_Button_Click(object sender, RoutedEventArgs e)
        {
            if (pathList.Count <= 0)
            {
                return;
            }

            showProcessBar(pathList.Count);
            new Thread(() =>
            {
                foreach (string path in pathList)
                {
                    //Console.WriteLine(path);
                    if (!Utils.WordToPDF(path))
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show(path + "  失败", "警告");
                        });

                    }
                    this.Dispatcher.Invoke(() =>
                    {
                        this.progressBar.Value += 1;
                        this.progressBarValue.Content = this.progressBar.Value;
                    });
                };
                this.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show("转换完成！！！", "提示");
                });

            }).Start();
        }
        private void showProcessBar(int maximum)
        { 
            this.progressBarStack.Visibility = Visibility.Visible;
            this.progressBar.Maximum = pathList.Count;
            this.progressBar.Value = 0;
            this.progressBarValue.Content = 0;
            this.progressBarMax.Content = maximum;
        }

        private void fileList_Drop(object sender, DragEventArgs e)
        {
            String[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (String s in files)
            {
                if (File.Exists(s))
                {
                    string extension = s.Substring(s.LastIndexOf('.') + 1);
                    if (extension.Equals("doc") || extension.Equals("docx"))
                    {
                        pathList.Add(s);
                    }

                }
            }
            this.fileList.Items.Refresh();
        }

        private void Clear_Button_Click(object sender, RoutedEventArgs e)
        {
            pathList.Clear();
            this.fileList.Items.Refresh();

        }
    }
}
