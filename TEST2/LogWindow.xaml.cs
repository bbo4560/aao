using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace TEST2
{
    public partial class LogWindow : Window
    {
        private readonly DatabaseService _dbService;
        private ObservableCollection<SystemLog> _logs = new ObservableCollection<SystemLog>();

        public LogWindow(DatabaseService dbService)
        {
            InitializeComponent();

            ExcelPackage.License.SetNonCommercialPersonal("User");


            _dbService = dbService;
            logGrid.ItemsSource = _logs;
            Loaded += LogWindow_Loaded;
        }

        private async void LogWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadLogsAsync();
        }

        private async Task LoadLogsAsync()
        {
            try
            {
                var logs = await _dbService.GetLogsAsync();
                _logs.Clear();
                foreach (var log in logs)
                {
                    _logs.Add(log);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入紀錄失敗: {ex.Message}", "錯誤");
            }
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_logs == null || _logs.Count == 0)
            {
                MessageBox.Show("目前沒有資料可供匯出。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel 檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*",
                FileName = $"Log.xlsx",
                Title = "匯出 Excel"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    this.IsEnabled = false;
                    Mouse.OverrideCursor = Cursors.Wait;

                    string filePath = saveFileDialog.FileName;

                    await Task.Run(() => ExportToExcel(filePath));

                    await _dbService.InsertLogAsync(new SystemLog
                    {
                        OperationTime = DateTime.Now,
                        UserName = Environment.UserName,
                        MachineName = Environment.MachineName,
                        OperationType = "紀錄匯出",
                        AffectedData = $"操作紀錄\n檔案：{Path.GetFileName(filePath)}",
                        DetailDescription = $"{_logs.Count} 筆紀錄\n{Path.GetFullPath(filePath)}"
                    });


                    await LoadLogsAsync();

                    MessageBox.Show($"匯出成功！\n檔案路徑：{filePath}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"匯出失敗：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    this.IsEnabled = true;
                    Mouse.OverrideCursor = null;
                }
            }
        }


        private void ExportToExcel(string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("操作紀錄");

                ws.Cells[1, 1].Value = "操作時間";
                ws.Cells[1, 2].Value = "機台名稱";
                ws.Cells[1, 3].Value = "操作類型";
                ws.Cells[1, 4].Value = "被操作資料";
                ws.Cells[1, 5].Value = "詳細描述";

                using (var range = ws.Cells[1, 1, 1, 5])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                int row = 2;
                foreach (var log in _logs)
                {
                    ws.Cells[row, 1].Value = log.OperationTime.ToString("yyyy/MM/dd HH:mm:ss");
                    ws.Cells[row, 2].Value = log.MachineName;
                    ws.Cells[row, 3].Value = log.OperationType;
                    ws.Cells[row, 4].Value = log.AffectedData;
                    ws.Cells[row, 5].Value = log.DetailDescription;
                    row++;
                }

                ws.Cells.AutoFitColumns();

                var fileInfo = new FileInfo(filePath);
                package.SaveAs(fileInfo);
            }
        }

        private async void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            bool isVerified = ShowPasswordConfirmDialog(
                "確定要清除所有操作紀錄嗎？\n此動作無法復原！",
                "確認清除");

            if (!isVerified) return;

            try
            {
                await _dbService.ClearAllLogsAsync();

                await _dbService.InsertLogAsync(new SystemLog
                {
                    OperationTime = DateTime.Now,
                    UserName = Environment.UserName,
                    MachineName = Environment.MachineName,
                    OperationType = "清除紀錄",
                    AffectedData = "All",
                    DetailDescription = "使用者手動清除了所有歷史紀錄"
                });

                await LoadLogsAsync();

                MessageBox.Show("紀錄已清除完畢。", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"清除失敗: {ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ShowPasswordConfirmDialog(string message, string btnText)
        {
            var dialog = new Window
            {
                Title = "安全驗證",
                Width = 360,
                SizeToContent = SizeToContent.Height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(250, 250, 252)),
                Owner = this
            };

            var mainGrid = new Grid { Margin = new Thickness(25) };
            var contentStack = new StackPanel();

            contentStack.Children.Add(new TextBlock 
            {
                Text = "雙重驗證",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = System.Windows.Media.Brushes.Black
            });

            contentStack.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 13,
                Foreground = System.Windows.Media.Brushes.Gray,
                Margin = new Thickness(0, 8, 0, 15),
                TextWrapping = TextWrapping.Wrap
            });

            var passwordBox = new PasswordBox
            {
                Margin = new Thickness(0, 0, 0, 25),
                Height = 32,
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(5, 0, 5, 0),
                FontSize = 14
            };
            contentStack.Children.Add(passwordBox);

            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 0, 5)
            };

            var btnCancel = new Button
            {
                Content = "取消",
                Width = 80,
                Height = 32,
                IsCancel = true,
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = System.Windows.Media.Brushes.LightGray,
                Foreground = System.Windows.Media.Brushes.DimGray,
                Margin = new Thickness(0, 0, 10, 0)
            };

            var btnOk = new Button
            {
                Content = btnText,
                Width = 90,
                Height = 32,
                IsDefault = true,
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 77, 79)),
                Foreground = System.Windows.Media.Brushes.White,
                FontWeight = FontWeights.Bold,
                BorderThickness = new Thickness(0)
            };

            bool result = false;

            btnOk.Click += (s, e) =>
            {
                if (passwordBox.Password == "1234")
                {
                    result = true;
                    dialog.Close();
                }
                else
                {
                    MessageBox.Show("密碼錯誤！", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                    passwordBox.Clear();
                    passwordBox.Focus();
                }
            };

            btnCancel.Click += (s, e) => dialog.Close();
            btnPanel.Children.Add(btnCancel);
            btnPanel.Children.Add(btnOk);

            contentStack.Children.Add(btnPanel);
            mainGrid.Children.Add(contentStack);

            dialog.Content = mainGrid;

            dialog.Loaded += (s, e) => passwordBox.Focus();

            dialog.ShowDialog();
            return result;
        }
    }
}



