using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace TEST2
{
    public partial class MainWindow : Window
    {
        private readonly DatabaseService _dbService;
        private ObservableCollection<OperationRecord> _records = new ObservableCollection<OperationRecord>();
        private readonly string ConnectionString = ConfigurationManager.ConnectionStrings["MyDbConnection"].ConnectionString;

        public MainWindow()
        {
            InitializeComponent();
            this.Title = Environment.MachineName;
            _dbService = new DatabaseService(ConnectionString);
            dataGrid.ItemsSource = _records;
            Loaded += MainWindow_Loaded;
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                DatabaseService.EnsureDatabaseExists(ConnectionString);
                await _dbService.InitializeDatabaseAsync();
                await ReloadDataAndUpdateTimeAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Initialization Error: {ex.Message}", "Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateLastOpTime()
        {
            txtLastOpTime.Text = $"最後操作時間: {DateTime.Now:yyyy/MM/dd HH:mm:ss}";
        }

        private async Task LoadDataAsync()
        {
            try
            {
                var data = await _dbService.GetAllAsync();
                _records.Clear();
                foreach (var item in data)
                {
                    _records.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ReloadDataAndUpdateTimeAsync()
        {
            await LoadDataAsync();
            UpdateLastOpTime();
        }

        private async Task ExecuteDbOperationAsync(Func<Task> operation, string? successMessage = null)
        {
            try
            {
                await operation();
                await ReloadDataAndUpdateTimeAsync();

                if (!string.IsNullOrEmpty(successMessage))
                {
                    MessageBox.Show(successMessage, "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"發生錯誤: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                BtnAdd_Click(BtnAdd, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.F2)
            {
                BtnQuery_Click(BtnQuery, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.F3)
            {
                BtnImport_Click(BtnImport, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.F4)
            {
                BtnExport_Click(BtnExport, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.F5)
            {
                BtnShowAll_Click(BtnShowAll, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.F6)
            {
                BtnLog_Click(BtnLog, new RoutedEventArgs());
                e.Handled = true;
            }
        }

        private void DataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                BtnBatchDelete_Click(BtnBatchDelete, new RoutedEventArgs());
                e.Handled = true;
            }
        }

        private async void BtnShowAll_Click(object sender, RoutedEventArgs e)
        {
            await ReloadDataAndUpdateTimeAsync();
        }

        private async void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            var addWindow = new AddRecordWindow
            {
                Owner = this
            };

            if (addWindow.ShowDialog() == true)
            {
                var newRecord = addWindow.NewRecord;
                if (newRecord != null)
                {
                    await ExecuteDbOperationAsync(
                        () => _dbService.InsertAsync(newRecord)
                    );
                }
            }
        }

        private async void BtnQuery_Click(object sender, RoutedEventArgs e)
        {
            var queryWindow = new QueryWindow { Owner = this };

            if (queryWindow.ShowDialog() == true)
            {
                try
                {
                    var data = await _dbService.GetFilteredAsync(
                        queryWindow.QueryPanelID,
                        queryWindow.QueryLOTID,
                        queryWindow.QueryCarrierID,
                        queryWindow.QueryDate,
                        queryWindow.QueryTimeInput
                    );

                    _records.Clear();

                    if (data == null || !data.Any())
                    {
                        MessageBox.Show("查無符合條件的資料。", "查詢結果",MessageBoxButton.OK, MessageBoxImage.Information);
                        await ReloadDataAndUpdateTimeAsync();
                        return;
                    }

                    foreach (var item in data)
                        _records.Add(item);

                    UpdateLastOpTime();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"查詢失敗：{ex.Message}", "Error",MessageBoxButton.OK, MessageBoxImage.Error);
                    await ReloadDataAndUpdateTimeAsync();
                }
            }
        }

        private async void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    var result = await _dbService.ImportFromExcelAsync(openFileDialog.FileName);
                    await ReloadDataAndUpdateTimeAsync();
                    MessageBox.Show($"匯入成功!\n\n{result}", "Success");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"匯入失敗: {ex.Message}", "Error");
                }
            }
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_records.Count == 0)
            {
                MessageBox.Show("目前沒有資料可供匯出。", "提示",MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var format = MessageBox.Show("要匯出成 Excel (.xlsx) 嗎？\n\n按「是」= Excel\n按「否」= SQLite 資料庫 (.db)",
                                         "選擇匯出格式",
                                         MessageBoxButton.YesNoCancel,
                                         MessageBoxImage.Question);

            if (format == MessageBoxResult.Cancel)
                return;

            bool exportExcel = (format == MessageBoxResult.Yes);

            var saveFileDialog = new SaveFileDialog
            {
                Filter = exportExcel
                    ? "Excel 檔案 (*.xlsx)|*.xlsx"
                    : "SQLite 資料庫 (*.db)|*.db",
                FileName = exportExcel ? "PanelLog.xlsx" : "PanelLog.db"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    BtnExport.IsEnabled = false;
                    Cursor = Cursors.Wait;

                    var dataToExport = _records.ToList();
                    string filePath = saveFileDialog.FileName;

                    if (exportExcel)
                    {
                        await _dbService.ExportToExcelAsync(dataToExport, filePath);
                    }
                    else
                    {
                        await Task.Run(() => _dbService.ExportPanelRecordsToDb(filePath, dataToExport));

                        await _dbService.InsertLogAsync(new SystemLog
                        {
                            OperationTime = DateTime.Now,
                            UserName = Environment.UserName,
                            MachineName = Environment.MachineName,
                            OperationType = "匯出 Panel 紀錄",
                            AffectedData = $"{dataToExport.Count} 筆",
                            DetailDescription = $"匯出至檔案路徑: {filePath}"
                        });
                    }

                    await ReloadDataAndUpdateTimeAsync();
                    MessageBox.Show($"匯出成功!\n檔案已儲存至: {filePath}",
                                    "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"匯出失敗: {ex.Message}",
                                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    BtnExport.IsEnabled = true;
                    Cursor = Cursors.Arrow;
                }
            }
        }

        private void BtnLog_Click(object sender, RoutedEventArgs e)
        {
            var logWindow = new LogWindow(_dbService)
            {
                Owner = this
            };
            logWindow.ShowDialog();
        }

        private async void BtnBatchDelete_Click(object sender, RoutedEventArgs e)
        {
            var selectedRecords = dataGrid.SelectedItems.Cast<OperationRecord>().ToList();

            if (selectedRecords.Count == 0)
            {
                return;
            }

            if (MessageBox.Show($"確定要刪除選取的 {selectedRecords.Count} 筆資料嗎?",
                                "確認",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                if (ShowPasswordConfirmDialog())
                {
                    var ids = selectedRecords.Select(r => r.Id);
                    await ExecuteDbOperationAsync(
                        () => _dbService.DeleteBatchAsync(ids),
                        "批量刪除成功!");
                }
            }
        }

        private async void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is OperationRecord record)
            {
                if (MessageBox.Show($"確定要刪除 PanelID: {record.PanelID} 嗎?", "確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    if (ShowPasswordConfirmDialog())
                    {
                        await ExecuteDbOperationAsync(
                            () => _dbService.DeleteAsync(record.Id)
                        );
                    }
                }
            }
        }

        private async void BtnModify_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is OperationRecord record)
            {
                var modifyWindow = new ModifyRecordWindow(record)
                {
                    Owner = this
                };

                if (modifyWindow.ShowDialog() == true)
                {
                    var updatedRecord = modifyWindow.ModifiedRecord;
                    if (updatedRecord != null)
                    {
                        await ExecuteDbOperationAsync(
                            () => _dbService.UpdateAsync(updatedRecord)
                        );
                    }
                }
            }
        }

        private bool ShowPasswordConfirmDialog()
        {
            var dialog = new Window
            {
                Title = "安全驗證",
                Width = 360,
                SizeToContent = SizeToContent.Height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Background = new SolidColorBrush(Color.FromRgb(250, 250, 252)),
                Owner = this
            };

            var mainGrid = new Grid { Margin = new Thickness(25) };
            var contentStack = new StackPanel();

            contentStack.Children.Add(new TextBlock
            {
                Text = "雙重驗證",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.Black
            });

            contentStack.Children.Add(new TextBlock
            {
                Text = "請輸入密碼以確認操作",
                FontSize = 13,
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 8, 0, 15)
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
                Background = Brushes.White,
                BorderBrush = Brushes.LightGray,
                Foreground = Brushes.DimGray,
                Margin = new Thickness(0, 0, 10, 0)
            };

            var btnOk = new Button
            {
                Content = "確認",
                Width = 90,
                Height = 32,
                IsDefault = true,
                Background = new SolidColorBrush(Color.FromRgb(255, 77, 79)),
                Foreground = Brushes.White,
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


