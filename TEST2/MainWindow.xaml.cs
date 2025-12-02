using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

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
            string machineName = Environment.MachineName;
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
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Initialization Error (Check Connection String): {ex.Message}", "Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            UpdateLastOpTime();
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

        private async void BtnShowAll_Click(object sender, RoutedEventArgs e)
        {
            await LoadDataAsync();
            UpdateLastOpTime();
        }

        private async void BtnAdd_Click(object sender, RoutedEventArgs e)

        {
            var addWindow = new AddRecordWindow();
            addWindow.Owner = this;

            if (addWindow.ShowDialog() == true)
            {
                var newRecord = addWindow.NewRecord;
                if (newRecord != null)
                {
                    try
                    {
                        await _dbService.InsertAsync(newRecord);
                        await LoadDataAsync();
                        UpdateLastOpTime();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error adding record: {ex.Message}");
                    }
                }
            }
        }

        private async void BtnQuery_Click(object sender, RoutedEventArgs e)
        {
            var queryWindow = new QueryWindow();
            queryWindow.Owner = this;

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
                    foreach (var item in data)
                    {
                        _records.Add(item);
                    }
                    UpdateLastOpTime();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"查詢失敗: {ex.Message}", "Error");
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
                    await LoadDataAsync();
                    MessageBox.Show($"匯入成功!\n\n{result}", "Success");
                    UpdateLastOpTime();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"匯入失敗: {ex.Message}", "Error");
                }
            }
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"PanelLog.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    BtnExport.IsEnabled = false;
                    Cursor = Cursors.Wait; 
                    await _dbService.ExportToExcelAsync(_records, saveFileDialog.FileName);

                    MessageBox.Show("匯出成功!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    UpdateLastOpTime(); 
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"匯出失敗: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
            var logWindow = new LogWindow(_dbService);
            logWindow.Owner = this;
            logWindow.ShowDialog();
        }

        private async void BtnBatchDelete_Click(object sender, RoutedEventArgs e)
        {
            var selectedRecords = dataGrid.SelectedItems.Cast<OperationRecord>().ToList();

            if (selectedRecords.Count == 0)
            {
                MessageBox.Show("請先選取要刪除的資料", "提示");
                return;
            }

            if (MessageBox.Show($"確定要刪除選取的 {selectedRecords.Count} 筆資料嗎?", "確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                if (ShowPasswordConfirmDialog())
                {
                    try
                    {
                        var ids = selectedRecords.Select(r => r.Id);
                        await _dbService.DeleteBatchAsync(ids);
                        await LoadDataAsync();
                        UpdateLastOpTime();
                        MessageBox.Show("批量刪除成功!", "Success");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"批量刪除失敗: {ex.Message}", "Error");
                    }
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
                        try
                        {
                            await _dbService.DeleteAsync(record.Id);
                            await LoadDataAsync();
                            UpdateLastOpTime();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"刪除失敗: {ex.Message}", "Error");
                        }
                    }
                }
            }
        }

        private async void BtnModify_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is OperationRecord record)
            {
                var modifyWindow = new ModifyRecordWindow(record);
                modifyWindow.Owner = this;

                if (modifyWindow.ShowDialog() == true)
                {
                    var updatedRecord = modifyWindow.ModifiedRecord;

                    if (updatedRecord != null)
                    {
                        try
                        {
                            await _dbService.UpdateAsync(updatedRecord);
                            await LoadDataAsync();
                            UpdateLastOpTime();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"修改失敗: {ex.Message}", "Error");
                        }
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
                Text = "請輸入密碼以確認刪除操作",
                FontSize = 13,
                Foreground = System.Windows.Media.Brushes.Gray,
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
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = System.Windows.Media.Brushes.LightGray,
                Foreground = System.Windows.Media.Brushes.DimGray,
                Margin = new Thickness(0, 0, 10, 0)
            };

            var btnOk = new Button
            {
                Content = "確認刪除",
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
            dialog.ShowDialog();
            return result;
        }
    }
}
