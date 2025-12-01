using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows;

namespace TEST2
{
    public partial class LogWindow : Window
    {
        private readonly DatabaseService _dbService;
        private ObservableCollection<SystemLog> _logs = new ObservableCollection<SystemLog>();

        public LogWindow(DatabaseService dbService)
        {
            InitializeComponent();
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

        private async void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "確定要清除所有操作紀錄嗎？\n此動作無法復原！",
                "清除確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes) return;

            try
            {
                await _dbService.ClearAllLogsAsync();
                await _dbService.InsertLogAsync(new SystemLog
                {
                    OperationTime = DateTime.Now,
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
    }
}


