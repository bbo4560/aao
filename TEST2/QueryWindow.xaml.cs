using System;
using System.Windows;

namespace TEST2
{
    public partial class QueryWindow : Window
    {
        public string? QueryPanelID { get; private set; }
        public string? QueryLOTID { get; private set; }
        public string? QueryCarrierID { get; private set; }

        public DateTime? QueryDate { get; private set; }

        public string? QueryTimeInput { get; private set; }

        public QueryWindow()
        {
            InitializeComponent();
        }

        private void BtnQuery_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(txtPanelID.Text))
                {
                    var text = txtPanelID.Text.Trim();
                    if (!long.TryParse(text, out _))
                        throw new ArgumentException("PanelID 必須為數字格式。");
                    QueryPanelID = text;
                }

                if (!string.IsNullOrWhiteSpace(txtLOTID.Text))
                {
                    var text = txtLOTID.Text.Trim();
                    if (!long.TryParse(text, out _))
                        throw new ArgumentException("LOTID 必須為數字格式。");
                    QueryLOTID = text;
                }

                if (!string.IsNullOrWhiteSpace(txtCarrierID.Text))
                {
                    var text = txtCarrierID.Text.Trim();
                    if (!long.TryParse(text, out _))
                        throw new ArgumentException("CarrierID 必須為數字格式。");
                    QueryCarrierID = text;
                }

                if (dpDate.SelectedDate.HasValue)
                    QueryDate = dpDate.SelectedDate.Value.Date;

                if (!string.IsNullOrWhiteSpace(txtTime.Text))
                {
                    string t = txtTime.Text.Trim();
                    if (!System.Text.RegularExpressions.Regex.IsMatch(t, @"^[0-9:]+$"))
                        throw new ArgumentException("時間格式僅允許數字與冒號 (例如 14 或 14:30)");
                    QueryTimeInput = t;
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "查詢條件錯誤",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                DialogResult = false;   
                Close();                
            }
        }
    }
}


