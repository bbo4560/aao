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
            if (!string.IsNullOrWhiteSpace(txtPanelID.Text))
                QueryPanelID = txtPanelID.Text.Trim();

            if (!string.IsNullOrWhiteSpace(txtLOTID.Text))
                QueryLOTID = txtLOTID.Text.Trim();

            if (!string.IsNullOrWhiteSpace(txtCarrierID.Text))
                QueryCarrierID = txtCarrierID.Text.Trim();

            if (dpDate.SelectedDate.HasValue)
                QueryDate = dpDate.SelectedDate.Value.Date;

            if (!string.IsNullOrWhiteSpace(txtTime.Text))
            {
                string t = txtTime.Text.Trim();
                if (System.Text.RegularExpressions.Regex.IsMatch(t, @"^[0-9:]+$"))
                {
                    QueryTimeInput = t;
                }
                else
                {
                    MessageBox.Show("時間格式僅允許數字與冒號 (例如 14 或 14:30)", "格式錯誤");
                    return;
                }
            }
            DialogResult = true;
            Close();
        }
    }
}


