using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Nico2PDF.Views
{
    public partial class CategoryEditDialog : Window
    {
        public string CategoryName { get; private set; } = "";
        public string CategoryIcon { get; private set; } = "📁";
        public string CategoryColor { get; private set; } = "#E9ECEF";

        public CategoryEditDialog(string? name, string? icon = null, string? color = null)
        {
            InitializeComponent();
            
            // 既存カテゴリの編集か新規作成かを判定
            bool isEditing = !string.IsNullOrEmpty(name);
            Title = isEditing ? "カテゴリ編集" : "新規カテゴリ作成";
            
            CategoryName = name ?? "新しいカテゴリ";
            CategoryIcon = icon ?? "📁";
            CategoryColor = color ?? "#E9ECEF";
            
            LoadCategoryInfo();
            UpdatePreview();
        }

        private void LoadCategoryInfo()
        {
            txtCategoryName.Text = CategoryName;
            txtIcon.Text = CategoryIcon;
            
            try
            {
                var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(CategoryColor);
                btnColorPreview.Background = new SolidColorBrush(color);
            }
            catch
            {
                btnColorPreview.Background = new SolidColorBrush(Colors.LightGray);
            }

            // イベントハンドラを追加
            txtCategoryName.TextChanged += (s, e) => UpdatePreview();
            txtIcon.TextChanged += (s, e) => UpdatePreview();
        }

        private void UpdatePreview()
        {
            previewIcon.Text = txtIcon.Text;
            previewName.Text = txtCategoryName.Text;
            
            if (btnColorPreview?.Background != null)
            {
                previewBorder.Background = btnColorPreview.Background;
            }
        }

        private void BtnSelectIcon_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button)
            {
                txtIcon.Text = button.Content.ToString();
                UpdatePreview();
            }
        }

        private void BtnSelectColor_Click(object sender, RoutedEventArgs e)
        {
            var colorDialog = new System.Windows.Forms.ColorDialog();
            
            try
            {
                var currentBrush = btnColorPreview.Background as SolidColorBrush;
                if (currentBrush != null)
                {
                    var currentColor = currentBrush.Color;
                    colorDialog.Color = System.Drawing.Color.FromArgb(currentColor.A, currentColor.R, currentColor.G, currentColor.B);
                }
            }
            catch { }

            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var newColor = colorDialog.Color;
                var wpfColor = System.Windows.Media.Color.FromArgb(newColor.A, newColor.R, newColor.G, newColor.B);
                btnColorPreview.Background = new SolidColorBrush(wpfColor);
                UpdatePreview();
            }
        }

        private void BtnPresetColor_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Background != null)
            {
                btnColorPreview.Background = button.Background;
                UpdatePreview();
            }
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            // 入力検証
            if (string.IsNullOrWhiteSpace(txtCategoryName.Text))
            {
                System.Windows.MessageBox.Show("カテゴリ名を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtCategoryName.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtIcon.Text))
            {
                System.Windows.MessageBox.Show("アイコンを入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtIcon.Focus();
                return;
            }

            // カテゴリ情報を更新
            CategoryName = txtCategoryName.Text.Trim();
            CategoryIcon = txtIcon.Text.Trim();
            
            if (btnColorPreview.Background is SolidColorBrush brush)
            {
                var color = brush.Color;
                CategoryColor = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }

            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}