using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Nico2PDF.Models;
using Nico2PDF.Services;

namespace Nico2PDF.Views
{
    public partial class CategoryManageDialog : Window
    {
        private List<ProjectData> projects;
        private ObservableCollection<CategoryViewModel> categories;
        private bool hasChanges = false;

        public CategoryManageDialog(List<ProjectData> projectList)
        {
            InitializeComponent();
            projects = projectList;
            categories = new ObservableCollection<CategoryViewModel>();
            dgCategories.ItemsSource = categories;
            LoadCategories();
            UpdateStatistics();
        }

        private void LoadCategories()
        {
            categories.Clear();
            
            // プロジェクトからカテゴリ情報を抽出
            var categoryGroups = projects.GroupBy(p => string.IsNullOrEmpty(p.Category) ? "未分類" : p.Category);
            
            foreach (var group in categoryGroups.Where(g => g.Key != "未分類").OrderBy(g => g.Key))
            {
                var firstProject = group.First();
                categories.Add(new CategoryViewModel
                {
                    Name = group.Key,
                    OriginalName = group.Key,
                    Icon = firstProject.CategoryIcon,
                    Color = firstProject.CategoryColor,
                    ProjectCount = group.Count(),
                    IsDefault = false
                });
            }
            
            // 未分類カテゴリを追加（削除不可）
            var uncategorizedCount = projects.Count(p => string.IsNullOrEmpty(p.Category));
            if (uncategorizedCount > 0)
            {
                categories.Add(new CategoryViewModel
                {
                    Name = "未分類",
                    OriginalName = "未分類",
                    Icon = "📁",
                    Color = "#E9ECEF",
                    ProjectCount = uncategorizedCount,
                    IsDefault = true
                });
            }
        }

        private void UpdateStatistics()
        {
            txtTotalCategories.Text = categories.Count(c => !c.IsDefault).ToString();
            txtTotalProjects.Text = projects.Count.ToString();
            txtUncategorizedProjects.Text = categories.FirstOrDefault(c => c.IsDefault)?.ProjectCount.ToString() ?? "0";
        }

        private void BtnAddCategory_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CategoryEditDialog(null);
            if (dialog.ShowDialog() == true)
            {
                var newCategory = new CategoryViewModel
                {
                    Name = dialog.CategoryName,
                    OriginalName = dialog.CategoryName,
                    Icon = dialog.CategoryIcon,
                    Color = dialog.CategoryColor,
                    ProjectCount = 0,
                    IsDefault = false
                };

                categories.Insert(categories.Count(c => !c.IsDefault), newCategory);
                hasChanges = true;
                UpdateStatistics();
            }
        }

        private void BtnEditCategory_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Tag is CategoryViewModel category && !category.IsDefault)
            {
                var dialog = new CategoryEditDialog(category.Name, category.Icon, category.Color);
                if (dialog.ShowDialog() == true)
                {
                    category.Name = dialog.CategoryName;
                    category.Icon = dialog.CategoryIcon;
                    category.Color = dialog.CategoryColor;
                    category.UpdateColorBrush();
                    hasChanges = true;
                }
            }
        }

        private void BtnDeleteCategory_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Tag is CategoryViewModel category && !category.IsDefault)
            {
                var result = System.Windows.MessageBox.Show(
                    $"カテゴリ「{category.Name}」を削除しますか？\n\nこのカテゴリに属する{category.ProjectCount}個のプロジェクトは「未分類」になります。",
                    "カテゴリ削除の確認",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    // カテゴリを削除し、該当プロジェクトを未分類に移動
                    foreach (var project in projects.Where(p => p.Category == category.OriginalName))
                    {
                        project.Category = "";
                        project.CategoryIcon = "📁";
                        project.CategoryColor = "#E9ECEF";
                    }

                    categories.Remove(category);
                    hasChanges = true;
                    LoadCategories(); // 統計を再計算するために再読み込み
                    UpdateStatistics();
                }
            }
        }

        private void BtnChangeColor_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Tag is CategoryViewModel category)
            {
                var colorDialog = new System.Windows.Forms.ColorDialog();
                
                try
                {
                    var currentColor = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(category.Color);
                    colorDialog.Color = System.Drawing.Color.FromArgb(currentColor.A, currentColor.R, currentColor.G, currentColor.B);
                }
                catch { }

                if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var newColor = colorDialog.Color;
                    category.Color = $"#{newColor.R:X2}{newColor.G:X2}{newColor.B:X2}";
                    category.UpdateColorBrush();
                    hasChanges = true;
                }
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (hasChanges)
            {
                // カテゴリ変更をプロジェクトに適用
                foreach (var category in categories.Where(c => !c.IsDefault))
                {
                    // 元のカテゴリ名でプロジェクトを検索
                    var projectsInCategory = projects.Where(p => p.Category == category.OriginalName);
                    foreach (var project in projectsInCategory)
                    {
                        // カテゴリ名、アイコン、色をすべて更新
                        project.Category = category.Name;
                        project.CategoryIcon = category.Icon;
                        project.CategoryColor = category.Color;
                    }
                }

                ProjectManager.SaveProjects(projects);
                System.Windows.MessageBox.Show("カテゴリの変更を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            DialogResult = true;
            Close();
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            if (hasChanges)
            {
                var result = System.Windows.MessageBox.Show("変更をリセットしますか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    LoadCategories();
                    UpdateStatistics();
                    hasChanges = false;
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            if (hasChanges)
            {
                var result = System.Windows.MessageBox.Show("変更が保存されていません。保存しますか？", 
                    "未保存の変更", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    BtnSave_Click(sender, e);
                    return;
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    return;
                }
            }

            DialogResult = false;
            Close();
        }
    }

    public class CategoryViewModel : INotifyPropertyChanged
    {
        private string _name = "";
        private string _icon = "📁";
        private string _color = "#E9ECEF";
        private SolidColorBrush _colorBrush;

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public string OriginalName { get; set; } = "";

        public string Icon
        {
            get => _icon;
            set
            {
                _icon = value;
                OnPropertyChanged(nameof(Icon));
            }
        }

        public string Color
        {
            get => _color;
            set
            {
                _color = value;
                UpdateColorBrush();
                OnPropertyChanged(nameof(Color));
            }
        }

        public SolidColorBrush ColorBrush
        {
            get => _colorBrush;
            private set
            {
                _colorBrush = value;
                OnPropertyChanged(nameof(ColorBrush));
            }
        }

        public int ProjectCount { get; set; }
        public string ProjectCountText => $"{ProjectCount}個";
        public bool IsDefault { get; set; }
        public Visibility DeleteButtonVisibility => IsDefault ? Visibility.Collapsed : Visibility.Visible;

        public CategoryViewModel()
        {
            UpdateColorBrush();
        }

        public void UpdateColorBrush()
        {
            try
            {
                var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(_color);
                ColorBrush = new SolidColorBrush(color);
            }
            catch
            {
                ColorBrush = new SolidColorBrush(Colors.LightGray);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}