using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace Nico2PDF.Models
{
    /// <summary>
    /// プロジェクトカテゴリグループ
    /// </summary>
    public class ProjectCategoryGroup : INotifyPropertyChanged
    {
        private bool _isExpanded = true;
        private string _categoryName = "";

        /// <summary>
        /// カテゴリ名
        /// </summary>
        public string CategoryName
        {
            get => _categoryName;
            set
            {
                _categoryName = value;
                OnPropertyChanged(nameof(CategoryName));
            }
        }

        /// <summary>
        /// カテゴリアイコン
        /// </summary>
        public string CategoryIcon { get; set; } = "??";

        /// <summary>
        /// カテゴリの色
        /// </summary>
        public string CategoryColor { get; set; } = "#E9ECEF";

        /// <summary>
        /// 展開状態
        /// </summary>
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                _isExpanded = value;
                OnPropertyChanged(nameof(IsExpanded));
            }
        }

        /// <summary>
        /// このカテゴリのプロジェクト
        /// </summary>
        public ObservableCollection<ProjectData> Projects { get; set; } = new ObservableCollection<ProjectData>();

        /// <summary>
        /// プロジェクト数
        /// </summary>
        public int ProjectCount => Projects.Count;

        /// <summary>
        /// 表示名
        /// </summary>
        public string DisplayName => $"{CategoryIcon} {CategoryName} ({ProjectCount})";

        /// <summary>
        /// プロパティ変更イベント
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// プロパティ変更通知
        /// </summary>
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}