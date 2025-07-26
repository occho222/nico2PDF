using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace Nico2PDF.Models
{
    /// <summary>
    /// �v���W�F�N�g�J�e�S���O���[�v
    /// </summary>
    public class ProjectCategoryGroup : INotifyPropertyChanged
    {
        private bool _isExpanded = true;
        private string _categoryName = "";

        /// <summary>
        /// �J�e�S����
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
        /// �J�e�S���A�C�R��
        /// </summary>
        public string CategoryIcon { get; set; } = "??";

        /// <summary>
        /// �J�e�S���̐F
        /// </summary>
        public string CategoryColor { get; set; } = "#E9ECEF";

        /// <summary>
        /// �W�J���
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
        /// ���̃J�e�S���̃v���W�F�N�g
        /// </summary>
        public ObservableCollection<ProjectData> Projects { get; set; } = new ObservableCollection<ProjectData>();

        /// <summary>
        /// �v���W�F�N�g��
        /// </summary>
        public int ProjectCount => Projects.Count;

        /// <summary>
        /// �\����
        /// </summary>
        public string DisplayName => $"{CategoryIcon} {CategoryName} ({ProjectCount})";

        /// <summary>
        /// �v���p�e�B�ύX�C�x���g
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// �v���p�e�B�ύX�ʒm
        /// </summary>
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}