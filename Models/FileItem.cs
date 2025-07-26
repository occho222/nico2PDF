using System;
using System.ComponentModel;
using System.IO;

namespace Nico2PDF.Models
{
    /// <summary>
    /// �t�@�C���A�C�e�����f��
    /// </summary>
    public class FileItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        private string _targetPages = "";
        private int _number;
        private string _relativePath = "";
        private string _displayName = "";
        private string _originalFileName = "";
        private string _pdfStatus = "";

        /// <summary>
        /// �I�����
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        /// <summary>
        /// �Ώۃy�[�W
        /// </summary>
        public string TargetPages
        {
            get => _targetPages;
            set
            {
                _targetPages = value;
                OnPropertyChanged(nameof(TargetPages));
            }
        }

        /// <summary>
        /// �ԍ�
        /// </summary>
        public int Number
        {
            get => _number;
            set
            {
                _number = value;
                OnPropertyChanged(nameof(Number));
            }
        }

        /// <summary>
        /// �t�@�C����
        /// </summary>
        public string FileName { get; set; } = "";

        /// <summary>
        /// �\�����i���[�U�[���ύX�\�j
        /// </summary>
        public string DisplayName
        {
            get => string.IsNullOrEmpty(_displayName) ? FileName : _displayName;
            set
            {
                _displayName = value;
                OnPropertyChanged(nameof(DisplayName));
            }
        }

        /// <summary>
        /// ���̃t�@�C�����i�ύX�O�̖��O��ێ��j
        /// </summary>
        public string OriginalFileName
        {
            get => string.IsNullOrEmpty(_originalFileName) ? FileName : _originalFileName;
            set
            {
                _originalFileName = value;
                OnPropertyChanged(nameof(OriginalFileName));
            }
        }

        /// <summary>
        /// �t�@�C�������ύX����Ă��邩�ǂ���
        /// </summary>
        public bool IsRenamed => !string.IsNullOrEmpty(_displayName) && _displayName != FileName;

        /// <summary>
        /// �t�@�C���p�X
        /// </summary>
        public string FilePath { get; set; } = "";

        /// <summary>
        /// �g���q
        /// </summary>
        public string Extension { get; set; } = "";

        /// <summary>
        /// �ŏI�X�V����
        /// </summary>
        public DateTime LastModified { get; set; }

        /// <summary>
        /// PDF�X�e�[�^�X
        /// </summary>
        public string PdfStatus
        {
            get => _pdfStatus;
            set
            {
                if (_pdfStatus != value)
                {
                    _pdfStatus = value;
                    OnPropertyChanged(nameof(PdfStatus));
                }
            }
        }

        /// <summary>
        /// �\������
        /// </summary>
        public int DisplayOrder { get; set; } = 0;

        /// <summary>
        /// ���΃p�X�i�T�u�t�H���_�ǂݍ��ݗp�j
        /// </summary>
        public string RelativePath
        {
            get => _relativePath;
            set
            {
                _relativePath = value;
                OnPropertyChanged(nameof(RelativePath));
                OnPropertyChanged(nameof(FolderName));
            }
        }

        /// <summary>
        /// �t�H���_���̂݁i���΃p�X����t�H���_�������݂̂𒊏o�j
        /// </summary>
        public string FolderName
        {
            get
            {
                if (string.IsNullOrEmpty(RelativePath))
                {
                    return "";
                }

                // RelativePath���t�@�C���p�X�̏ꍇ�A�f�B���N�g�������݂̂��擾
                var directoryPath = Path.GetDirectoryName(RelativePath);
                
                if (string.IsNullOrEmpty(directoryPath))
                {
                    return "";
                }

                // �Ō�̃t�H���_���݂̂�Ԃ�
                return Path.GetFileName(directoryPath);
            }
        }

        /// <summary>
        /// �\���������Z�b�g�i���̃t�@�C�����ɖ߂��j
        /// </summary>
        public void ResetDisplayName()
        {
            DisplayName = "";
            OnPropertyChanged(nameof(DisplayName));
            OnPropertyChanged(nameof(IsRenamed));
        }

        /// <summary>
        /// �v���p�e�B�ύX�C�x���g
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// �v���p�e�B�ύX�ʒm
        /// </summary>
        /// <param name="propertyName">�v���p�e�B��</param>
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}