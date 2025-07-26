using System;
using System.ComponentModel;
using System.IO;
using System.Windows;

namespace Nico2PDF.Views
{
    /// <summary>
    /// �t�@�C�����ύX�_�C�A���O
    /// </summary>
    public partial class FileRenameDialog : Window, INotifyPropertyChanged
    {
        private string _currentFileName = "";
        private string _newFileName = "";
        private string _validationError = "";
        private bool _hasValidationError = false;

        public FileRenameDialog()
        {
            InitializeComponent();
            DataContext = this;
            
            // ���{��e�L�X�g��ݒ�
            Title = "�t�@�C�����ύX";
            txtTitle.Text = "�t�@�C�����ύX";
            btnHelp.Content = "�w���v";
            btnHelp.ToolTip = "�t�@�C�����ύX���@�̏ڍא�����\��";
            
            lblCurrentFileName.Text = "���݂̃t�@�C����:";
            lblNewFileName.Text = "�V�����t�@�C����:";
            lblPreview.Text = "�ύX��̃t�@�C����:";
            
            txtWarningTitle.Text = "���ӎ���";
            txtWarning1.Text = "�E�t�@�C������ύX����ƁA���̃t�@�C�����ɖ߂����Ƃ͂ł��܂���";
            txtWarning2.Text = "�E���̃A�v���P�[�V�����Ŏg�p���̃t�@�C���͕ύX�ł��܂���";
            txtWarning3.Text = "�E�ύX���PDF�t�@�C���̍Đ������K�v�ɂȂ�ꍇ������܂�";
            
            txtErrorTitle.Text = "���̓G���[";
            btnOK.Content = "�ύX";
            btnCancel.Content = "�L�����Z��";
        }

        /// <summary>
        /// ���݂̃t�@�C����
        /// </summary>
        public string CurrentFileName
        {
            get => _currentFileName;
            set
            {
                _currentFileName = value;
                OnPropertyChanged(nameof(CurrentFileName));
                if (string.IsNullOrEmpty(_newFileName))
                {
                    NewFileName = Path.GetFileNameWithoutExtension(value);
                }
            }
        }

        /// <summary>
        /// �V�����t�@�C����
        /// </summary>
        public string NewFileName
        {
            get => _newFileName;
            set
            {
                _newFileName = value;
                OnPropertyChanged(nameof(NewFileName));
                OnPropertyChanged(nameof(PreviewFileName));
                ValidateFileName();
            }
        }

        /// <summary>
        /// �v���r���[�t�@�C����
        /// </summary>
        public string PreviewFileName
        {
            get
            {
                if (string.IsNullOrEmpty(_newFileName))
                    return "";

                var extension = Path.GetExtension(_currentFileName);
                return _newFileName + extension;
            }
        }

        /// <summary>
        /// ���؃G���[
        /// </summary>
        public string ValidationError
        {
            get => _validationError;
            set
            {
                _validationError = value;
                OnPropertyChanged(nameof(ValidationError));
            }
        }

        /// <summary>
        /// ���؃G���[�����邩�ǂ���
        /// </summary>
        public bool HasValidationError
        {
            get => _hasValidationError;
            set
            {
                _hasValidationError = value;
                OnPropertyChanged(nameof(HasValidationError));
                OnPropertyChanged(nameof(IsValid));
            }
        }

        /// <summary>
        /// ���͒l���L�����ǂ���
        /// </summary>
        public bool IsValid => !HasValidationError && !string.IsNullOrWhiteSpace(NewFileName);

        /// <summary>
        /// ���̃t�@�C���p�X
        /// </summary>
        public string OriginalFilePath { get; set; } = "";

        /// <summary>
        /// �t�@�C����������
        /// </summary>
        private void ValidateFileName()
        {
            if (string.IsNullOrWhiteSpace(_newFileName))
            {
                ValidationError = "�t�@�C��������͂��Ă��������B";
                HasValidationError = true;
                return;
            }

            // �����ȕ������`�F�b�N
            var invalidChars = Path.GetInvalidFileNameChars();
            foreach (var invalidChar in invalidChars)
            {
                if (_newFileName.Contains(invalidChar))
                {
                    ValidationError = $"�t�@�C�����ɖ����ȕ��� '{invalidChar}' ���܂܂�Ă��܂��B";
                    HasValidationError = true;
                    return;
                }
            }

            // �\�����`�F�b�N
            var reservedNames = new[] { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
            if (Array.Exists(reservedNames, name => string.Equals(name, _newFileName, StringComparison.OrdinalIgnoreCase)))
            {
                ValidationError = $"'{_newFileName}' �͗\���̂��ߎg�p�ł��܂���B";
                HasValidationError = true;
                return;
            }

            // �����t�@�C���̑��݃`�F�b�N
            if (!string.IsNullOrEmpty(OriginalFilePath))
            {
                var directory = Path.GetDirectoryName(OriginalFilePath);
                var extension = Path.GetExtension(OriginalFilePath);
                var newFilePath = Path.Combine(directory!, _newFileName + extension);
                
                if (File.Exists(newFilePath) && !string.Equals(newFilePath, OriginalFilePath, StringComparison.OrdinalIgnoreCase))
                {
                    ValidationError = "�������O�̃t�@�C�������ɑ��݂��܂��B";
                    HasValidationError = true;
                    return;
                }
            }

            // �����������`�F�b�N
            if (_newFileName.Length > 200)
            {
                ValidationError = "�t�@�C�������������܂��B200�����ȓ��œ��͂��Ă��������B";
                HasValidationError = true;
                return;
            }

            ValidationError = "";
            HasValidationError = false;
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            if (IsValid)
            {
                DialogResult = true;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            var helpMessage = @"?? �t�@�C�����ύX�@�\�̎g����

�y��{�I�Ȏg�����z
1. �u�V�����t�@�C�����v���Ɋ�]����t�@�C��������͂��Ă�������
2. �v���r���[�Ō��ʂ��m�F���Ă�������
3. �u�ύX�v�{�^���ŕύX�����s���Ă�������

�y�t�@�C�����ύX�ɂ��āz
�EWindows�G�N�X�v���[���[�ł̃t�@�C�������ύX����܂�
�E�ύX��͌��ɖ߂����Ƃ��ł��܂���
�E�t�@�C�������̃A�v���P�[�V�����Ŏg�p���̏ꍇ�͕ύX�ł��܂���
�EPDF�t�@�C���̍Đ������K�v�ɂȂ�ꍇ������܂�

�y�t�@�C�����̐����z
�E�g�p�ł��Ȃ�����: \ / : * ? "" < > |
�E�\���͎g�p�ł��܂���iCON, PRN, AUX���j
�E�����̃t�@�C���Ɠ������O�͎g�p�ł��܂���
�E200�����ȓ��œ��͂��Ă�������

�y���ӎ����z
�E�g���q�͎����I�ɕt������邽�߁A���͕s�v�ł�
�E�ύX�O�Ƀt�@�C���̃o�b�N�A�b�v�𐄏����܂�
�E�d�v�ȃt�@�C���͐T�d�ɑ��삵�Ă�������";

            System.Windows.MessageBox.Show(helpMessage, "�t�@�C�����ύX�@�\�̃w���v", 
                MessageBoxButton.OK, MessageBoxImage.Information);
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