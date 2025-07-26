using System.ComponentModel;

namespace Nico2PDF.Models
{
    /// <summary>
    /// �ꊇ�ύX�p�t�@�C���A�C�e��
    /// </summary>
    public class BatchRenameItem : INotifyPropertyChanged
    {
        private string _newFileName = "";
        private bool _hasError = false;
        private string _errorMessage = "";

        /// <summary>
        /// ���̃t�@�C���A�C�e��
        /// </summary>
        public FileItem OriginalItem { get; set; } = new FileItem();

        /// <summary>
        /// ���݂̃t�@�C�����i�g���q�Ȃ��j
        /// </summary>
        public string CurrentFileName => System.IO.Path.GetFileNameWithoutExtension(OriginalItem.FileName);

        /// <summary>
        /// �V�����t�@�C�����i�g���q�Ȃ��j
        /// </summary>
        public string NewFileName
        {
            get => string.IsNullOrEmpty(_newFileName) ? CurrentFileName : _newFileName;
            set
            {
                _newFileName = value;
                OnPropertyChanged(nameof(NewFileName));
                OnPropertyChanged(nameof(PreviewFileName));
                OnPropertyChanged(nameof(IsChanged));
                ValidateFileName();
            }
        }

        /// <summary>
        /// �v���r���[�t�@�C�����i�g���q����j
        /// </summary>
        public string PreviewFileName
        {
            get
            {
                var extension = System.IO.Path.GetExtension(OriginalItem.FileName);
                return NewFileName + extension;
            }
        }

        /// <summary>
        /// �t�@�C�������ύX����Ă��邩�ǂ���
        /// </summary>
        public bool IsChanged => !string.Equals(CurrentFileName, NewFileName, System.StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// �G���[�����邩�ǂ���
        /// </summary>
        public bool HasError
        {
            get => _hasError;
            set
            {
                _hasError = value;
                OnPropertyChanged(nameof(HasError));
            }
        }

        /// <summary>
        /// �G���[���b�Z�[�W
        /// </summary>
        public string ErrorMessage
        {
            get => _errorMessage;
            set
            {
                _errorMessage = value;
                OnPropertyChanged(nameof(ErrorMessage));
            }
        }

        /// <summary>
        /// �t�@�C����������
        /// </summary>
        private void ValidateFileName()
        {
            if (string.IsNullOrWhiteSpace(NewFileName))
            {
                ErrorMessage = "�t�@�C��������͂��Ă��������B";
                HasError = true;
                return;
            }

            // �����ȕ������`�F�b�N
            var invalidChars = System.IO.Path.GetInvalidFileNameChars();
            foreach (var invalidChar in invalidChars)
            {
                if (NewFileName.Contains(invalidChar))
                {
                    ErrorMessage = $"�����ȕ��� '{invalidChar}' ���܂܂�Ă��܂��B";
                    HasError = true;
                    return;
                }
            }

            // �\�����`�F�b�N
            var reservedNames = new[] { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
            if (System.Array.Exists(reservedNames, name => string.Equals(name, NewFileName, System.StringComparison.OrdinalIgnoreCase)))
            {
                ErrorMessage = $"'{NewFileName}' �͗\���̂��ߎg�p�ł��܂���B";
                HasError = true;
                return;
            }

            // �����������`�F�b�N
            if (NewFileName.Length > 200)
            {
                ErrorMessage = "�t�@�C�������������܂��B200�����ȓ��œ��͂��Ă��������B";
                HasError = true;
                return;
            }

            ErrorMessage = "";
            HasError = false;
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