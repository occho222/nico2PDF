using System.ComponentModel;

namespace Nico2PDF.Models
{
    /// <summary>
    /// 一括変更用ファイルアイテム
    /// </summary>
    public class BatchRenameItem : INotifyPropertyChanged
    {
        private string _newFileName = "";
        private bool _hasError = false;
        private string _errorMessage = "";

        /// <summary>
        /// 元のファイルアイテム
        /// </summary>
        public FileItem OriginalItem { get; set; } = new FileItem();

        /// <summary>
        /// 現在のファイル名（拡張子なし）
        /// </summary>
        public string CurrentFileName => System.IO.Path.GetFileNameWithoutExtension(OriginalItem.FileName);

        /// <summary>
        /// 新しいファイル名（拡張子なし）
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
        /// プレビューファイル名（拡張子あり）
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
        /// ファイル名が変更されているかどうか
        /// </summary>
        public bool IsChanged => !string.Equals(CurrentFileName, NewFileName, System.StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// エラーがあるかどうか
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
        /// エラーメッセージ
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
        /// ファイル名を検証
        /// </summary>
        private void ValidateFileName()
        {
            if (string.IsNullOrWhiteSpace(NewFileName))
            {
                ErrorMessage = "ファイル名を入力してください。";
                HasError = true;
                return;
            }

            // 無効な文字をチェック
            var invalidChars = System.IO.Path.GetInvalidFileNameChars();
            foreach (var invalidChar in invalidChars)
            {
                if (NewFileName.Contains(invalidChar))
                {
                    ErrorMessage = $"無効な文字 '{invalidChar}' が含まれています。";
                    HasError = true;
                    return;
                }
            }

            // 予約語をチェック
            var reservedNames = new[] { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
            if (System.Array.Exists(reservedNames, name => string.Equals(name, NewFileName, System.StringComparison.OrdinalIgnoreCase)))
            {
                ErrorMessage = $"'{NewFileName}' は予約語のため使用できません。";
                HasError = true;
                return;
            }

            // 文字数制限チェック
            if (NewFileName.Length > 200)
            {
                ErrorMessage = "ファイル名が長すぎます。200文字以内で入力してください。";
                HasError = true;
                return;
            }

            ErrorMessage = "";
            HasError = false;
        }

        /// <summary>
        /// プロパティ変更イベント
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// プロパティ変更通知
        /// </summary>
        /// <param name="propertyName">プロパティ名</param>
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}