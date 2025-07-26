using System;
using System.ComponentModel;
using System.IO;
using System.Windows;

namespace Nico2PDF.Views
{
    /// <summary>
    /// ファイル名変更ダイアログ
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
            
            // 日本語テキストを設定
            Title = "ファイル名変更";
            txtTitle.Text = "ファイル名変更";
            btnHelp.Content = "ヘルプ";
            btnHelp.ToolTip = "ファイル名変更方法の詳細説明を表示";
            
            lblCurrentFileName.Text = "現在のファイル名:";
            lblNewFileName.Text = "新しいファイル名:";
            lblPreview.Text = "変更後のファイル名:";
            
            txtWarningTitle.Text = "注意事項";
            txtWarning1.Text = "・ファイル名を変更すると、元のファイル名に戻すことはできません";
            txtWarning2.Text = "・他のアプリケーションで使用中のファイルは変更できません";
            txtWarning3.Text = "・変更後はPDFファイルの再生成が必要になる場合があります";
            
            txtErrorTitle.Text = "入力エラー";
            btnOK.Content = "変更";
            btnCancel.Content = "キャンセル";
        }

        /// <summary>
        /// 現在のファイル名
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
        /// 新しいファイル名
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
        /// プレビューファイル名
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
        /// 検証エラー
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
        /// 検証エラーがあるかどうか
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
        /// 入力値が有効かどうか
        /// </summary>
        public bool IsValid => !HasValidationError && !string.IsNullOrWhiteSpace(NewFileName);

        /// <summary>
        /// 元のファイルパス
        /// </summary>
        public string OriginalFilePath { get; set; } = "";

        /// <summary>
        /// ファイル名を検証
        /// </summary>
        private void ValidateFileName()
        {
            if (string.IsNullOrWhiteSpace(_newFileName))
            {
                ValidationError = "ファイル名を入力してください。";
                HasValidationError = true;
                return;
            }

            // 無効な文字をチェック
            var invalidChars = Path.GetInvalidFileNameChars();
            foreach (var invalidChar in invalidChars)
            {
                if (_newFileName.Contains(invalidChar))
                {
                    ValidationError = $"ファイル名に無効な文字 '{invalidChar}' が含まれています。";
                    HasValidationError = true;
                    return;
                }
            }

            // 予約語をチェック
            var reservedNames = new[] { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
            if (Array.Exists(reservedNames, name => string.Equals(name, _newFileName, StringComparison.OrdinalIgnoreCase)))
            {
                ValidationError = $"'{_newFileName}' は予約語のため使用できません。";
                HasValidationError = true;
                return;
            }

            // 既存ファイルの存在チェック
            if (!string.IsNullOrEmpty(OriginalFilePath))
            {
                var directory = Path.GetDirectoryName(OriginalFilePath);
                var extension = Path.GetExtension(OriginalFilePath);
                var newFilePath = Path.Combine(directory!, _newFileName + extension);
                
                if (File.Exists(newFilePath) && !string.Equals(newFilePath, OriginalFilePath, StringComparison.OrdinalIgnoreCase))
                {
                    ValidationError = "同じ名前のファイルが既に存在します。";
                    HasValidationError = true;
                    return;
                }
            }

            // 文字数制限チェック
            if (_newFileName.Length > 200)
            {
                ValidationError = "ファイル名が長すぎます。200文字以内で入力してください。";
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
            var helpMessage = @"?? ファイル名変更機能の使い方

【基本的な使い方】
1. 「新しいファイル名」欄に希望するファイル名を入力してください
2. プレビューで結果を確認してください
3. 「変更」ボタンで変更を実行してください

【ファイル名変更について】
・Windowsエクスプローラーでのファイル名も変更されます
・変更後は元に戻すことができません
・ファイルが他のアプリケーションで使用中の場合は変更できません
・PDFファイルの再生成が必要になる場合があります

【ファイル名の制限】
・使用できない文字: \ / : * ? "" < > |
・予約語は使用できません（CON, PRN, AUX等）
・既存のファイルと同じ名前は使用できません
・200文字以内で入力してください

【注意事項】
・拡張子は自動的に付加されるため、入力不要です
・変更前にファイルのバックアップを推奨します
・重要なファイルは慎重に操作してください";

            System.Windows.MessageBox.Show(helpMessage, "ファイル名変更機能のヘルプ", 
                MessageBoxButton.OK, MessageBoxImage.Information);
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