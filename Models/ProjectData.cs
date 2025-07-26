using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.Json.Serialization;

namespace Nico2PDF.Models
{
    /// <summary>
    /// プロジェクトデータモデル
    /// </summary>
    public class ProjectData : INotifyPropertyChanged
    {
        private string _name = "";
        private bool _isActive = false;
        private bool _includeSubfolders = false;
        private bool _useCustomPdfPath = false;
        private string _customPdfPath = "";

        /// <summary>
        /// プロジェクトID
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// プロジェクト名
        /// </summary>
        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        /// <summary>
        /// プロジェクトの階層レベル
        /// </summary>
        public int Level { get; set; } = 0;

        /// <summary>
        /// アクティブ状態
        /// </summary>
        public bool IsActive
        {
            get => _isActive;
            set
            {
                _isActive = value;
                OnPropertyChanged(nameof(IsActive));
            }
        }

        /// <summary>
        /// プロジェクトフォルダのパス
        /// </summary>
        public string FolderPath { get; set; } = "";

        /// <summary>
        /// サブフォルダを含むかどうか
        /// </summary>
        public bool IncludeSubfolders
        {
            get => _includeSubfolders;
            set
            {
                _includeSubfolders = value;
                OnPropertyChanged(nameof(IncludeSubfolders));
            }
        }

        /// <summary>
        /// カスタムPDF保存パスを使用するかどうか
        /// </summary>
        public bool UseCustomPdfPath
        {
            get => _useCustomPdfPath;
            set
            {
                _useCustomPdfPath = value;
                OnPropertyChanged(nameof(UseCustomPdfPath));
            }
        }

        /// <summary>
        /// カスタムPDF保存パス
        /// </summary>
        public string CustomPdfPath
        {
            get => _customPdfPath;
            set
            {
                _customPdfPath = value;
                OnPropertyChanged(nameof(CustomPdfPath));
            }
        }

        /// <summary>
        /// PDF出力フォルダのパス
        /// </summary>
        public string PdfOutputFolder 
        { 
            get 
            {
                if (UseCustomPdfPath && !string.IsNullOrEmpty(CustomPdfPath))
                {
                    // カスタムパスを使用する場合、PDFフォルダを作成
                    return Path.Combine(CustomPdfPath, "PDF");
                }
                return Path.Combine(FolderPath, "PDF");
            }
            set
            {
                // 後方互換性のため、setter は保持
                if (!UseCustomPdfPath)
                {
                    // 従来通りの動作
                }
            }
        }

        /// <summary>
        /// 結合PDF保存フォルダのパス
        /// </summary>
        [JsonIgnore]
        public string MergePdfFolder
        {
            get
            {
                if (UseCustomPdfPath && !string.IsNullOrEmpty(CustomPdfPath))
                {
                    // カスタムパスを使用する場合、カスタムパス下にmergePDFフォルダを作成
                    return Path.Combine(CustomPdfPath, "mergePDF");
                }
                // 通常はプロジェクトフォルダ下にmergePDFフォルダを作成
                return Path.Combine(FolderPath, "mergePDF");
            }
        }

        /// <summary>
        /// 結合PDFファイル名
        /// </summary>
        public string MergeFileName { get; set; } = "結合PDF";

        /// <summary>
        /// ページ番号追加フラグ
        /// </summary>
        public bool AddPageNumber { get; set; } = false;

        /// <summary>
        /// しおり追加フラグ
        /// </summary>
        public bool AddBookmarks { get; set; } = true;

        /// <summary>
        /// フォルダ別グループ化フラグ
        /// </summary>
        public bool GroupByFolder { get; set; } = false;

        /// <summary>
        /// ヘッダ・フッタ追加フラグ
        /// </summary>
        public bool AddHeaderFooter { get; set; } = false;

        /// <summary>
        /// ヘッダ・フッタテキスト
        /// </summary>
        public string HeaderFooterText { get; set; } = "";

        /// <summary>
        /// ヘッダ・フッタフォントサイズ
        /// </summary>
        public float HeaderFooterFontSize { get; set; } = 10.0f;

        /// <summary>
        /// 最新の結合PDFファイルパス
        /// </summary>
        public string LatestMergedPdfPath { get; set; } = "";

        /// <summary>
        /// 作成日時
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;

        /// <summary>
        /// 最終アクセス日時
        /// </summary>
        public DateTime LastAccessDate { get; set; } = DateTime.Now;

        /// <summary>
        /// ファイルアイテムリスト
        /// </summary>
        public List<FileItemData> FileItems { get; set; } = new List<FileItemData>();

        /// <summary>
        /// プロジェクトカテゴリ（フォルダ分け用）
        /// </summary>
        public string Category { get; set; } = "";

        /// <summary>
        /// カテゴリの説明
        /// </summary>
        public string CategoryDescription { get; set; } = "";

        /// <summary>
        /// カテゴリの色（表示用）
        /// </summary>
        public string CategoryColor { get; set; } = "#E9ECEF";

        /// <summary>
        /// カテゴリアイコン（表示用）
        /// </summary>
        public string CategoryIcon { get; set; } = "??";

        /// <summary>
        /// 表示名（JSON非対象）
        /// </summary>
        [JsonIgnore]
        public string DisplayName => string.IsNullOrEmpty(Category) ? Name : $"{Name}";

        /// <summary>
        /// カテゴリ付き表示名（JSON非対象）
        /// </summary>
        [JsonIgnore]
        public string CategoryDisplayName => string.IsNullOrEmpty(Category) ? $"?? {Name}" : $"{CategoryIcon} {Name}";

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