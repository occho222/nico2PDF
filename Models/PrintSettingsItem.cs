using System.ComponentModel;

namespace Nico2PDF.Models
{
    /// <summary>
    /// Excel印刷設定項目
    /// </summary>
    public class PrintSettingsItem : INotifyPropertyChanged
    {
        private PaperSize? _paperSize = null;
        private Orientation? _orientation = null;
        private FitToPageOption? _fitToPageOption = null;

        /// <summary>
        /// 対象ファイルアイテム
        /// </summary>
        public FileItem FileItem { get; set; } = new FileItem();

        /// <summary>
        /// 用紙サイズ
        /// </summary>
        public PaperSize? PaperSize
        {
            get => _paperSize;
            set
            {
                _paperSize = value;
                OnPropertyChanged(nameof(PaperSize));
                OnPropertyChanged(nameof(PaperSizeDisplay));
                OnPropertyChanged(nameof(SettingsSummary));
                OnPropertyChanged(nameof(IsPaperSizeChanged));
                OnPropertyChanged(nameof(HasChanges));
            }
        }

        /// <summary>
        /// 用紙の向き
        /// </summary>
        public Orientation? Orientation
        {
            get => _orientation;
            set
            {
                _orientation = value;
                OnPropertyChanged(nameof(Orientation));
                OnPropertyChanged(nameof(OrientationDisplay));
                OnPropertyChanged(nameof(SettingsSummary));
                OnPropertyChanged(nameof(IsOrientationChanged));
                OnPropertyChanged(nameof(HasChanges));
            }
        }

        /// <summary>
        /// ページに合わせる設定
        /// </summary>
        public FitToPageOption? FitToPageOption
        {
            get => _fitToPageOption;
            set
            {
                _fitToPageOption = value;
                OnPropertyChanged(nameof(FitToPageOption));
                OnPropertyChanged(nameof(FitToPageOptionDisplay));
                OnPropertyChanged(nameof(SettingsSummary));
                OnPropertyChanged(nameof(IsFitToPageOptionChanged));
                OnPropertyChanged(nameof(HasChanges));
            }
        }

        /// <summary>
        /// 用紙サイズが変更されているか
        /// </summary>
        public bool IsPaperSizeChanged => _paperSize.HasValue;

        /// <summary>
        /// 用紙の向きが変更されているか
        /// </summary>
        public bool IsOrientationChanged => _orientation.HasValue;

        /// <summary>
        /// ページ設定が変更されているか
        /// </summary>
        public bool IsFitToPageOptionChanged => _fitToPageOption.HasValue;

        /// <summary>
        /// 何らかの設定が変更されているか
        /// </summary>
        public bool HasChanges => IsPaperSizeChanged || IsOrientationChanged || IsFitToPageOptionChanged;

        /// <summary>
        /// 設定の概要文字列（変更されたもののみ表示）
        /// </summary>
        public string SettingsSummary
        {
            get
            {
                var changes = new System.Collections.Generic.List<string>();

                if (IsPaperSizeChanged && _paperSize.HasValue)
                {
                    var sizeText = _paperSize == Models.PaperSize.A4 ? "A4" : "A3";
                    changes.Add($"用紙: {sizeText}");
                }

                if (IsOrientationChanged && _orientation.HasValue)
                {
                    var orientationText = _orientation == Models.Orientation.Portrait ? "縦" : "横";
                    changes.Add($"向き: {orientationText}");
                }

                if (IsFitToPageOptionChanged && _fitToPageOption.HasValue)
                {
                    var fitText = _fitToPageOption switch
                    {
                        Models.FitToPageOption.FitSheetOnOnePage => "シートを1ページに印刷",
                        Models.FitToPageOption.FitAllColumnsOnOnePage => "全ての列を1ページに印刷",
                        Models.FitToPageOption.FitAllRowsOnOnePage => "全ての行を1ページに印刷",
                        _ => "標準"
                    };
                    changes.Add($"印刷: {fitText}");
                }

                return changes.Count > 0 ? string.Join(", ", changes) : "変更なし";
            }
        }

        /// <summary>
        /// 表示用の用紙サイズ文字列
        /// </summary>
        public string PaperSizeDisplay
        {
            get => _paperSize switch
            {
                Models.PaperSize.A4 => "A4",
                Models.PaperSize.A3 => "A3",
                _ => ""
            };
        }

        /// <summary>
        /// 表示用の向き文字列
        /// </summary>
        public string OrientationDisplay
        {
            get => _orientation switch
            {
                Models.Orientation.Portrait => "縦",
                Models.Orientation.Landscape => "横",
                _ => ""
            };
        }

        /// <summary>
        /// 表示用の印刷範囲文字列
        /// </summary>
        public string FitToPageOptionDisplay
        {
            get => _fitToPageOption switch
            {
                Models.FitToPageOption.None => "標準",
                Models.FitToPageOption.FitSheetOnOnePage => "シートを1ページに印刷",
                Models.FitToPageOption.FitAllColumnsOnOnePage => "全ての列を1ページに印刷",
                Models.FitToPageOption.FitAllRowsOnOnePage => "全ての行を1ページに印刷",
                _ => ""
            };
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

    /// <summary>
    /// 用紙サイズ
    /// </summary>
    public enum PaperSize
    {
        A4,
        A3
    }

    /// <summary>
    /// 用紙の向き
    /// </summary>
    public enum Orientation
    {
        Portrait,   // 縦
        Landscape   // 横
    }

    /// <summary>
    /// ページに合わせる設定
    /// </summary>
    public enum FitToPageOption
    {
        None,                        // 標準
        FitSheetOnOnePage,          // シートを1ページに印刷
        FitAllColumnsOnOnePage,     // 全ての列を1ページに印刷
        FitAllRowsOnOnePage         // 全ての行を1ページに印刷
    }
}