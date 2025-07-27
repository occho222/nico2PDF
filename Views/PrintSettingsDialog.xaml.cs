using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Nico2PDF.Models;
using Nico2PDF.Services;
using MessageBox = System.Windows.MessageBox;
using MessageBoxButton = System.Windows.MessageBoxButton;
using MessageBoxImage = System.Windows.MessageBoxImage;
using MessageBoxResult = System.Windows.MessageBoxResult;

namespace Nico2PDF.Views
{
    /// <summary>
    /// Excel印刷設定ダイアログ
    /// </summary>
    public partial class PrintSettingsDialog : Window, INotifyPropertyChanged
    {
        private ObservableCollection<PrintSettingsItem> _printSettingsItems = new();
        private string _statusMessage = "";

        public PrintSettingsDialog()
        {
            InitializeComponent();
            DataContext = this;
            InitializeControls();
        }

        /// <summary>
        /// 印刷設定アイテムコレクション
        /// </summary>
        public ObservableCollection<PrintSettingsItem> PrintSettingsItems
        {
            get => _printSettingsItems;
            set
            {
                _printSettingsItems = value;
                OnPropertyChanged(nameof(PrintSettingsItems));
                OnPropertyChanged(nameof(TargetFilesInfo));
                OnPropertyChanged(nameof(SelectedFilesInfo));
            }
        }

        /// <summary>
        /// ステータスメッセージ
        /// </summary>
        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                _statusMessage = value;
                OnPropertyChanged(nameof(StatusMessage));
            }
        }

        /// <summary>
        /// 対象ファイル情報
        /// </summary>
        public string TargetFilesInfo
        {
            get
            {
                var excelFiles = PrintSettingsItems?.Where(item => IsExcelFile(item.FileItem.Extension)).Count() ?? 0;
                var totalFiles = PrintSettingsItems?.Count ?? 0;
                return $"Excelファイル: {excelFiles}件 / 全ファイル: {totalFiles}件";
            }
        }

        /// <summary>
        /// 選択ファイル情報
        /// </summary>
        public string SelectedFilesInfo
        {
            get
            {
                var selectedFiles = PrintSettingsItems?.Where(item => item.FileItem.IsSelected && IsExcelFile(item.FileItem.Extension)).Count() ?? 0;
                return $"選択中: {selectedFiles}件";
            }
        }

        /// <summary>
        /// ダイアログ結果
        /// </summary>
        public new bool? DialogResult { get; private set; }

        /// <summary>
        /// ファイルリストを設定
        /// </summary>
        /// <param name="fileItems">ファイルアイテムリスト</param>
        public void SetFileItems(System.Collections.Generic.List<FileItem> fileItems)
        {
            PrintSettingsItems.Clear();
            foreach (var fileItem in fileItems)
            {
                // Excelファイルのみを対象とする
                if (IsExcelFile(fileItem.Extension))
                {
                    PrintSettingsItems.Add(new PrintSettingsItem { FileItem = fileItem });
                }
            }

            // 各アイテムの変更イベントを監視
            foreach (var item in PrintSettingsItems)
            {
                item.FileItem.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(FileItem.IsSelected))
                    {
                        OnPropertyChanged(nameof(SelectedFilesInfo));
                        UpdateSelectAllCheckbox();
                    }
                };
            }

            OnPropertyChanged(nameof(TargetFilesInfo));
            OnPropertyChanged(nameof(SelectedFilesInfo));
            UpdateSelectAllCheckbox();
            StatusMessage = "印刷設定を編集してください。";
        }

        /// <summary>
        /// コントロールの初期化
        /// </summary>
        private void InitializeControls()
        {
            // 初期状態では何も選択しない（ユーザーが明示的に選択した場合のみ適用）
            cbPaperSize.SelectedIndex = -1;
            cbOrientation.SelectedIndex = -1;
            cbFitToPage.SelectedIndex = -1;
        }

        /// <summary>
        /// Excelファイルかどうかを判定
        /// </summary>
        /// <param name="extension">拡張子</param>
        /// <returns>Excelファイルの場合true</returns>
        private static bool IsExcelFile(string extension)
        {
            return extension.ToUpper() is "XLS" or "XLSX" or "XLSM";
        }

        /// <summary>
        /// 印刷設定適用ボタンクリック
        /// </summary>
        private async void BtnApplyPrintSettings_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = PrintSettingsItems.Where(item => item.FileItem.IsSelected && item.HasChanges).ToList();
            if (!selectedItems.Any())
            {
                MessageBox.Show("印刷設定に変更があるファイルを選択してください。", "選択エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show(
                $"選択された{selectedItems.Count}件のExcelファイルに印刷設定を適用しますか？\n\n注意: ファイルが上書き保存されます。",
                "印刷設定適用確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;

            try
            {
                btnApplyPrintSettings.IsEnabled = false;
                StatusMessage = "印刷設定を適用中...";

                var successCount = 0;
                var errorCount = 0;
                var errors = new System.Collections.Generic.List<string>();

                foreach (var item in selectedItems)
                {
                    try
                    {
                        StatusMessage = $"印刷設定適用中... ({successCount + errorCount + 1}/{selectedItems.Count}) {item.FileItem.FileName}";
                        
                        await System.Threading.Tasks.Task.Run(() =>
                            ExcelPrintSettingsService.ApplyPrintSettings(item.FileItem.FilePath, item));
                        
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        errors.Add($"{item.FileItem.FileName}: {ex.Message}");
                    }
                }

                StatusMessage = $"完了: 成功 {successCount}件, エラー {errorCount}件";

                if (errors.Any())
                {
                    var errorMessage = string.Join("\n", errors.Take(10));
                    if (errors.Count > 10)
                        errorMessage += $"\n... 他{errors.Count - 10}件のエラー";

                    MessageBox.Show($"印刷設定の適用が完了しました。\n\n成功: {successCount}件\nエラー: {errorCount}件\n\nエラー詳細:\n{errorMessage}",
                        "適用結果", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show($"印刷設定の適用が完了しました。\n\n成功: {successCount}件",
                        "適用完了", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                DialogResult = true;
            }
            catch (Exception ex)
            {
                StatusMessage = "エラーが発生しました。";
                MessageBox.Show($"印刷設定の適用中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnApplyPrintSettings.IsEnabled = true;
            }
        }

        /// <summary>
        /// キャンセルボタンクリック
        /// </summary>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        /// <summary>
        /// ヘルプボタンクリック
        /// </summary>
        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            var helpMessage = "Excel印刷設定ヘルプ\n\n" +
                            "機能:\n" +
                            "• 選択したExcelファイルの印刷設定を一括で変更できます\n\n" +
                            "設定項目:\n" +
                            "• 用紙サイズ: A4またはA3を選択\n" +
                            "• 用紙の向き: 縦または横を選択\n" +
                            "• 印刷範囲: 標準、シートを1ページに印刷、全ての列を1ページに印刷、全ての行を1ページに印刷から選択\n\n" +
                            "使用方法:\n" +
                            "1. 上部の一括設定で共通の設定を選択\n" +
                            "2. 「設定概要に反映」ボタンで設定を反映\n" +
                            "3. または個別にファイルごとに設定を変更\n" +
                            "4. 「印刷設定適用」ボタンで設定を保存\n\n" +
                            "注意:\n" +
                            "• Excelファイルのみが対象です\n" +
                            "• ファイルが上書き保存されます";

            MessageBox.Show(helpMessage, "ヘルプ", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// リセットボタンクリック
        /// </summary>
        private void BtnResetAll_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("全ての印刷設定をリセットしますか？", "リセット確認", 
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                foreach (var item in PrintSettingsItems)
                {
                    item.PaperSize = null;
                    item.Orientation = null;
                    item.FitToPageOption = null;
                }

                // 一括設定のコンボボックスもリセット
                cbPaperSize.SelectedIndex = -1;
                cbOrientation.SelectedIndex = -1;
                cbFitToPage.SelectedIndex = -1;

                StatusMessage = "設定をリセットしました。";
            }
        }

        /// <summary>
        /// 設定概要に反映ボタンクリック
        /// </summary>
        private void BtnApplyToAll_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = PrintSettingsItems.Where(item => item.FileItem.IsSelected).ToList();
            if (!selectedItems.Any())
            {
                MessageBox.Show("設定を反映するファイルを選択してください。", "選択エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 一括設定からどの項目が変更されたかを判定
            var hasPaperSizeChange = cbPaperSize.SelectedIndex >= 0;
            var hasOrientationChange = cbOrientation.SelectedIndex >= 0;
            var hasFitToPageChange = cbFitToPage.SelectedIndex >= 0;

            if (!hasPaperSizeChange && !hasOrientationChange && !hasFitToPageChange)
            {
                MessageBox.Show("反映する設定を選択してください。", "設定エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 選択されたファイルに変更された設定のみを設定概要に反映
            foreach (var item in selectedItems)
            {
                if (hasPaperSizeChange)
                {
                    item.PaperSize = GetSelectedPaperSize();
                }
                if (hasOrientationChange)
                {
                    item.Orientation = GetSelectedOrientation();
                }
                if (hasFitToPageChange)
                {
                    item.FitToPageOption = GetSelectedFitToPageOption();
                }
            }

            var appliedSettings = new System.Collections.Generic.List<string>();
            if (hasPaperSizeChange) appliedSettings.Add("用紙サイズ");
            if (hasOrientationChange) appliedSettings.Add("用紙の向き");
            if (hasFitToPageChange) appliedSettings.Add("印刷範囲");

            StatusMessage = $"{selectedItems.Count}件のファイルの設定概要に{string.Join("、", appliedSettings)}を反映しました。";
        }

        /// <summary>
        /// 選択された用紙サイズを取得
        /// </summary>
        private PaperSize? GetSelectedPaperSize()
        {
            if (cbPaperSize.SelectedItem is ComboBoxItem item && item.Tag is string tagString)
            {
                return tagString switch
                {
                    "A4" => PaperSize.A4,
                    "A3" => PaperSize.A3,
                    _ => null
                };
            }
            return null;
        }

        /// <summary>
        /// 選択された用紙の向きを取得
        /// </summary>
        private Models.Orientation? GetSelectedOrientation()
        {
            if (cbOrientation.SelectedItem is ComboBoxItem item && item.Tag is string tagString)
            {
                return tagString switch
                {
                    "Portrait" => Models.Orientation.Portrait,
                    "Landscape" => Models.Orientation.Landscape,
                    _ => null
                };
            }
            return null;
        }

        /// <summary>
        /// 選択されたページ設定を取得
        /// </summary>
        private FitToPageOption? GetSelectedFitToPageOption()
        {
            if (cbFitToPage.SelectedItem is ComboBoxItem item && item.Tag is string tagString)
            {
                return tagString switch
                {
                    "None" => FitToPageOption.None,
                    "FitSheetOnOnePage" => FitToPageOption.FitSheetOnOnePage,
                    "FitAllColumnsOnOnePage" => FitToPageOption.FitAllColumnsOnOnePage,
                    "FitAllRowsOnOnePage" => FitToPageOption.FitAllRowsOnOnePage,
                    _ => null
                };
            }
            return null;
        }

        /// <summary>
        /// プロパティ変更イベント
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// 全選択チェックボックスクリック
        /// </summary>
        private void ChkSelectAll_Click(object sender, RoutedEventArgs e)
        {
            var isChecked = chkSelectAll.IsChecked == true;
            foreach (var item in PrintSettingsItems)
            {
                item.FileItem.IsSelected = isChecked;
            }
        }

        /// <summary>
        /// 全選択チェックボックスの状態を更新
        /// </summary>
        private void UpdateSelectAllCheckbox()
        {
            if (!PrintSettingsItems.Any())
            {
                chkSelectAll.IsChecked = false;
                return;
            }

            var selectedCount = PrintSettingsItems.Count(item => item.FileItem.IsSelected);
            var totalCount = PrintSettingsItems.Count;

            if (selectedCount == 0)
            {
                chkSelectAll.IsChecked = false;
            }
            else if (selectedCount == totalCount)
            {
                chkSelectAll.IsChecked = true;
            }
            else
            {
                chkSelectAll.IsChecked = null; // 部分選択状態
            }
        }


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