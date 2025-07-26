using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using Nico2PDF.Models;
using Nico2PDF.Services;
using MessageBox = System.Windows.MessageBox;
using DragEventArgs = System.Windows.DragEventArgs;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;

namespace Nico2PDF.Views
{
    /// <summary>
    /// プロジェクト編集ダイアログ
    /// </summary>
    public partial class ProjectEditDialog : Window
    {
        #region プロパティ
        /// <summary>
        /// プロジェクト名
        /// </summary>
        public string ProjectName { get; set; } = "";

        /// <summary>
        /// プロジェクトカテゴリ
        /// </summary>
        public string Category { get; set; } = "";

        /// <summary>
        /// フォルダパス
        /// </summary>
        public string FolderPath { get; set; } = "";

        /// <summary>
        /// サブフォルダを含むかどうか
        /// </summary>
        public bool IncludeSubfolders { get; set; } = false;

        /// <summary>
        /// カスタムPDF保存パスを使用するかどうか
        /// </summary>
        public bool UseCustomPdfPath { get; set; } = false;

        /// <summary>
        /// カスタムPDF保存パス
        /// </summary>
        public string CustomPdfPath { get; set; } = "";

        /// <summary>
        /// 利用可能なカテゴリリスト
        /// </summary>
        private List<string> availableCategories = new List<string>();
        #endregion

        #region コンストラクタ
        public ProjectEditDialog()
        {
            InitializeComponent();
            LoadAvailableCategories();
        }
        #endregion

        #region 初期化
        /// <summary>
        /// 利用可能なカテゴリを読み込み
        /// </summary>
        private void LoadAvailableCategories()
        {
            var allProjects = ProjectManager.LoadProjects();
            availableCategories = ProjectManager.GetAvailableCategories(allProjects);
            
            // よく使われるカテゴリを追加
            var defaultCategories = new List<string> { "業務", "個人", "開発", "資料", "アーカイブ" };
            foreach (var category in defaultCategories)
            {
                if (!availableCategories.Contains(category))
                {
                    availableCategories.Add(category);
                }
            }
            
            cmbCategory.ItemsSource = availableCategories;
        }
        #endregion

        #region イベントハンドラ
        /// <summary>
        /// ウィンドウ読み込み時
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtProjectName.Text = ProjectName;
            txtFolderPath.Text = FolderPath;
            cmbCategory.Text = Category;
            chkIncludeSubfolders.IsChecked = IncludeSubfolders;
            chkUseCustomPdfPath.IsChecked = UseCustomPdfPath;
            txtCustomPdfPath.Text = CustomPdfPath;

            // カスタムPDFパスの有効/無効を設定
            UpdateCustomPdfPathEnabled();
            
            // ヒントテキストの表示制御
            UpdateDropHints();
        }

        /// <summary>
        /// ドラッグ&ドロップヒントの表示を更新
        /// </summary>
        private void UpdateDropHints()
        {
            if (txtFolderDropHint != null)
            {
                txtFolderDropHint.Visibility = string.IsNullOrEmpty(txtFolderPath.Text) ? Visibility.Visible : Visibility.Collapsed;
            }
            
            if (txtPdfDropHint != null)
            {
                txtPdfDropHint.Visibility = string.IsNullOrEmpty(txtCustomPdfPath.Text) ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        /// <summary>
        /// フォルダ選択ボタンクリック時
        /// </summary>
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "プロジェクトフォルダを選択してください";
                if (!string.IsNullOrEmpty(txtFolderPath.Text))
                {
                    dialog.SelectedPath = txtFolderPath.Text;
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtFolderPath.Text = dialog.SelectedPath;
                    
                    // プロジェクト名が空の場合はフォルダ名を設定
                    if (string.IsNullOrEmpty(txtProjectName.Text))
                    {
                        txtProjectName.Text = Path.GetFileName(dialog.SelectedPath);
                    }
                    
                    // ヒント表示を更新
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// カスタムPDF保存パス選択ボタンクリック時
        /// </summary>
        private void BtnSelectCustomPdfPath_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "PDF保存フォルダを選択してください（フォルダパスのみが設定されます）";
                if (!string.IsNullOrEmpty(txtCustomPdfPath.Text))
                {
                    // 既存のパスがファイル名を含む場合は、ディレクトリパスのみを取得
                    var existingPath = txtCustomPdfPath.Text;
                    if (File.Exists(existingPath))
                    {
                        dialog.SelectedPath = Path.GetDirectoryName(existingPath) ?? "";
                    }
                    else if (Directory.Exists(existingPath))
                    {
                        dialog.SelectedPath = existingPath;
                    }
                    else
                    {
                        // パスの親ディレクトリが存在するかチェック
                        var parentDir = Path.GetDirectoryName(existingPath);
                        if (!string.IsNullOrEmpty(parentDir) && Directory.Exists(parentDir))
                        {
                            dialog.SelectedPath = parentDir;
                        }
                    }
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // フォルダパスのみを設定（ファイル名は含めない）
                    txtCustomPdfPath.Text = dialog.SelectedPath;
                    
                    // ヒント表示を更新
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// サブフォルダ読み込みチェック時
        /// </summary>
        private void ChkIncludeSubfolders_Checked(object sender, RoutedEventArgs e)
        {
            // サブフォルダを含む場合、カスタムPDFパスを必須にする
            chkUseCustomPdfPath.IsChecked = true;
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// サブフォルダ読み込みチェック解除時
        /// </summary>
        private void ChkIncludeSubfolders_Unchecked(object sender, RoutedEventArgs e)
        {
            // サブフォルダを含まない場合は任意
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// カスタムPDF保存パス使用チェック時
        /// </summary>
        private void ChkUseCustomPdfPath_Checked(object sender, RoutedEventArgs e)
        {
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// カスタムPDF保存パス使用チェック解除時
        /// </summary>
        private void ChkUseCustomPdfPath_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// カスタムPDF保存パス入力欄の有効/無効を更新
        /// </summary>
        private void UpdateCustomPdfPathEnabled()
        {
            var includeSubfolders = chkIncludeSubfolders.IsChecked == true;
            var useCustomPdfPath = chkUseCustomPdfPath.IsChecked == true;
            
            // サブフォルダを含む場合は、カスタムPDFパスを強制的に有効にする
            if (includeSubfolders)
            {
                chkUseCustomPdfPath.IsChecked = true;
                chkUseCustomPdfPath.IsEnabled = false; // チェックボックスを無効化（必須）
                gridCustomPdfPath.IsEnabled = true;
            }
            else
            {
                chkUseCustomPdfPath.IsEnabled = true; // チェックボックスを有効化（任意）
                gridCustomPdfPath.IsEnabled = useCustomPdfPath;
            }
        }

        /// <summary>
        /// OKボタンクリック時
        /// </summary>
        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtProjectName.Text))
            {
                MessageBox.Show("プロジェクト名を入力してください。", "エラー", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtFolderPath.Text))
            {
                MessageBox.Show("フォルダを選択してください。", "エラー", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!Directory.Exists(txtFolderPath.Text))
            {
                MessageBox.Show("選択されたフォルダが存在しません。", "エラー", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (chkIncludeSubfolders.IsChecked == true && chkUseCustomPdfPath.IsChecked != true)
            {
                MessageBox.Show("サブフォルダを含む設定の場合、カスタムPDF保存パスの設定が必須です。", "エラー", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (chkUseCustomPdfPath.IsChecked == true)
            {
                if (string.IsNullOrWhiteSpace(txtCustomPdfPath.Text))
                {
                    MessageBox.Show("カスタムPDF保存パスを選択してください。", "エラー", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!Directory.Exists(txtCustomPdfPath.Text))
                {
                    var result = MessageBox.Show("指定されたPDF保存フォルダが存在しません。作成しますか？", "確認", 
                        MessageBoxButton.YesNo, MessageBoxImage.Question);
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        try
                        {
                            Directory.CreateDirectory(txtCustomPdfPath.Text);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"フォルダの作成に失敗しました: {ex.Message}", "エラー", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }
                }
            }

            ProjectName = txtProjectName.Text.Trim();
            FolderPath = txtFolderPath.Text.Trim();
            Category = cmbCategory.Text?.Trim() ?? "";
            IncludeSubfolders = chkIncludeSubfolders.IsChecked == true;
            UseCustomPdfPath = chkUseCustomPdfPath.IsChecked == true;
            CustomPdfPath = txtCustomPdfPath.Text.Trim();
            
            DialogResult = true;
            Close();
        }

        /// <summary>
        /// キャンセルボタンクリック時
        /// </summary>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        #endregion

        #region ドラッグ&ドロップ処理
        /// <summary>
        /// フォルダパス用ドラッグ&ドロップエリアのDragEnter
        /// </summary>
        private void FolderDropArea_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // ドラッグオーバー時の視覚的フィードバック
                if (sender is Border border)
                {
                    border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightBlue);
                    border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.DodgerBlue);
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// フォルダパス用ドラッグ&ドロップエリアのDragOver
        /// </summary>
        private void FolderDropArea_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// フォルダパス用ドラッグ&ドロップエリアのDragLeave
        /// </summary>
        private void FolderDropArea_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            if (sender is Border border)
            {
                border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250));
                border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 122, 204));
            }
        }

        /// <summary>
        /// フォルダパス用ドラッグ&ドロップエリアのDrop
        /// </summary>
        private void FolderDropArea_Drop(object sender, DragEventArgs e)
        {
            // ドラッグオーバー時の視覚的フィードバックを元に戻す
            if (sender is Border border)
            {
                border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250));
                border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 122, 204));
            }
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // フォルダかファイルかを判定
                    if (Directory.Exists(droppedPath))
                    {
                        txtFolderPath.Text = droppedPath;
                        
                        // プロジェクト名が空の場合はフォルダ名を設定
                        if (string.IsNullOrEmpty(txtProjectName.Text))
                        {
                            txtProjectName.Text = Path.GetFileName(droppedPath);
                        }
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            txtFolderPath.Text = parentFolder;
                            
                            // プロジェクト名が空の場合はフォルダ名を設定
                            if (string.IsNullOrEmpty(txtProjectName.Text))
                            {
                                txtProjectName.Text = Path.GetFileName(parentFolder);
                            }
                        }
                    }
                    
                    // ヒント表示を更新
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// PDF保存パス用ドラッグ&ドロップエリアのDragEnter
        /// </summary>
        private void PdfDropArea_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // ドラッグオーバー時の視覚的フィードバック
                if (sender is Border border)
                {
                    border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightBlue);
                    border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.DodgerBlue);
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// PDF保存パス用ドラッグ&ドロップエリアのDragOver
        /// </summary>
        private void PdfDropArea_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// PDF保存パス用ドラッグ&ドロップエリアのDragLeave
        /// </summary>
        private void PdfDropArea_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            if (sender is Border border)
            {
                border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250));
                border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 122, 204));
            }
        }

        /// <summary>
        /// PDF保存パス用ドラッグ&ドロップエリアのDrop
        /// </summary>
        private void PdfDropArea_Drop(object sender, DragEventArgs e)
        {
            // ドラッグオーバー時の視覚的フィードバックを元に戻す
            if (sender is Border border)
            {
                border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250));
                border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 122, 204));
            }
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // フォルダかファイルかを判定
                    if (Directory.Exists(droppedPath))
                    {
                        // フォルダパスのみを設定（ファイル名は含めない）
                        txtCustomPdfPath.Text = droppedPath;
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            txtCustomPdfPath.Text = parentFolder;
                        }
                    }
                    
                    // ヒント表示を更新
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// フォルダパステキストボックスのドラッグエンター（旧版の互換性維持）
        /// </summary>
        private void TxtFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // ドラッグオーバー時の視覚的フィードバック
                txtFolderPath.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightBlue);
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// フォルダパステキストボックスのドラッグオーバー（旧版の互換性維持）
        /// </summary>
        private void TxtFolderPath_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// フォルダパステキストボックスのドラッグリーブ（旧版の互換性維持）
        /// </summary>
        private void TxtFolderPath_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            txtFolderPath.Background = System.Windows.Media.Brushes.White;
        }

        /// <summary>
        /// フォルダパステキストボックスのドロップ（旧版の互換性維持）
        /// </summary>
        private void TxtFolderPath_Drop(object sender, DragEventArgs e)
        {
            // ドラッグオーバー時の視覚的フィードバックを元に戻す
            txtFolderPath.Background = System.Windows.Media.Brushes.White;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // フォルダかファイルかを判定
                    if (Directory.Exists(droppedPath))
                    {
                        txtFolderPath.Text = droppedPath;
                        
                        // プロジェクト名が空の場合はフォルダ名を設定
                        if (string.IsNullOrEmpty(txtProjectName.Text))
                        {
                            txtProjectName.Text = Path.GetFileName(droppedPath);
                        }
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            txtFolderPath.Text = parentFolder;
                            
                            // プロジェクト名が空の場合はフォルダ名を設定
                            if (string.IsNullOrEmpty(txtProjectName.Text))
                            {
                                txtProjectName.Text = Path.GetFileName(parentFolder);
                            }
                        }
                    }
                    
                    // ヒント表示を更新
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// カスタムPDF保存パステキストボックスのドラッグエンター（旧版の互換性維持）
        /// </summary>
        private void TxtCustomPdfPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // ドラッグオーバー時の視覚的フィードバック
                txtCustomPdfPath.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightBlue);
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// カスタムPDF保存パステキストボックスのドラッグオーバー（旧版の互換性維持）
        /// </summary>
        private void TxtCustomPdfPath_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// カスタムPDF保存パステキストボックスのドラッグリーブ（旧版の互換性維持）
        /// </summary>
        private void TxtCustomPdfPath_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            txtCustomPdfPath.Background = System.Windows.Media.Brushes.White;
        }

        /// <summary>
        /// カスタムPDF保存パステキストボックスのドロップ（旧版の互換性維持）
        /// </summary>
        private void TxtCustomPdfPath_Drop(object sender, DragEventArgs e)
        {
            // ドラッグオーバー時の視覚的フィードバックを元に戻す
            txtCustomPdfPath.Background = System.Windows.Media.Brushes.White;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // フォルダかファイルかを判定
                    if (Directory.Exists(droppedPath))
                    {
                        // フォルダパスのみを設定（ファイル名は含めない）
                        txtCustomPdfPath.Text = droppedPath;
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            txtCustomPdfPath.Text = parentFolder;
                        }
                    }
                    
                    // ヒント表示を更新
                    UpdateDropHints();
                }
            }
        }
        #endregion
    }
}