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
        public int SubfolderDepth { get; set; } = 1;

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

        #region メソッド
        /// <summary>
        /// 指定されたパスがベースパスのサブディレクトリかどうかを判定
        /// </summary>
        /// <param name="basePath">ベースパス</param>
        /// <param name="targetPath">判定対象パス</param>
        /// <returns>サブディレクトリの場合はtrue</returns>
        private bool IsSubdirectory(string basePath, string targetPath)
        {
            try
            {
                var baseUri = new Uri(Path.GetFullPath(basePath).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar);
                var targetUri = new Uri(Path.GetFullPath(targetPath).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar);
                return baseUri.IsBaseOf(targetUri);
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 利用可能なカテゴリを読み込み
        /// </summary>
        private void LoadAvailableCategories()
        {
            var allProjects = ProjectManager.LoadProjects();
            availableCategories = ProjectManager.GetAvailableCategories(allProjects);
            
            // よく使うカテゴリを追加
            var defaultCategories = new List<string> { "仕事", "個人", "開発", "資料", "アーカイブ" };
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
            txtSubfolderDepth.Text = SubfolderDepth.ToString();
            chkUseCustomPdfPath.IsChecked = UseCustomPdfPath;
            txtCustomPdfPath.Text = CustomPdfPath;

            // カスタムPDFパスの有効/無効設定
            UpdateCustomPdfPathEnabled();
            
            // ヒントテキストの表示更新
            UpdateDropHints();
        }

        /// <summary>
        /// ドラッグ&ドロップヒントの表示更新
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
                    
                    // プロジェクト名空の場合はフォルダ名設定
                    if (string.IsNullOrEmpty(txtProjectName.Text))
                    {
                        txtProjectName.Text = Path.GetFileName(dialog.SelectedPath);
                    }
                    
                    // ヒント表示更新
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
                dialog.Description = "PDF保存フォルダを選択してください（フォルダパスのみ設定されます）";
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
                    // フォルダパスのみ設定（ファイル名は含めない）
                    // プロジェクトフォルダ配下の選択を禁止
                    if (!string.IsNullOrWhiteSpace(txtFolderPath.Text) && 
                        IsSubdirectory(txtFolderPath.Text, dialog.SelectedPath))
                    {
                        MessageBox.Show("PDF保存パスはプロジェクトフォルダ配下以外を選択してください。\n" +
                                      "配下を選択するとPDFファイルが対象に含まれてしまいます。", "エラー", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    txtCustomPdfPath.Text = dialog.SelectedPath;
                    
                    // ヒント表示更新
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// サブフォルダ読み込みチェック時
        /// </summary>
        private void ChkIncludeSubfolders_Checked(object sender, RoutedEventArgs e)
        {
            // サブフォルダを含む場合、カスタムPDFパス必須にする
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
        /// カスタムPDF保存パス入力欄の有効/無効更新
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

                // プロジェクトフォルダ配下の選択を禁止
                if (!string.IsNullOrWhiteSpace(txtFolderPath.Text) && 
                    IsSubdirectory(txtFolderPath.Text, txtCustomPdfPath.Text))
                {
                    MessageBox.Show("PDF保存パスはプロジェクトフォルダ配下以外を選択してください。\n" +
                                  "配下を選択するとPDFファイルが対象に含まれてしまいます。", "エラー", 
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
            
            // 階層数の取得と検証
            if (int.TryParse(txtSubfolderDepth.Text, out int depth))
            {
                SubfolderDepth = Math.Max(1, Math.Min(5, depth));
            }
            else
            {
                SubfolderDepth = 1; // デフォルト値
            }
            
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
                        
                        // プロジェクト名空の場合はフォルダ名設定
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
                            
                            // プロジェクト名空の場合はフォルダ名設定
                            if (string.IsNullOrEmpty(txtProjectName.Text))
                            {
                                txtProjectName.Text = Path.GetFileName(parentFolder);
                            }
                        }
                    }
                    
                    // ヒント表示更新
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
                        // フォルダパスのみ設定（ファイル名は含めない）
                        // プロジェクトフォルダ配下の選択を禁止
                        if (!string.IsNullOrWhiteSpace(txtFolderPath.Text) && 
                            IsSubdirectory(txtFolderPath.Text, droppedPath))
                        {
                            MessageBox.Show("PDF保存パスはプロジェクトフォルダ配下以外を選択してください。\n" +
                                          "配下を選択するとPDFファイルが対象に含まれてしまいます。", "エラー", 
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                        
                        txtCustomPdfPath.Text = droppedPath;
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            // プロジェクトフォルダ配下の選択を禁止
                            if (!string.IsNullOrWhiteSpace(txtFolderPath.Text) && 
                                IsSubdirectory(txtFolderPath.Text, parentFolder))
                            {
                                MessageBox.Show("PDF保存パスはプロジェクトフォルダ配下以外を選択してください。\n" +
                                              "配下を選択するとPDFファイルが対象に含まれてしまいます。", "エラー", 
                                    MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                            }
                            
                            txtCustomPdfPath.Text = parentFolder;
                        }
                    }
                    
                    // ヒント表示更新
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// フォルダパステキストボックスのドラッグエンター（互換性保持）
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
        /// フォルダパステキストボックスのドラッグオーバー（互換性保持）
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
        /// フォルダパステキストボックスのドラッグリーブ（互換性保持）
        /// </summary>
        private void TxtFolderPath_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            txtFolderPath.Background = System.Windows.Media.Brushes.White;
        }

        /// <summary>
        /// フォルダパステキストボックスのドロップ（互換性保持）
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
                        
                        // プロジェクト名空の場合はフォルダ名設定
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
                            
                            // プロジェクト名空の場合はフォルダ名設定
                            if (string.IsNullOrEmpty(txtProjectName.Text))
                            {
                                txtProjectName.Text = Path.GetFileName(parentFolder);
                            }
                        }
                    }
                    
                    // ヒント表示更新
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// カスタムPDF保存パステキストボックスのドラッグエンター（互換性保持）
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
        /// カスタムPDF保存パステキストボックスのドラッグオーバー（互換性保持）
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
        /// カスタムPDF保存パステキストボックスのドラッグリーブ（互換性保持）
        /// </summary>
        private void TxtCustomPdfPath_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            txtCustomPdfPath.Background = System.Windows.Media.Brushes.White;
        }

        /// <summary>
        /// カスタムPDF保存パステキストボックスのドロップ（互換性保持）
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
                        // フォルダパスのみ設定（ファイル名は含めない）
                        // プロジェクトフォルダ配下の選択を禁止
                        if (!string.IsNullOrWhiteSpace(txtFolderPath.Text) && 
                            IsSubdirectory(txtFolderPath.Text, droppedPath))
                        {
                            MessageBox.Show("PDF保存パスはプロジェクトフォルダ配下以外を選択してください。\n" +
                                          "配下を選択するとPDFファイルが対象に含まれてしまいます。", "エラー", 
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                        
                        txtCustomPdfPath.Text = droppedPath;
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            // プロジェクトフォルダ配下の選択を禁止
                            if (!string.IsNullOrWhiteSpace(txtFolderPath.Text) && 
                                IsSubdirectory(txtFolderPath.Text, parentFolder))
                            {
                                MessageBox.Show("PDF保存パスはプロジェクトフォルダ配下以外を選択してください。\n" +
                                              "配下を選択するとPDFファイルが対象に含まれてしまいます。", "エラー", 
                                    MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                            }
                            
                            txtCustomPdfPath.Text = parentFolder;
                        }
                    }
                    
                    // ヒント表示更新
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// 階層数テキストボックスの入力制限（数字のみ、1-5）
        /// </summary>
        private void TxtSubfolderDepth_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // 数字のみ許可
            if (!char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
                return;
            }

            var textBox = sender as System.Windows.Controls.TextBox;
            var newText = textBox.Text + e.Text;
            
            // 1-5の範囲のみ許可
            if (int.TryParse(newText, out int value))
            {
                if (value < 1 || value > 5)
                {
                    e.Handled = true;
                }
            }
            else
            {
                e.Handled = true;
            }
        }

        #endregion
    }
}