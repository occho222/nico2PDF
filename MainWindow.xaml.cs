using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using Nico2PDF.Models;
using Nico2PDF.Services;
using Nico2PDF.Views;
using MessageBox = System.Windows.MessageBox;
using DragEventArgs = System.Windows.DragEventArgs;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;

namespace Nico2PDF
{
    /// <summary>
    /// メインウィンドウ
    /// </summary>
    public partial class MainWindow : Window
    {
        #region プライベートフィールド
        private ObservableCollection<FileItem> fileItems = new ObservableCollection<FileItem>();
        private ObservableCollection<ProjectCategoryGroup> categoryGroups = new ObservableCollection<ProjectCategoryGroup>();
        private ProjectData? currentProject = null;
        private string selectedFolderPath = "";
        private string pdfOutputFolder = "";
        private AppSettings appSettings = new AppSettings();
        #endregion

        #region コンストラクタ
        public MainWindow()
        {
            InitializeComponent();
            InitializeDataBindings();
            LoadAppSettings();
            LoadProjects();
            RestoreActiveProject();
            UpdateProjectDisplay();
            SetVersionInfo();
        }
        #endregion

        #region 初期化
        private void InitializeDataBindings()
        {
            dgFiles.ItemsSource = fileItems;
            treeProjects.ItemsSource = categoryGroups;
        }

        private void LoadProjects()
        {
            categoryGroups.Clear();
            var projectList = ProjectManager.LoadProjects();
            
            // 既存プロジェクトのアイコンを修正
            FixExistingProjectIcons(projectList);
            
            // 各プロジェクトのアイコンを確認・設定
            foreach (var project in projectList)
            {
                if (string.IsNullOrEmpty(project.CategoryIcon))
                {
                    project.CategoryIcon = GetDefaultCategoryIcon(project.Category);
                }
            }
            
            // カテゴリ別にグループ化
            var groupedProjects = projectList.GroupBy(p => string.IsNullOrEmpty(p.Category) ? "未分類" : p.Category)
                                            .OrderBy(g => g.Key == "未分類" ? "z" : g.Key)
                                            .ToList();

            // アクティブプロジェクトのカテゴリを取得
            var activeProject = projectList.FirstOrDefault(p => p.IsActive);
            var activeCategory = activeProject?.Category ?? "";
            if (string.IsNullOrEmpty(activeCategory))
            {
                activeCategory = "未分類";
            }

            foreach (var group in groupedProjects)
            {
                var categoryGroup = new ProjectCategoryGroup
                {
                    CategoryName = group.Key,
                    CategoryIcon = GetCategoryIcon(group.Key, group.First().CategoryIcon),
                    CategoryColor = GetCategoryColor(group.Key, group.First().CategoryColor),
                    // アクティブプロジェクトがあるカテゴリのみ展開、それ以外は閉じる
                    IsExpanded = group.Key == activeCategory
                };

                // カテゴリ内でプロジェクト名順に並び替え
                var sortedProjects = group.OrderBy(p => p.Name).ToList();
                foreach (var project in sortedProjects)
                {
                    categoryGroup.Projects.Add(project);
                }

                categoryGroups.Add(categoryGroup);
            }
        }

        /// <summary>
        /// 既存プロジェクトのアイコンを修正
        /// </summary>
        private void FixExistingProjectIcons(List<ProjectData> projects)
        {
            bool needsSave = false;
            
            foreach (var project in projects)
            {
                // 空のアイコンやデフォルト値の修正
                if (string.IsNullOrEmpty(project.CategoryIcon) || project.CategoryIcon == "??")
                {
                    project.CategoryIcon = GetDefaultCategoryIcon(project.Category);
                    needsSave = true;
                }
                
                // 空の色やデフォルト値の修正
                if (string.IsNullOrEmpty(project.CategoryColor))
                {
                    project.CategoryColor = GetCategoryColor(project.Category, "");
                    needsSave = true;
                }
            }
            
            // 修正があった場合は保存
            if (needsSave)
            {
                ProjectManager.SaveProjects(projects);
            }
        }

        private string GetDefaultCategoryIcon(string category)
        {
            return category switch
            {
                "業務" => "💼",
                "プロジェクト" => "📊",
                "資料" => "📋",
                "マニュアル" => "📖",
                "提案書" => "📝",
                "報告書" => "📄",
                "会議" => "🗣️",
                "設計" => "⚙️",
                "テスト" => "🧪",
                "開発" => "💻",
                "運用" => "🔧",
                "保守" => "🛠️",
                "バックアップ" => "💾",
                "アーカイブ" => "📦",
                "一時的" => "⏱️",
                "進行中" => "🔄",
                "完了" => "✅",
                "保留" => "⏸️",
                "重要" => "⭐",
                "緊急" => "🚨",
                _ => "📁"
            };
        }

        private string GetCategoryIcon(string categoryName, string existingIcon)
        {
            if (!string.IsNullOrEmpty(existingIcon) && existingIcon != "📁")
            {
                return existingIcon;
            }
            return GetDefaultCategoryIcon(categoryName);
        }

        private string GetCategoryColor(string categoryName, string existingColor)
        {
            if (!string.IsNullOrEmpty(existingColor) && existingColor != "#E9ECEF")
            {
                return existingColor;
            }
            
            return categoryName switch
            {
                "業務" => "#007ACC",
                "プロジェクト" => "#28A745",
                "資料" => "#6C757D",
                "マニュアル" => "#17A2B8",
                "提案書" => "#FFC107",
                "報告書" => "#DC3545",
                "会議" => "#6F42C1",
                "設計" => "#FD7E14",
                "テスト" => "#20C997",
                "開発" => "#E83E8C",
                "運用" => "#6C757D",
                "保守" => "#495057",
                "バックアップ" => "#ADB5BD",
                "アーカイブ" => "#868E96",
                "一時的" => "#F8F9FA",
                "進行中" => "#007BFF",
                "完了" => "#28A745",
                "保留" => "#FFC107",
                "重要" => "#FF6B6B",
                "緊急" => "#DC3545",
                _ => "#E9ECEF"
            };
        }

        private void RestoreActiveProject()
        {
            var activeProject = GetAllProjects().FirstOrDefault(p => p.IsActive);
            if (activeProject != null)
            {
                SwitchToProject(activeProject);
            }
            else
            {
                UpdateLatestMergedPdfDisplay();
            }
        }

        /// <summary>
        /// 全プロジェクトを取得
        /// </summary>
        /// <returns>全プロジェクトのリスト</returns>
        private List<ProjectData> GetAllProjects()
        {
            var allProjects = new List<ProjectData>();
            foreach (var categoryGroup in categoryGroups)
            {
                allProjects.AddRange(categoryGroup.Projects);
            }
            return allProjects;
        }

        /// <summary>
        /// バージョン情報を設定
        /// </summary>
        private void SetVersionInfo()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                if (version != null)
                {
                    txtVersion.Text = $"v{version.Major}.{version.Minor}.{version.Build}";
                }
            }
            catch
            {
                txtVersion.Text = "v1.4.0";
            }
        }

        /// <summary>
        /// アプリケーション設定を読み込む
        /// </summary>
        private void LoadAppSettings()
        {
            appSettings = AppSettings.Load();
            
            // UI要素に設定を反映
            txtHeaderText.Text = appSettings.HeaderText;
            txtFooterText.Text = appSettings.FooterText;
            chkAddHeader.IsChecked = appSettings.AddHeader;
            chkAddFooter.IsChecked = appSettings.AddFooter;
            txtMergeFileName.Text = appSettings.MergeFileName;
            chkAddPageNumber.IsChecked = appSettings.AddPageNumber;
            chkAddBookmarks.IsChecked = appSettings.AddBookmarks;
            chkGroupByFolder.IsChecked = appSettings.GroupByFolder;
        }

        /// <summary>
        /// アプリケーション設定を保存する
        /// </summary>
        private void SaveAppSettings()
        {
            appSettings.UpdateFromMainWindow(
                txtHeaderText.Text ?? "",
                txtFooterText.Text ?? "",
                chkAddHeader.IsChecked ?? false,
                chkAddFooter.IsChecked ?? false,
                txtMergeFileName.Text ?? "結合PDF",
                chkAddPageNumber.IsChecked ?? false,
                chkAddBookmarks.IsChecked ?? true,
                chkGroupByFolder.IsChecked ?? false
            );
            
            appSettings.Save();
        }

        /// <summary>
        /// 設定変更時に呼び出されるイベントハンドラ
        /// </summary>
        private void OnSettingsChanged(object sender, RoutedEventArgs e)
        {
            // アプリ起動中は設定の自動保存を行う
            if (IsLoaded)
            {
                SaveAppSettings();
            }
        }
        #endregion

        #region プロジェクト管理
        private void BtnNewProject_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new ProjectEditDialog();
            if (dialog.ShowDialog() == true)
            {
                var newProject = new ProjectData
                {
                    Name = dialog.ProjectName,
                    FolderPath = dialog.FolderPath,
                    Category = dialog.Category,
                    IncludeSubfolders = dialog.IncludeSubfolders,
                    SubfolderDepth = dialog.SubfolderDepth,
                    UseCustomPdfPath = dialog.UseCustomPdfPath,
                    CustomPdfPath = dialog.CustomPdfPath,
                    CategoryIcon = GetDefaultCategoryIcon(dialog.Category),
                    CategoryColor = GetCategoryColor(dialog.Category, "")
                };

                // カテゴリグループに追加
                AddProjectToCategoryGroup(newProject);
                SwitchToProject(newProject);
                
                // プロジェクトリストを再構築
                RefreshProjectList();
            }
        }

        private void BtnEditProject_Click(object sender, RoutedEventArgs e)
        {
            if (treeProjects.SelectedItem is ProjectData selectedProject)
            {
                var dialog = new ProjectEditDialog();
                dialog.ProjectName = selectedProject.Name;
                dialog.FolderPath = selectedProject.FolderPath;
                dialog.Category = selectedProject.Category;
                dialog.IncludeSubfolders = selectedProject.IncludeSubfolders;
                dialog.SubfolderDepth = selectedProject.SubfolderDepth;
                dialog.UseCustomPdfPath = selectedProject.UseCustomPdfPath;
                dialog.CustomPdfPath = selectedProject.CustomPdfPath;

                if (dialog.ShowDialog() == true)
                {
                    selectedProject.Name = dialog.ProjectName;
                    selectedProject.FolderPath = dialog.FolderPath;
                    selectedProject.Category = dialog.Category;
                    selectedProject.IncludeSubfolders = dialog.IncludeSubfolders;
                    selectedProject.SubfolderDepth = dialog.SubfolderDepth;
                    selectedProject.UseCustomPdfPath = dialog.UseCustomPdfPath;
                    selectedProject.CustomPdfPath = dialog.CustomPdfPath;
                    
                    // カテゴリが変更された場合、アイコンと色を更新
                    selectedProject.CategoryIcon = GetDefaultCategoryIcon(dialog.Category);
                    selectedProject.CategoryColor = GetCategoryColor(dialog.Category, "");

                    if (selectedProject == currentProject)
                    {
                        selectedFolderPath = selectedProject.FolderPath;
                        pdfOutputFolder = selectedProject.PdfOutputFolder;
                        txtFolderPath.Text = selectedFolderPath;
                        UpdateProjectDisplay();
                    }

                    SaveProjects();
                    RefreshProjectList();
                }
            }
            else
            {
                MessageBox.Show("編集するプロジェクトを選択してください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnDeleteProject_Click(object sender, RoutedEventArgs e)
        {
            if (treeProjects.SelectedItem is ProjectData selectedProject)
            {
                var result = MessageBox.Show($"プロジェクト '{selectedProject.Name}' を削除しますか？",
                    "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    // カテゴリグループから削除
                    RemoveProjectFromCategoryGroup(selectedProject);

                    if (selectedProject == currentProject)
                    {
                        currentProject = null;
                        fileItems.Clear();
                        selectedFolderPath = "";
                        pdfOutputFolder = "";
                        txtFolderPath.Text = "";
                        UpdateProjectDisplay();
                    }

                    SaveProjects();
                }
            }
            else
            {
                MessageBox.Show("削除するプロジェクトを選択してください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnSwitchProject_Click(object sender, RoutedEventArgs e)
        {
            if (treeProjects.SelectedItem is ProjectData selectedProject)
            {
                SwitchToProject(selectedProject);
            }
            else
            {
                MessageBox.Show("プロジェクトを選択してください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void TreeProjects_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (treeProjects.SelectedItem is ProjectData selectedProject)
            {
                SwitchToProject(selectedProject);
            }
        }

        private void BtnConvertToProject_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("先にフォルダを選択してください。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var folderName = Path.GetFileName(selectedFolderPath);
            var dialog = new ProjectEditDialog();
            dialog.ProjectName = folderName;
            dialog.FolderPath = selectedFolderPath;

            if (dialog.ShowDialog() == true)
            {
                var newProject = new ProjectData
                {
                    Name = dialog.ProjectName,
                    FolderPath = dialog.FolderPath,
                    Category = dialog.Category,
                    IncludeSubfolders = dialog.IncludeSubfolders,
                    SubfolderDepth = dialog.SubfolderDepth,
                    UseCustomPdfPath = dialog.UseCustomPdfPath,
                    CustomPdfPath = dialog.CustomPdfPath,
                    MergeFileName = txtMergeFileName.Text,
                    AddPageNumber = chkAddPageNumber.IsChecked ?? false,
                    AddBookmarks = chkAddBookmarks.IsChecked ?? true,
                    GroupByFolder = chkGroupByFolder.IsChecked ?? false,
                    CategoryIcon = GetDefaultCategoryIcon(dialog.Category),
                    CategoryColor = GetCategoryColor(dialog.Category, "")
                };

                // ヘッダフッタ設定
                var chkAddHeaderFooter = FindName("chkAddHeaderFooter") as System.Windows.Controls.CheckBox;
                var txtHeaderFooterText = FindName("txtHeaderFooterText") as System.Windows.Controls.TextBox;
                var txtHeaderFooterFontSize = FindName("txtHeaderFooterFontSize") as System.Windows.Controls.TextBox;
                
                if (chkAddHeaderFooter != null)
                    newProject.AddHeaderFooter = chkAddHeaderFooter.IsChecked ?? false;
                if (txtHeaderFooterText != null)
                    newProject.HeaderFooterText = txtHeaderFooterText.Text;

                // ヘッダフッタフォントサイズを設定
                if (txtHeaderFooterFontSize != null && float.TryParse(txtHeaderFooterFontSize.Text, out float fontSize))
                {
                    newProject.HeaderFooterFontSize = fontSize;
                }

                // 現在のファイル状態を保存
                foreach (var item in fileItems)
                {
                    newProject.FileItems.Add(new FileItemData
                    {
                        IsSelected = item.IsSelected,
                        TargetPages = item.TargetPages,
                        FilePath = item.FilePath,
                        LastModified = item.LastModified,
                        DisplayOrder = item.DisplayOrder,
                        RelativePath = item.RelativePath
                    });
                }

                AddProjectToCategoryGroup(newProject);
                SwitchToProject(newProject);

                MessageBox.Show($"プロジェクト '{newProject.Name}' を作成しました。", "完了",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void SwitchToProject(ProjectData project)
        {
            SaveCurrentProjectState();

            // 全プロジェクトのアクティブ状態をリセット
            foreach (var p in GetAllProjects())
            {
                p.IsActive = false;
            }

            // 新しいプロジェクトをアクティブに設定
            project.IsActive = true;
            currentProject = project;

            // UIを更新
            selectedFolderPath = project.FolderPath;
            pdfOutputFolder = project.PdfOutputFolder;
            txtFolderPath.Text = selectedFolderPath;
            txtMergeFileName.Text = project.MergeFileName;
            chkAddPageNumber.IsChecked = project.AddPageNumber;
            chkAddBookmarks.IsChecked = project.AddBookmarks;
            chkGroupByFolder.IsChecked = project.GroupByFolder;
            
            // ヘッダ・フッタの個別設定をUIに反映
            chkAddHeader.IsChecked = project.AddHeader;
            chkAddFooter.IsChecked = project.AddFooter;
            txtHeaderText.Text = project.HeaderText;
            txtFooterText.Text = project.FooterText;


            UpdateLatestMergedPdfDisplay();
            RestoreFileItems(project);
            UpdateProjectDisplay();
            UpdateProjectTitle();
            SaveProjects();
        }

        private void SaveProjects()
        {
            var allProjects = GetAllProjects();
            ProjectManager.SaveProjects(allProjects);
        }

        private void SaveCurrentProjectState()
        {
            if (currentProject != null)
            {
                currentProject.FolderPath = selectedFolderPath;
                currentProject.PdfOutputFolder = pdfOutputFolder;
                currentProject.MergeFileName = txtMergeFileName.Text;
                currentProject.AddPageNumber = chkAddPageNumber.IsChecked ?? false;
                currentProject.AddBookmarks = chkAddBookmarks.IsChecked ?? true;
                currentProject.GroupByFolder = chkGroupByFolder.IsChecked ?? false;
                
                // ヘッダ・フッタ設定をプロジェクトに保存
                currentProject.AddHeader = chkAddHeader.IsChecked ?? false;
                currentProject.AddFooter = chkAddFooter.IsChecked ?? false;
                currentProject.HeaderText = txtHeaderText.Text ?? "";
                currentProject.FooterText = txtFooterText.Text ?? "";

                // ヘッダフッタのUIコントロールを検索して設定
                var chkAddHeaderFooter = FindName("chkAddHeaderFooter") as System.Windows.Controls.CheckBox;
                var txtHeaderFooterText = FindName("txtHeaderFooterText") as System.Windows.Controls.TextBox;
                var txtHeaderFooterFontSize = FindName("txtHeaderFooterFontSize") as System.Windows.Controls.TextBox;
                
                if (chkAddHeaderFooter != null)
                    currentProject.AddHeaderFooter = chkAddHeaderFooter.IsChecked ?? false;
                if (txtHeaderFooterText != null)
                    currentProject.HeaderFooterText = txtHeaderFooterText.Text;
                if (txtHeaderFooterFontSize != null && float.TryParse(txtHeaderFooterFontSize.Text, out float fontSize))
                {
                    currentProject.HeaderFooterFontSize = fontSize;
                }

                
                currentProject.LastAccessDate = DateTime.Now;

                // ファイルアイテムの状態を保存
                currentProject.FileItems.Clear();
                foreach (var item in fileItems)
                {
                    currentProject.FileItems.Add(new FileItemData
                    {
                        IsSelected = item.IsSelected,
                        TargetPages = item.TargetPages,
                        FilePath = item.FilePath,
                        LastModified = item.LastModified,
                        DisplayOrder = item.DisplayOrder,
                        RelativePath = item.RelativePath,
                        DisplayName = item.DisplayName,
                        OriginalFileName = item.OriginalFileName
                    });
                }

                SaveProjects();
            }
            
            // アプリケーション設定も同時に保存
            SaveAppSettings();
        }

        private void UpdateProjectTitle()
        {
            if (currentProject != null)
            {
                txtCurrentProjectTitle.Text = $"- {currentProject.Name}";
            }
            else
            {
                txtCurrentProjectTitle.Text = "";
            }
        }

        private void UpdateProjectDisplay()
        {
            if (currentProject != null)
            {
                lblCurrentProject.Content = $"現在のプロジェクト: {currentProject.Name}";
                
                // バージョン情報も含めてタイトルを設定
                var assembly = Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                var versionText = version != null ? $"v{version.Major}.{version.Minor}.{version.Build}" : "v1.4.0";
                Title = $"nico²PDF {versionText} - {currentProject.Name}";
            }
            else
            {
                lblCurrentProject.Content = "現在のプロジェクト: なし";
                
                // バージョン情報も含めてタイトルを設定
                var assembly = Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                var versionText = version != null ? $"v{version.Major}.{version.Minor}.{version.Build}" : "v1.4.0";
                Title = $"nico²PDF {versionText}";
            }
            
            // プロジェクトタイトルも更新
            UpdateProjectTitle();
            
            // ヒントテキストの表示制御
            if (txtDropHint != null)
            {
                txtDropHint.Visibility = string.IsNullOrEmpty(selectedFolderPath) ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        /// <summary>
        /// プロジェクトリストをカテゴリ順序で再構築
        /// </summary>
        private void RefreshProjectList()
        {
            var currentSelectedProject = treeProjects.SelectedItem as ProjectData;
            
            // カテゴリグループをクリアして再構築
            LoadProjects();
            
            // 選択状態を復元
            if (currentSelectedProject != null)
            {
                SelectProjectInTree(currentSelectedProject.Id);
            }
        }

        /// <summary>
        /// TreeViewで指定IDのプロジェクトを選択
        /// </summary>
        private void SelectProjectInTree(string projectId)
        {
            foreach (var categoryGroup in categoryGroups)
            {
                var project = categoryGroup.Projects.FirstOrDefault(p => p.Id == projectId);
                if (project != null)
                {
                    // TreeViewItemを見つけて選択
                    var treeViewItem = FindTreeViewItem(treeProjects, project);
                    if (treeViewItem != null)
                    {
                        treeViewItem.IsSelected = true;
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// TreeViewItemを検索
        /// </summary>
        private TreeViewItem FindTreeViewItem(System.Windows.Controls.TreeView treeView, object item)
        {
            return FindTreeViewItem(treeView, item, treeView.ItemContainerGenerator);
        }

        /// <summary>
        /// TreeViewItemを再帰的に検索
        /// </summary>
        private TreeViewItem FindTreeViewItem(ItemsControl parent, object item, ItemContainerGenerator generator)
        {
            if (parent == null || item == null) return null;

            for (int i = 0; i < parent.Items.Count; i++)
            {
                var container = generator.ContainerFromIndex(i) as TreeViewItem;
                if (container != null)
                {
                    if (container.DataContext == item)
                        return container;

                    var child = FindTreeViewItem(container, item, container.ItemContainerGenerator);
                    if (child != null)
                        return child;
                }
            }
            return null;
        }

        private void BtnCategoryManage_Click(object sender, RoutedEventArgs e)
        {
            var allProjects = GetAllProjects();
            var dialog = new Views.CategoryManageDialog(allProjects);
            
            if (dialog.ShowDialog() == true)
            {
                // カテゴリ管理で変更があった場合、プロジェクト一覧を更新
                RefreshProjectList();
                
                // 現在のプロジェクトがアクティブな場合、UIを更新
                if (currentProject != null)
                {
                    UpdateProjectDisplay();
                }
            }
        }

        /// <summary>
        /// プロジェクトをカテゴリグループに追加
        /// </summary>
        private void AddProjectToCategoryGroup(ProjectData project)
        {
            var categoryName = string.IsNullOrEmpty(project.Category) ? "未分類" : project.Category;
            var existingGroup = categoryGroups.FirstOrDefault(g => g.CategoryName == categoryName);
            
            if (existingGroup == null)
            {
                existingGroup = new ProjectCategoryGroup
                {
                    CategoryName = categoryName,
                    CategoryIcon = GetCategoryIcon(categoryName, project.CategoryIcon),
                    CategoryColor = GetCategoryColor(categoryName, project.CategoryColor)
                };
                categoryGroups.Add(existingGroup);
            }
            
            existingGroup.Projects.Add(project);
        }

        /// <summary>
        /// プロジェクトをカテゴリグループから削除
        /// </summary>
        private void RemoveProjectFromCategoryGroup(ProjectData project)
        {
            foreach (var categoryGroup in categoryGroups.ToList())
            {
                if (categoryGroup.Projects.Contains(project))
                {
                    categoryGroup.Projects.Remove(project);
                    
                    // プロジェクトが空になったカテゴリグループは削除
                    if (categoryGroup.Projects.Count == 0)
                    {
                        categoryGroups.Remove(categoryGroup);
                    }
                    break;
                }
            }
        }
        #endregion

        #region ファイル管理
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "対象フォルダを選択してください（フォルダパスのみが設定されます）";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // フォルダパスのみを設定（ファイル名は含めない）
                    selectedFolderPath = dialog.SelectedPath;
                    txtFolderPath.Text = selectedFolderPath;
                    
                    // PDFアウトプットフォルダもフォルダパスのみに設定
                    if (currentProject != null && currentProject.UseCustomPdfPath && !string.IsNullOrEmpty(currentProject.CustomPdfPath))
                    {
                        pdfOutputFolder = currentProject.CustomPdfPath;
                    }
                    else
                    {
                        pdfOutputFolder = Path.Combine(selectedFolderPath, "PDF");
                    }
                    
                    if (currentProject != null)
                    {
                        currentProject.FolderPath = selectedFolderPath;
                        // PdfOutputFolderはプロパティで自動計算されるので直接設定しない
                        SaveProjects();
                    }
                    
                    txtStatus.Text = "フォルダが選択されました";
                    
                    // 自動的にファイル読み込みを実行
                    LoadFilesFromFolder();
                }
            }
        }

        /// <summary>
        /// フォルダからファイルを読み込む共通メソッド
        /// </summary>
        private void LoadFilesFromFolder()
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("フォルダを選択してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
            var subfolderDepth = currentProject?.SubfolderDepth ?? 1;
            
            // 削除されたファイルの処理も含めてファイル更新
            var (updatedItems, changedFiles, addedFiles, deletedFiles) = 
                FileManagementService.UpdateFiles(selectedFolderPath, pdfOutputFolder, fileItems.ToList(), includeSubfolders, subfolderDepth);
            
            fileItems.Clear();
            
            // 名前順でソートしてから追加
            var sortedItems = updatedItems.OrderBy(f => f.FileName).ToList();
            for (int i = 0; i < sortedItems.Count; i++)
            {
                sortedItems[i].Number = i + 1;
                sortedItems[i].DisplayOrder = i;
                fileItems.Add(sortedItems[i]);
            }

            // PDFステータスを更新（UIスレッドで実行）
            Dispatcher.BeginInvoke(() => RefreshAllPdfStatus());

            // 削除されたファイルがある場合はメッセージを表示
            var statusMessage = $"{fileItems.Count}個のファイルを読み込みました";
            if (deletedFiles.Count > 0)
            {
                statusMessage += $"（{deletedFiles.Count}個の不要なPDFファイルを削除しました）";
            }
            if (includeSubfolders)
            {
                statusMessage += " (サブフォルダを含む)";
            }
            txtStatus.Text = statusMessage;
            SaveCurrentProjectState();
        }

        private void BtnUpdateFiles_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("フォルダを選択してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
            var subfolderDepth = currentProject?.SubfolderDepth ?? 1;
            var (updatedItems, changedFiles, addedFiles, deletedFiles) = 
                FileManagementService.UpdateFiles(selectedFolderPath, pdfOutputFolder, fileItems.ToList(), includeSubfolders, subfolderDepth);

            fileItems.Clear();
            
            // 名前順でソートしてから追加
            var sortedItems = updatedItems.OrderBy(f => f.FileName).ToList();
            for (int i = 0; i < sortedItems.Count; i++)
            {
                sortedItems[i].Number = i + 1;
                sortedItems[i].DisplayOrder = i;
                fileItems.Add(sortedItems[i]);
            }

            // PDFステータスを更新（UIスレッドで実行）
            Dispatcher.BeginInvoke(() => RefreshAllPdfStatus());

            // 結果メッセージを作成
            var statusMessages = new List<string>();
            statusMessages.Add($"{fileItems.Count}個のファイルを更新しました");
            
            if (includeSubfolders)
            {
                statusMessages.Add("(サブフォルダを含む)");
            }

            if (changedFiles.Any())
                statusMessages.Add($"変更されたファイル: {changedFiles.Count}個");

            if (addedFiles.Any())
                statusMessages.Add($"追加されたファイル: {addedFiles.Count}個");

            if (deletedFiles.Any())
            {
                statusMessages.Add($"削除されたファイル: {deletedFiles.Count}個");
                
                var deletedMessage = $"以下のファイルが削除されました：\n{string.Join("\n", deletedFiles)}";
                if (deletedFiles.Any(f => !f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)))
                {
                    deletedMessage += "\n\n対応するPDFファイルも削除されました。";
                }
                MessageBox.Show(deletedMessage, "削除されたファイル", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            txtStatus.Text = string.Join(" / ", statusMessages);
            SaveCurrentProjectState();
        }

        private void RestoreFileItems(ProjectData project)
        {
            fileItems.Clear();

            if (string.IsNullOrEmpty(project.FolderPath) || !Directory.Exists(project.FolderPath))
            {
                txtStatus.Text = "プロジェクトフォルダが見つかりません";
                return;
            }

            // まず通常のLoadFilesFromFolderでファイルを読み込み
            var loadedFileItems = FileManagementService.LoadFilesFromFolder(project.FolderPath, project.PdfOutputFolder, project.IncludeSubfolders, project.SubfolderDepth);
            
            // 削除されたファイルの処理（現在のUIにあるfileItemsコレクションと比較）
            var deletedFiles = new List<string>();
            var currentFilePaths = loadedFileItems.Select(f => f.FilePath).ToHashSet();
            
            foreach (var item in fileItems.ToList())
            {
                if (!currentFilePaths.Contains(item.FilePath))
                {
                    // 対応するPDFファイルを削除
                    if (item.Extension.ToUpper() != "PDF")
                    {
                        var pdfPath = FileManagementService.GetPdfPath(item.FilePath, project.PdfOutputFolder, project.FolderPath, project.IncludeSubfolders);
                        if (File.Exists(pdfPath))
                        {
                            try
                            {
                                File.Delete(pdfPath);
                                deletedFiles.Add(item.FileName);
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"PDFファイル削除エラー: {ex.Message}");
                            }
                        }
                    }
                }
            }
            
            // 削除されたファイルがある場合はメッセージを表示
            if (deletedFiles.Count > 0)
            {
                MessageBox.Show($"{deletedFiles.Count}個の不要なPDFファイルを削除しました。\n削除されたファイル: {string.Join(", ", deletedFiles)}", 
                    "PDF削除完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            
            // 保存された状態を復元
            foreach (var item in loadedFileItems)
            {
                var savedItem = project.FileItems.FirstOrDefault(f => f.FilePath == item.FilePath);
                if (savedItem != null)
                {
                    item.IsSelected = savedItem.IsSelected;
                    item.TargetPages = savedItem.TargetPages;
                    item.DisplayOrder = savedItem.DisplayOrder >= 0 ? savedItem.DisplayOrder : loadedFileItems.Count;
                    item.DisplayName = savedItem.DisplayName;
                    item.OriginalFileName = savedItem.OriginalFileName;
                }
            }

            // 表示順序で並び替え
            var orderedItems = loadedFileItems
                .OrderBy(f => f.DisplayOrder)
                .ThenBy(f => f.RelativePath)
                .ThenBy(f => f.FileName)
                .ToList();

            for (int i = 0; i < orderedItems.Count; i++)
            {
                orderedItems[i].Number = i + 1;
                orderedItems[i].DisplayOrder = i;
                fileItems.Add(orderedItems[i]);
            }

            // PDFステータスを更新（UIスレッドで実行）
            Dispatcher.BeginInvoke(() => RefreshAllPdfStatus());

            var statusMessage = $"プロジェクト '{project.Name}' を読み込みました ({fileItems.Count}個のファイル)";
            if (project.IncludeSubfolders)
            {
                statusMessage += " (サブフォルダを含む)";
            }
            txtStatus.Text = statusMessage;
        }
        #endregion

        #region ファイル操作
        private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            if (dgFiles.SelectedIndex > 0)
            {
                var selectedIndex = dgFiles.SelectedIndex;
                var item = fileItems[selectedIndex];
                fileItems.RemoveAt(selectedIndex);
                fileItems.Insert(selectedIndex - 1, item);
                
                UpdateFileNumbers();
                dgFiles.SelectedIndex = selectedIndex - 1;
                SaveCurrentProjectState();
            }
        }

        private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            if (dgFiles.SelectedIndex >= 0 && dgFiles.SelectedIndex < fileItems.Count - 1)
            {
                var selectedIndex = dgFiles.SelectedIndex;
                var item = fileItems[selectedIndex];
                fileItems.RemoveAt(selectedIndex);
                fileItems.Insert(selectedIndex + 1, item);
                
                UpdateFileNumbers();
                dgFiles.SelectedIndex = selectedIndex + 1;
                SaveCurrentProjectState();
            }
        }

        private void BtnSortByName_Click(object sender, RoutedEventArgs e)
        {
            var sortedItems = fileItems.OrderBy(f => f.FileName).ToList();
            fileItems.Clear();
            
            for (int i = 0; i < sortedItems.Count; i++)
            {
                sortedItems[i].Number = i + 1;
                sortedItems[i].DisplayOrder = i;
                fileItems.Add(sortedItems[i]);
            }
            
            SaveCurrentProjectState();
            txtStatus.Text = "ファイル名順に並び替えました";
        }

        private void UpdateFileNumbers()
        {
            for (int i = 0; i < fileItems.Count; i++)
            {
                fileItems[i].Number = i + 1;
                fileItems[i].DisplayOrder = i;
            }
        }

        private void ChkSelectAll_Click(object sender, RoutedEventArgs e)
        {
            var isChecked = chkSelectAll.IsChecked ?? false;
            foreach (var item in fileItems)
            {
                item.IsSelected = isChecked;
            }
            SaveCurrentProjectState();
        }

        private void FileName_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBlock textBlock && textBlock.DataContext is FileItem fileItem)
            {
                OpenFile(fileItem.FilePath);
            }
        }

        private void OpenFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ファイルを開けませんでした: {ex.Message}", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("ファイルが見つかりません。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void BtnExportFileList_Click(object sender, RoutedEventArgs e)
        {
            if (!fileItems.Any())
            {
                MessageBox.Show("ファイル一覧が空です。先にフォルダを読み込んでください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new System.Windows.Forms.SaveFileDialog
            {
                Title = "ファイル一覧の保存先を選択",
                Filter = "テキストファイル (*.txt)|*.txt|CSVファイル (*.csv)|*.csv",
                FileName = $"ファイル一覧_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    var selectedOnly = MessageBox.Show("選択されたファイルのみを出力しますか？", 
                        "出力対象", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;
                    
                    var includeDetails = MessageBox.Show("詳細情報を含めますか？", 
                        "出力形式", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;

                    var extension = Path.GetExtension(dialog.FileName).ToLower();
                    
                    if (extension == ".csv")
                    {
                        FileManagementService.ExportFileListToCsv(fileItems.ToList(), dialog.FileName, 
                            includeHeaders: true, includeDetails: includeDetails, selectedOnly: selectedOnly);
                    }
                    else
                    {
                        FileManagementService.ExportFileList(fileItems.ToList(), dialog.FileName, 
                            includeHeaders: true, includeDetails: includeDetails, selectedOnly: selectedOnly);
                    }

                    txtStatus.Text = $"ファイル一覧を出力しました: {Path.GetFileName(dialog.FileName)}";
                    
                    var result = MessageBox.Show("ファイル一覧を出力しました。保存先フォルダを開きますか？", 
                        "出力完了", MessageBoxButton.YesNo, MessageBoxImage.Information);
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        try
                        {
                            System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{dialog.FileName}\"");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ファイル一覧の出力に失敗しました: {ex.Message}", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnCopyToClipboard_Click(object sender, RoutedEventArgs e)
        {
            if (!fileItems.Any())
            {
                MessageBox.Show("ファイル一覧が空です。先にフォルダを読み込んでください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                // 選択されたファイルのみをコピーするか確認
                var selectedOnly = MessageBox.Show("選択されたファイルのみをコピーしますか？", 
                    "コピー対象", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;

                // ファイル名のみをシンプルな改行区切り形式でクリップボードにコピー
                FileManagementService.CopyFileNamesToClipboard(fileItems.ToList(), selectedOnly);

                var itemCount = selectedOnly ? fileItems.Count(f => f.IsSelected) : fileItems.Count;
                txtStatus.Text = $"ファイル名をクリップボードにコピーしました ({itemCount}個のファイル)";
                
                MessageBox.Show($"ファイル名をクリップボードにコピーしました。\n\n" +
                               $"コピー件数: {itemCount}個\n" +
                               $"形式: ファイル名のみ（改行区切り）\n" +
                               $"対象: {(selectedOnly ? "選択されたファイルのみ" : "全ファイル")}", 
                               "コピー完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"クリップボードへのコピーに失敗しました: {ex.Message}", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region ファイル名変更機能
        private void BtnRenameFile_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = dgFiles.SelectedItem as FileItem;
            if (selectedItem == null)
            {
                MessageBox.Show("名前を変更するファイルを選択してください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            ShowRenameDialog(selectedItem);
        }

        private void BtnBatchRename_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = fileItems.Where(f => f.IsSelected).ToList();
            if (!selectedItems.Any())
            {
                MessageBox.Show("名前を変更するファイルを選択してください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            ShowBatchRenameDialog(selectedItems);
        }

        private void MenuRenameFile_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = dgFiles.SelectedItem as FileItem;
            if (selectedItem != null)
            {
                ShowRenameDialog(selectedItem);
            }
        }

        private void MenuBatchRename_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = fileItems.Where(f => f.IsSelected).ToList();
            if (selectedItems.Any())
            {
                ShowBatchRenameDialog(selectedItems);
            }
        }

        private void BtnExcelPrintSettings_Click(object sender, RoutedEventArgs e)
        {
            ShowExcelPrintSettingsDialog();
        }

        private void MenuExcelPrintSettings_Click(object sender, RoutedEventArgs e)
        {
            ShowExcelPrintSettingsDialog();
        }

        private void MenuResetDisplayName_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = dgFiles.SelectedItem as FileItem;
            if (selectedItem != null && selectedItem.IsRenamed)
            {
                var result = MessageBox.Show($"'{selectedItem.DisplayName}' の表示名を元のファイル名に戻しますか？", 
                    "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    selectedItem.ResetDisplayName();
                    SaveCurrentProjectState();
                    txtStatus.Text = "表示名をリセットしました";
                }
            }
        }

        private void MenuOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = dgFiles.SelectedItem as FileItem;
            if (selectedItem != null)
            {
                OpenFile(selectedItem.FilePath);
            }
        }

        private void MenuOpenInFolder_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = dgFiles.SelectedItem as FileItem;
            if (selectedItem != null)
            {
                try
                {
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{selectedItem.FilePath}\"");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ShowRenameDialog(FileItem fileItem)
        {
            var dialog = new Nico2PDF.Views.FileRenameDialog();
            dialog.Owner = this;
            dialog.CurrentFileName = fileItem.FileName;
            dialog.NewFileName = Path.GetFileNameWithoutExtension(fileItem.DisplayName);
            dialog.OriginalFilePath = fileItem.FilePath;

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var extension = Path.GetExtension(fileItem.FileName);
                    var newDisplayName = dialog.NewFileName + extension;
                    var oldFilePath = fileItem.FilePath;

                    // 物理ファイル名を変更
                    var newFilePath = FileManagementService.RenamePhysicalFile(fileItem.FilePath, dialog.NewFileName);
                    fileItem.FilePath = newFilePath;
                    fileItem.FileName = Path.GetFileName(newFilePath);
                    fileItem.DisplayName = newDisplayName;
                    
                    // 相対パスも更新
                    if (!string.IsNullOrEmpty(selectedFolderPath))
                    {
                        fileItem.RelativePath = GetRelativePath(selectedFolderPath, newFilePath);
                    }

                    // PDFファイルの処理（古いPDFファイル削除と新しいPDFファイル存在確認）
                    var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
                    var hasPdf = FileManagementService.HandlePdfFileAfterRename(
                        oldFilePath, newFilePath, pdfOutputFolder, selectedFolderPath, includeSubfolders);
                    
                    // PDFステータスを更新（リアルタイム）
                    fileItem.PdfStatus = hasPdf ? "変換済" : "未変換";
                    
                    // 未変換の場合は選択状態にする
                    if (!hasPdf)
                    {
                        fileItem.IsSelected = true;
                    }

                    SaveCurrentProjectState();
                    txtStatus.Text = $"ファイル名を変更しました: {newDisplayName}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ファイル名の変更に失敗しました: {ex.Message}", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ShowBatchRenameDialog(List<FileItem> selectedItems)
        {
            var dialog = new Nico2PDF.Views.BatchRenameDialog();
            dialog.Owner = this;
            dialog.TargetFiles = selectedItems;
            dialog.Initialize(); // 新しい初期化メソッドを呼び出し

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var successCount = 0;
                    var failCount = 0;
                    var errors = new List<string>();

                    // 変更されたファイルのみを処理
                    var changedItems = dialog.RenameItems.Where(item => item.IsChanged && !item.HasError).ToList();

                    foreach (var renameItem in changedItems)
                    {
                        try
                        {
                            var originalItem = renameItem.OriginalItem;
                            var oldFilePath = originalItem.FilePath;
                            var newFilePath = FileManagementService.RenamePhysicalFile(originalItem.FilePath, renameItem.NewFileName);
                            
                            // ファイルアイテムを更新
                            originalItem.FilePath = newFilePath;
                            originalItem.FileName = Path.GetFileName(newFilePath);
                            originalItem.DisplayName = renameItem.PreviewFileName;
                            
                            // 相対パスも更新
                            if (!string.IsNullOrEmpty(selectedFolderPath))
                            {
                                originalItem.RelativePath = GetRelativePath(selectedFolderPath, newFilePath);
                            }

                            // PDFファイルの処理（古いPDFファイル削除と新しいPDFファイル存在確認）
                            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
                            var hasPdf = FileManagementService.HandlePdfFileAfterRename(
                                oldFilePath, newFilePath, pdfOutputFolder, selectedFolderPath, includeSubfolders);
                            
                            // PDFステータスを更新（リアルタイム）
                            originalItem.PdfStatus = hasPdf ? "変換済" : "未変換";
                            
                            // 未変換の場合は選択状態にする
                            if (!hasPdf)
                            {
                                originalItem.IsSelected = true;
                            }
                            
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"{renameItem.CurrentFileName}: {ex.Message}");
                            failCount++;
                        }
                    }

                    if (failCount > 0)
                    {
                        var errorMessage = $"一部のファイルの名前変更に失敗しました:\n\n{string.Join("\n", errors)}";
                        MessageBox.Show(errorMessage, "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }

                    SaveCurrentProjectState();
                    txtStatus.Text = $"一括リネーム完了: 成功 {successCount}件, 失敗 {failCount}件";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"一括リネームに失敗しました: {ex.Message}", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ShowExcelPrintSettingsDialog()
        {
            var excelFiles = fileItems.Where(f => IsExcelFile(f.Extension)).ToList();
            if (!excelFiles.Any())
            {
                MessageBox.Show("Excelファイル（.xls, .xlsx, .xlsm）が見つかりません。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new Nico2PDF.Views.PrintSettingsDialog();
            dialog.Owner = this;
            dialog.SetFileItems(fileItems.ToList());

            if (dialog.ShowDialog() == true)
            {
                // ダイアログで何かしらの設定が適用された場合の処理
                // 現在のプロジェクト状態を保存
                SaveCurrentProjectState();
                txtStatus.Text = "Excel印刷設定が適用されました。";
            }
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

        #endregion

        #region PDF処理
        private async void BtnConvertPDF_Click(object sender, RoutedEventArgs e)
        {
            var selectedFiles = fileItems.Where(f => f.IsSelected).ToList();
            if (!selectedFiles.Any())
            {
                MessageBox.Show("変換するファイルを選択してください", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!Directory.Exists(pdfOutputFolder))
                Directory.CreateDirectory(pdfOutputFolder);

            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
            var baseFolderPath = selectedFolderPath;

            progressBar.Visibility = Visibility.Visible;
            progressBar.Maximum = selectedFiles.Count;
            progressBar.Value = 0;

            var convertedFiles = new List<FileItem>();

            await Task.Run(() =>
            {
                foreach (var file in selectedFiles)
                {
                    try
                    {
                        // サブフォルダ構造を考慮した変換
                        if (includeSubfolders)
                        {
                            FileManagementService.EnsurePdfOutputDirectory(file.FilePath, pdfOutputFolder, baseFolderPath, includeSubfolders);
                        }

                        PdfConversionService.ConvertToPdf(file.FilePath, pdfOutputFolder, file.TargetPages, baseFolderPath, includeSubfolders);

                        // 変換成功したファイルをリストに追加
                        convertedFiles.Add(file);

                        Dispatcher.Invoke(() =>
                        {
                            // 個別ファイルのステータス更新（即座に反映）
                            UpdateIndividualPdfStatus(file, includeSubfolders);
                            file.IsSelected = false;
                            progressBar.Value++;
                            txtStatus.Text = $"変換中: {file.FileName}";
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show($"変換エラー: {file.FileName}\n{ex.Message}", "エラー",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                        });
                    }
                }
            });

            progressBar.Visibility = Visibility.Collapsed;
            
            // 変換完了後に全体のステータスを確認・更新
            Dispatcher.Invoke(() =>
            {
                RefreshAllPdfStatus();
                
                var convertedCount = convertedFiles.Count;
                var totalSelected = selectedFiles.Count;
                
                if (convertedCount == totalSelected)
                {
                    txtStatus.Text = $"PDF変換が完了しました ({convertedCount}個のファイル)";
                }
                else
                {
                    txtStatus.Text = $"PDF変換が完了しました ({convertedCount}/{totalSelected}個のファイル)";
                }
                
                SaveCurrentProjectState();
            });
        }

        /// <summary>
        /// 個別ファイルのPDFステータスを即座に更新
        /// </summary>
        /// <param name="file">対象ファイル</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        private void UpdateIndividualPdfStatus(FileItem file, bool includeSubfolders)
        {
            // PDFファイル自体はスキップ
            if (file.Extension.ToLower() == ".pdf")
            {
                file.PdfStatus = "PDF";
                return;
            }

            // 対応するPDFファイルの存在を確認
            string pdfPath = GetPdfPath(file.FilePath, pdfOutputFolder, selectedFolderPath, includeSubfolders);
            bool pdfExists = File.Exists(pdfPath);
            
            file.PdfStatus = pdfExists ? "変換済" : "未変換";
        }

        private async void BtnMergePDF_Click(object sender, RoutedEventArgs e)
        {
            var allFiles = fileItems.OrderBy(f => f.DisplayOrder).ToList();
            if (!allFiles.Any())
            {
                MessageBox.Show("結合対象のファイルがありません", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // PDFファイルパスを取得
            var pdfFilePaths = new List<string>();
            var missingPdfFiles = new List<string>();
            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
            var baseFolderPath = selectedFolderPath;

            foreach (var file in allFiles)
            {
                string pdfPath;
                if (file.Extension.ToLower() == "pdf")
                {
                    // PDFファイルの場合も、PDFフォルダ内のファイルを使用
                    pdfPath = GetPdfPath(file.FilePath, pdfOutputFolder, baseFolderPath, includeSubfolders);
                }
                else
                {
                    if (includeSubfolders)
                    {
                        // サブフォルダ構造を考慮したパス
                        var fileInfo = new FileInfo(file.FilePath);
                        var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                        var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                        pdfPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(file.FileName) + ".pdf");
                    }
                    else
                    {
                        pdfPath = Path.Combine(pdfOutputFolder, Path.GetFileNameWithoutExtension(file.FileName) + ".pdf");
                    }
                }

                if (File.Exists(pdfPath))
                {
                    pdfFilePaths.Add(pdfPath);
                }
                else
                {
                    missingPdfFiles.Add(file.FileName);
                }
            }

            if (missingPdfFiles.Any())
            {
                var message = "以下のファイルに対応するPDFファイルが見つかりません:\n\n";
                message += string.Join("\n", missingPdfFiles);
                message += "\n\n不足しているファイルを自動的にPDF化してから結合しますか？";
                
                var result = MessageBox.Show(message, "PDFファイル不足", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Cancel)
                {
                    return;
                }
                else if (result == MessageBoxResult.Yes)
                {
                    // 不足しているファイルをPDF化する
                    await ConvertMissingFilesToPdf(missingPdfFiles, allFiles);
                    
                    // PDF化後に再度パスを取得
                    pdfFilePaths.Clear();
                    missingPdfFiles.Clear();
                    
                    foreach (var file in allFiles)
                    {
                        string pdfPath;
                        if (file.Extension.ToLower() == "pdf")
                        {
                            pdfPath = GetPdfPath(file.FilePath, pdfOutputFolder, baseFolderPath, includeSubfolders);
                        }
                        else
                        {
                            if (includeSubfolders)
                            {
                                var fileInfo = new FileInfo(file.FilePath);
                                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                                pdfPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(file.FileName) + ".pdf");
                            }
                            else
                            {
                                pdfPath = Path.Combine(pdfOutputFolder, Path.GetFileNameWithoutExtension(file.FileName) + ".pdf");
                            }
                        }

                        if (File.Exists(pdfPath))
                        {
                            pdfFilePaths.Add(pdfPath);
                        }
                        else
                        {
                            missingPdfFiles.Add(file.FileName);
                        }
                    }
                    
                    // まだ不足しているファイルがある場合
                    if (missingPdfFiles.Any())
                    {
                        var remainingMessage = "以下のファイルはPDF化に失敗しました。これらのファイルを除外して結合を続行しますか？\n\n";
                        remainingMessage += string.Join("\n", missingPdfFiles);
                        
                        var continueResult = MessageBox.Show(remainingMessage, "PDF化失敗", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                        if (continueResult == MessageBoxResult.No)
                        {
                            return;
                        }
                    }
                }
                // result == MessageBoxResult.No の場合は、不足ファイルを無視して結合を続行
            }

            // mergePDFフォルダの場所を決定（カスタムPDF保存パスを考慮）
            var mergeFolder = GetMergePdfFolderPath();
            
            if (!Directory.Exists(mergeFolder))
                Directory.CreateDirectory(mergeFolder);

            var mergeFileName = txtMergeFileName.Text;
            var addPageNumber = chkAddPageNumber.IsChecked == true;
            var addBookmarks = chkAddBookmarks.IsChecked == true;
            var groupByFolder = chkGroupByFolder.IsChecked == true;
            var addHeader = chkAddHeader.IsChecked == true;
            var addFooter = chkAddFooter.IsChecked == true;
            var headerText = txtHeaderText.Text ?? "";
            var footerText = txtFooterText.Text ?? "";
            var addHeaderFooter = addHeader || addFooter;
            var headerFooterText = addHeader ? headerText : footerText;
            
            // メイン画面の設定をプロジェクトに保存
            if (currentProject != null)
            {
                currentProject.AddHeader = addHeader;
                currentProject.AddFooter = addFooter;
                currentProject.HeaderText = headerText;
                currentProject.FooterText = footerText;
                currentProject.MergeFileName = txtMergeFileName.Text;
                currentProject.AddPageNumber = chkAddPageNumber.IsChecked == true;
                currentProject.AddBookmarks = chkAddBookmarks.IsChecked == true;
                currentProject.GroupByFolder = chkGroupByFolder.IsChecked == true;
            }
            var headerFooterFontSize = 10.0f;
            
            // フォントサイズをパース
            var txtHeaderFooterFontSize = FindName("txtHeaderFooterFontSize") as System.Windows.Controls.TextBox;
            if (txtHeaderFooterFontSize != null && !float.TryParse(txtHeaderFooterFontSize.Text, out headerFooterFontSize))
            {
                headerFooterFontSize = 10.0f;
            }

            // 新しい位置設定を取得（currentProjectから）
            // 位置設定を取得（プロジェクトから）
            var pageNumberPosition = currentProject?.PageNumberPosition ?? 0;
            var pageNumberOffsetX = currentProject?.PageNumberOffsetX ?? 20.0f;
            var pageNumberOffsetY = currentProject?.PageNumberOffsetY ?? 20.0f;
            var pageNumberFontSize = currentProject?.PageNumberFontSize ?? 10.0f;
            var headerPosition = currentProject?.HeaderPosition ?? 0;
            var headerOffsetX = currentProject?.HeaderOffsetX ?? 20.0f;
            var headerOffsetY = currentProject?.HeaderOffsetY ?? 20.0f;
            var headerFontSize = currentProject?.HeaderFontSize ?? 10.0f;
            var footerPosition = currentProject?.FooterPosition ?? 2;
            var footerOffsetX = currentProject?.FooterOffsetX ?? 20.0f;
            var footerOffsetY = currentProject?.FooterOffsetY ?? 20.0f;
            var footerFontSize = currentProject?.FooterFontSize ?? 10.0f;
            
            var timestamp = DateTime.Now.ToString("yyMMddHHmmss");
            var outputFileName = $"{mergeFileName}_{timestamp}.pdf";
            var outputPath = Path.Combine(mergeFolder, outputFileName);

            progressBar.Visibility = Visibility.Visible;
            progressBar.IsIndeterminate = true;
            txtStatus.Text = "PDF結合中...";

            bool mergeSuccess = false;
            
            await Task.Run(() =>
            {
                try
                {
                    if (addBookmarks && (includeSubfolders && groupByFolder))
                    {
                        // 高度なしおり機能を使用（フォルダ別グループ化）
                        PdfMergeService.MergePdfFilesWithAdvancedBookmarks(pdfFilePaths, outputPath, allFiles, addPageNumber, true, addHeaderFooter, headerFooterText, headerFooterFontSize,
                            pageNumberPosition, pageNumberOffsetX, pageNumberOffsetY, pageNumberFontSize,
                            headerPosition, headerOffsetX, headerOffsetY, headerFontSize,
                            footerPosition, footerOffsetX, footerOffsetY, footerFontSize,
                            addHeader, addFooter, headerText, footerText);
                    }
                    else if (addBookmarks)
                    {
                        // 基本的なしおり機能を使用
                        PdfMergeService.MergePdfFiles(pdfFilePaths, outputPath, addPageNumber, true, allFiles, addHeaderFooter, headerFooterText, headerFooterFontSize,
                            pageNumberPosition, pageNumberOffsetX, pageNumberOffsetY, pageNumberFontSize,
                            headerPosition, headerOffsetX, headerOffsetY, headerFontSize,
                            footerPosition, footerOffsetX, footerOffsetY, footerFontSize,
                            addHeader, addFooter, headerText, footerText);
                    }
                    else
                    {
                        // しおりなしで結合
                        PdfMergeService.MergePdfFiles(pdfFilePaths, outputPath, addPageNumber, false, null, addHeaderFooter, headerFooterText, headerFooterFontSize,
                            pageNumberPosition, pageNumberOffsetX, pageNumberOffsetY, pageNumberFontSize,
                            headerPosition, headerOffsetX, headerOffsetY, headerFontSize,
                            footerPosition, footerOffsetX, footerOffsetY, footerFontSize,
                            addHeader, addFooter, headerText, footerText);
                    }
                    mergeSuccess = true;
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show($"PDF結合エラー: {ex.Message}", "エラー",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    });
                }
            });

            progressBar.Visibility = Visibility.Collapsed;
            
            if (mergeSuccess)
            {
                if (currentProject != null)
                {
                    currentProject.LatestMergedPdfPath = outputPath;
                    SaveProjects();
                }

                UpdateLatestMergedPdfDisplay();
                
                var statusMessage = "PDF結合が完了しました";
                var statusParts = new List<string>();
                
                if (addBookmarks)
                {
                    statusParts.Add("しおり付き");
                }
                if (addHeaderFooter && !string.IsNullOrEmpty(headerFooterText))
                {
                    statusParts.Add("ヘッダ・フッタ付き");
                }
                if (addPageNumber)
                {
                    statusParts.Add("ページ番号付き");
                }
                
                if (statusParts.Any())
                {
                    statusMessage += $" ({string.Join(", ", statusParts)})";
                }
                
                txtStatus.Text = statusMessage;
                
                var result = MessageBox.Show("PDFを開きますか？", "PDF結合完了", 
                    MessageBoxButton.YesNo, MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    OpenFile(outputPath);
                }
                else
                {
                    try
                    {
                        Process.Start("explorer.exe", mergeFolder);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void UpdateLatestMergedPdfDisplay()
        {
            if (currentProject != null && !string.IsNullOrEmpty(currentProject.LatestMergedPdfPath))
            {
                if (File.Exists(currentProject.LatestMergedPdfPath))
                {
                    txtLatestMergedPDF.Text = Path.GetFileName(currentProject.LatestMergedPdfPath);
                    btnOpenLatestMergedPDF.IsEnabled = true;
                }
                else
                {
                    // ファイルが存在しない場合、mergePDFフォルダから最新のファイルを検索
                    var latestPdf = FindLatestMergedPdf();
                    if (!string.IsNullOrEmpty(latestPdf))
                    {
                        currentProject.LatestMergedPdfPath = latestPdf;
                        txtLatestMergedPDF.Text = Path.GetFileName(latestPdf);
                        btnOpenLatestMergedPDF.IsEnabled = true;
                        SaveProjects();
                    }
                    else
                    {
                        txtLatestMergedPDF.Text = "ファイルが見つかりません";
                        btnOpenLatestMergedPDF.IsEnabled = false;
                    }
                }
            }
            else
            {
                // プロジェクトに保存されたパスがない場合、mergePDFフォルダから最新のファイルを検索
                var latestPdf = FindLatestMergedPdf();
                if (!string.IsNullOrEmpty(latestPdf))
                {
                    if (currentProject != null)
                    {
                        currentProject.LatestMergedPdfPath = latestPdf;
                        SaveProjects();
                    }
                    txtLatestMergedPDF.Text = Path.GetFileName(latestPdf);
                    btnOpenLatestMergedPDF.IsEnabled = true;
                }
                else
                {
                    txtLatestMergedPDF.Text = "まだ結合されていません";
                    btnOpenLatestMergedPDF.IsEnabled = false;
                }
            }
        }

        /// <summary>
        /// mergePDFフォルダから最新のPDFファイルを検索
        /// </summary>
        /// <returns>最新のPDFファイルパス、見つからない場合はnull</returns>
        private string? FindLatestMergedPdf()
        {
            try
            {
                var mergeFolder = GetMergePdfFolderPath();
                if (!Directory.Exists(mergeFolder))
                    return null;

                var pdfFiles = Directory.GetFiles(mergeFolder, "*.pdf")
                    .Where(f => Path.GetFileName(f).Contains(txtMergeFileName.Text))
                    .OrderByDescending(f => File.GetCreationTime(f))
                    .ToList();

                return pdfFiles.FirstOrDefault();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"最新PDFファイル検索エラー: {ex.Message}");
                return null;
            }
        }

        private void BtnOpenLatestMergedPDF_Click(object sender, RoutedEventArgs e)
        {
            // まず最新のファイル状態を確認
            UpdateLatestMergedPdfDisplay();
            
            if (currentProject != null && !string.IsNullOrEmpty(currentProject.LatestMergedPdfPath))
            {
                if (File.Exists(currentProject.LatestMergedPdfPath))
                {
                    OpenFile(currentProject.LatestMergedPdfPath);
                }
                else
                {
                    MessageBox.Show("結合PDFファイルが見つかりません。", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    UpdateLatestMergedPdfDisplay();
                }
            }
            else
            {
                // 直接mergePDFフォルダから最新ファイルを検索して開く
                var latestPdf = FindLatestMergedPdf();
                if (!string.IsNullOrEmpty(latestPdf))
                {
                    OpenFile(latestPdf);
                }
                else
                {
                    MessageBox.Show("結合PDFファイルがありません。", "情報",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }
        #endregion

        #region フォルダ操作
        private void BtnOpenProjectFolder_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Tag is ProjectData project)
            {
                if (Directory.Exists(project.FolderPath))
                {
                    try
                    {
                        Process.Start("explorer.exe", project.FolderPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show("フォルダが見つかりません。", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void BtnOpenCurrentProjectFolder_Click(object sender, RoutedEventArgs e)
        {
            if (currentProject != null && Directory.Exists(currentProject.FolderPath))
            {
                try
                {
                    Process.Start("explorer.exe", currentProject.FolderPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else if (currentProject == null)
            {
                MessageBox.Show("現在のプロジェクトが選択されていません。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                MessageBox.Show("フォルダが見つかりません。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        #endregion

        #region イベントハンドラ
        private void TreeProjects_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is ProjectData selectedProject)
            {
                // プロジェクトが選択された場合の処理はここに追加
                // 現在は何もしない（ダブルクリックで切り替え）
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            SaveCurrentProjectState();
            SaveAppSettings();
            base.OnClosed(e);
        }
        #endregion

        #region ドラッグ&ドロップ処理
        private void Window_DragEnter(object sender, DragEventArgs e)
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

        private void Window_DragOver(object sender, DragEventArgs e)
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

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // フォルダかファイルかを判定
                    if (Directory.Exists(droppedPath))
                    {
                        SetFolderPath(droppedPath);
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            SetFolderPath(parentFolder);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 新しいドラッグ&ドロップエリアのDragEnter
        /// </summary>
        private void DropArea_DragEnter(object sender, DragEventArgs e)
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
        /// 新しいドラッグ&ドロップエリアのDragOver
        /// </summary>
        private void DropArea_DragOver(object sender, DragEventArgs e)
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
        /// 新しいドラッグ&ドロップエリアのDragLeave
        /// </summary>
        private void DropArea_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            if (sender is Border border)
            {
                border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250));
                border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 122, 204));
            }
        }

        /// <summary>
        /// 新しいドラッグ&ドロップエリアのDrop
        /// </summary>
        private void DropArea_Drop(object sender, DragEventArgs e)
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
                        SetFolderPath(droppedPath);
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            SetFolderPath(parentFolder);
                        }
                    }
                }
            }
        }

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

        private void TxtFolderPath_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            txtFolderPath.Background = System.Windows.Media.Brushes.White;
        }

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
                        SetFolderPath(droppedPath);
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            SetFolderPath(parentFolder);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// フォルダパスを設定する共通メソッド
        /// </summary>
        /// <param name="folderPath">設定するフォルダパス</param>
        private void SetFolderPath(string folderPath)
        {
            // フォルダパスのみを設定（ファイル名は含めない）
            selectedFolderPath = folderPath;
            txtFolderPath.Text = selectedFolderPath;
            
            // ヒントテキストを非表示にする
            if (txtDropHint != null)
            {
                txtDropHint.Visibility = string.IsNullOrEmpty(selectedFolderPath) ? Visibility.Visible : Visibility.Collapsed;
            }
            
            // PDFアウトプットフォルダもフォルダパスのみに設定
            if (currentProject != null && currentProject.UseCustomPdfPath && !string.IsNullOrEmpty(currentProject.CustomPdfPath))
            {
                pdfOutputFolder = currentProject.CustomPdfPath;
            }
            else
            {
                pdfOutputFolder = Path.Combine(selectedFolderPath, "PDF");
            }
            
            if (currentProject != null)
            {
                currentProject.FolderPath = selectedFolderPath;
                // PdfOutputFolderはプロパティで自動計算されるので直接設定しない
                SaveProjects();
            }
            
            txtStatus.Text = "フォルダがドラッグ&ドロップで選択されました";
            
            // 自動的にファイル読み込みを実行
            LoadFilesFromFolder();
        }
        #endregion

        #region ヘルパーメソッド
        /// <summary>
        /// 相対パスを取得
        /// </summary>
        /// <param name="basePath">基準パス</param>
        /// <param name="fullPath">完全パス</param>
        /// <returns>相対パス</returns>
        private string GetRelativePath(string basePath, string fullPath)
        {
            var baseUri = new Uri(basePath.EndsWith(Path.DirectorySeparatorChar.ToString()) ? basePath : basePath + Path.DirectorySeparatorChar);
            var fullUri = new Uri(fullPath);
            
            if (baseUri.Scheme != fullUri.Scheme)
            {
                return fullPath;
            }

            var relativeUri = baseUri.MakeRelativeUri(fullUri);
            var relativePath = Uri.UnescapeDataString(relativeUri.ToString());
            
            return relativePath.Replace('/', Path.DirectorySeparatorChar);
        }

        /// <summary>
        /// mergePDFフォルダのパスを取得
        /// </summary>
        /// <returns>mergePDFフォルダのパス</returns>
        private string GetMergePdfFolderPath()
        {
            if (currentProject != null)
            {
                return currentProject.MergePdfFolder;
            }
            else
            {
                // プロジェクトがない場合は従来通り
                return Path.Combine(selectedFolderPath, "mergePDF");
            }
        }

        /// <summary>
        /// 全ファイルのPDFステータスをリアルタイムで更新
        /// </summary>
        private void RefreshAllPdfStatus()
        {
            if (string.IsNullOrEmpty(selectedFolderPath) || string.IsNullOrEmpty(pdfOutputFolder))
                return;

            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;

            // UIスレッドで実行されていることを確認
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => RefreshAllPdfStatus());
                return;
            }

            foreach (var file in fileItems)
            {
                // PDFファイル自体はスキップ
                if (file.Extension.ToLower() == ".pdf")
                {
                    file.PdfStatus = "PDF";
                    continue;
                }

                // 対応するPDFファイルの存在を確認
                string pdfPath = GetPdfPath(file.FilePath, pdfOutputFolder, selectedFolderPath, includeSubfolders);
                bool pdfExists = File.Exists(pdfPath);
                
                var newStatus = pdfExists ? "変換済" : "未変換";
                
                // ステータスが変更された場合のみ更新
                if (file.PdfStatus != newStatus)
                {
                    file.PdfStatus = newStatus;
                }
                
                // 未変換の場合は自動的に選択状態にする（オプション）
                if (!pdfExists && !file.IsSelected)
                {
                    // 自動選択は行わない（ユーザーが手動で選択）
                }
            }
        }

        /// <summary>
        /// PDFファイルのパスを取得
        /// </summary>
        /// <param name="originalFilePath">元ファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="baseFolderPath">基準フォルダパス</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>PDFファイルパス</returns>
        private string GetPdfPath(string originalFilePath, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            if (includeSubfolders)
            {
                // サブフォルダ構造を考慮したパス
                var fileInfo = new FileInfo(originalFilePath);
                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                return Path.Combine(outputDir, Path.GetFileNameWithoutExtension(originalFilePath) + ".pdf");
            }
            else
            {
                // 従来通り、すべて同じフォルダに出力
                return Path.Combine(pdfOutputFolder, Path.GetFileNameWithoutExtension(originalFilePath) + ".pdf");
            }
        }

        /// <summary>
        /// 位置設定ボタンクリック
        /// </summary>
        private void BtnPositionSettings_Click(object sender, RoutedEventArgs e)
        {
            if (currentProject == null)
            {
                MessageBox.Show("プロジェクトが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dialog = new PositionSettingsDialog(currentProject);
            dialog.Owner = this;
            
            if (dialog.ShowDialog() == true)
            {
                // ダイアログの結果をプロジェクトに反映
                dialog.SaveToProject(currentProject);
                
                // プロジェクトを保存
                SaveProjects();
                
                MessageBox.Show("位置設定を保存しました。", "設定完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        /// <summary>
        /// ページ設定ボタンクリック
        /// </summary>
        private void BtnPageSettings_Click(object sender, RoutedEventArgs e)
        {
            if (currentProject == null)
            {
                MessageBox.Show("プロジェクトが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // メイン画面のヘッダ・フッタ設定をプロジェクトに反映
            currentProject.AddHeader = chkAddHeader.IsChecked == true;
            currentProject.AddFooter = chkAddFooter.IsChecked == true;
            currentProject.HeaderText = txtHeaderText.Text ?? "";
            currentProject.FooterText = txtFooterText.Text ?? "";

            var dialog = new PositionSettingsDialog(currentProject);
            dialog.Owner = this;
            
            if (dialog.ShowDialog() == true)
            {
                // ダイアログの結果をプロジェクトに反映
                dialog.SaveToProject(currentProject);
                
                // プロジェクトを保存
                SaveProjects();
                
                MessageBox.Show("ページ設定を保存しました。", "設定完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        #endregion

        #region インポート・エクスポート機能
        /// <summary>
        /// プロジェクトデータをエクスポート
        /// </summary>
        private void MenuExportProjects_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "JSON ファイル (*.json)|*.json",
                    Title = "プロジェクトデータをエクスポート",
                    FileName = $"Nico2PDF_Projects_{DateTime.Now:yyyyMMdd_HHmmss}.json"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    var projects = ProjectManager.LoadProjects();
                    var options = new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                    };

                    var jsonString = JsonSerializer.Serialize(projects, options);
                    File.WriteAllText(saveFileDialog.FileName, jsonString);

                    MessageBox.Show($"プロジェクトデータを正常にエクスポートしました。\n\nファイル: {saveFileDialog.FileName}", 
                        "エクスポート完了", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"プロジェクトデータのエクスポートに失敗しました。\n\nエラー: {ex.Message}", 
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// プロジェクトデータをインポート
        /// </summary>
        private void MenuImportProjects_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "JSON ファイル (*.json)|*.json",
                    Title = "プロジェクトデータをインポート"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    var result = MessageBox.Show(
                        "プロジェクトデータをインポートします。\n\n" +
                        "既存のプロジェクトと同じ名前のプロジェクトがある場合は上書きされます。\n" +
                        "続行しますか？",
                        "インポート確認",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question
                    );

                    if (result == MessageBoxResult.Yes)
                    {
                        var jsonString = File.ReadAllText(openFileDialog.FileName);
                        var importedProjects = JsonSerializer.Deserialize<List<ProjectData>>(jsonString);

                        if (importedProjects != null)
                        {
                            var existingProjects = ProjectManager.LoadProjects();
                            var importCount = 0;
                            var updateCount = 0;

                            foreach (var importedProject in importedProjects)
                            {
                                var existingProject = existingProjects.FirstOrDefault(p => p.Name == importedProject.Name);
                                if (existingProject != null)
                                {
                                    // 既存プロジェクトを更新（アクティブ状態は保持）
                                    var wasActive = existingProject.IsActive;
                                    existingProjects.Remove(existingProject);
                                    importedProject.IsActive = wasActive;
                                    existingProjects.Add(importedProject);
                                    updateCount++;
                                }
                                else
                                {
                                    // 新規プロジェクトを追加（非アクティブにする）
                                    importedProject.IsActive = false;
                                    existingProjects.Add(importedProject);
                                    importCount++;
                                }
                            }

                            ProjectManager.SaveProjects(existingProjects);
                            LoadProjects();

                            MessageBox.Show(
                                $"プロジェクトデータを正常にインポートしました。\n\n" +
                                $"新規追加: {importCount}件\n" +
                                $"更新: {updateCount}件",
                                "インポート完了",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information
                            );
                        }
                        else
                        {
                            MessageBox.Show("ファイルの形式が正しくありません。", "エラー", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"プロジェクトデータのインポートに失敗しました。\n\nエラー: {ex.Message}", 
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 不足しているPDFファイルを自動的に変換する
        /// </summary>
        /// <param name="missingFileNames">不足しているファイル名のリスト</param>
        /// <param name="allFiles">全ファイルのリスト</param>
        private async Task ConvertMissingFilesToPdf(List<string> missingFileNames, List<FileItem> allFiles)
        {
            var filesToConvert = allFiles.Where(f => missingFileNames.Contains(f.FileName)).ToList();
            
            if (!filesToConvert.Any())
                return;

            if (!Directory.Exists(pdfOutputFolder))
                Directory.CreateDirectory(pdfOutputFolder);

            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
            var baseFolderPath = selectedFolderPath;

            progressBar.Visibility = Visibility.Visible;
            progressBar.Maximum = filesToConvert.Count;
            progressBar.Value = 0;

            var convertedFiles = new List<FileItem>();

            await Task.Run(() =>
            {
                foreach (var file in filesToConvert)
                {
                    try
                    {
                        Dispatcher.Invoke(() =>
                        {
                            txtStatus.Text = $"PDF化中: {file.DisplayName}";
                        });

                        // サブフォルダを含む場合、出力ディレクトリを確保
                        if (includeSubfolders)
                        {
                            FileManagementService.EnsurePdfOutputDirectory(file.FilePath, pdfOutputFolder, baseFolderPath, includeSubfolders);
                        }

                        PdfConversionService.ConvertToPdf(file.FilePath, pdfOutputFolder, file.TargetPages, baseFolderPath, includeSubfolders);

                        // 変換成功したファイルをリストに追加
                        convertedFiles.Add(file);

                        Dispatcher.Invoke(() =>
                        {
                            // 個別ファイルのステータス更新（即座に反映）
                            UpdateIndividualPdfStatus(file, includeSubfolders);
                            progressBar.Value++;
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            txtStatus.Text = $"変換エラー: {file.DisplayName} - {ex.Message}";
                        });
                    }
                }
            });

            progressBar.Visibility = Visibility.Collapsed;
            txtStatus.Text = $"{convertedFiles.Count}個のファイルをPDF化しました";
        }

        #endregion
    }
}