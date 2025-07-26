using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using Nico2PDF.Models;

namespace Nico2PDF.Views
{
    /// <summary>
    /// 一括リネームダイアログ
    /// </summary>
    public partial class BatchRenameDialog : Window, INotifyPropertyChanged
    {
        public BatchRenameDialog()
        {
            InitializeComponent();
            DataContext = this;
            RenameItems = new ObservableCollection<BatchRenameItem>();
            
            // 日本語テキストを設定
            Title = "一括リネーム";
            txtTitle.Text = "一括リネーム";
            btnHelp.Content = "ヘルプ";
            btnHelp.ToolTip = "一括リネーム方法の詳細説明を表示";
            btnOK.Content = "変更";
            btnCancel.Content = "キャンセル";
            
            // ボタンの日本語設定
            btnResetAll.Content = "全てリセット";
            btnResetAll.ToolTip = "全てのファイル名を元に戻します";
            btnAddPrefix.Content = "前置詞追加";
            btnAddPrefix.ToolTip = "全てのファイル名に前置詞を追加";
            btnAddSuffix.Content = "後置詞追加";
            btnAddSuffix.ToolTip = "全てのファイル名に後置詞を追加";
            
            // カラムヘッダーの日本語設定
            colNo.Header = "No";
            colFolder.Header = "フォルダ";
            colCurrentName.Header = "現在のファイル名";
            colNewName.Header = "新しいファイル名";
            colPreview.Header = "プレビュー";
            colExtension.Header = "拡張子";
            colStatus.Header = "状態";
            
            // 警告メッセージ
            txtWarningTitle.Text = "注意事項";
            txtWarning1.Text = "・ファイル名を変更すると、元のファイル名に戻すことはできません";
            txtWarning2.Text = "・他のアプリケーションで使用中のファイルは変更できません";
            txtWarning3.Text = "・変更後はPDFファイルの再生成が必要になる場合があります";

            // データテンプレート内のTextBlockにも日本語設定
            Loaded += (s, e) => 
            {
                // ChangedCountのテキストをバインディング
                txtChangedCount.SetBinding(System.Windows.Controls.TextBlock.TextProperty, 
                    new System.Windows.Data.Binding("ChangedFilesCount") 
                    { 
                        StringFormat = "変更対象: {0} 件",
                        Source = this
                    });
                    
                // データテンプレートの中にある TextBlock を見つけて設定
                SetDataTemplateTexts();
            };
        }

        /// <summary>
        /// 対象ファイルリスト
        /// </summary>
        public List<FileItem> TargetFiles { get; set; } = new List<FileItem>();

        /// <summary>
        /// 一括リネームアイテム
        /// </summary>
        public ObservableCollection<BatchRenameItem> RenameItems { get; set; }

        /// <summary>
        /// 対象ファイル情報
        /// </summary>
        public string TargetFilesInfo => $"対象ファイル数: {TargetFiles.Count}個";

        /// <summary>
        /// 変更されたファイル数
        /// </summary>
        public int ChangedFilesCount => RenameItems.Count(item => item.IsChanged && !item.HasError);

        /// <summary>
        /// ステータスメッセージ
        /// </summary>
        public string StatusMessage
        {
            get
            {
                var errorCount = RenameItems.Count(item => item.HasError);
                if (errorCount > 0)
                    return $"エラー: {errorCount} 件";
                
                var changedCount = ChangedFilesCount;
                if (changedCount > 0)
                    return $"変更準備完了: {changedCount} 件";
                
                return "変更するファイル名を編集してください";
            }
        }

        /// <summary>
        /// 入力値が有効かどうか
        /// </summary>
        public bool IsValid => RenameItems.Any(item => item.IsChanged) && !RenameItems.Any(item => item.HasError);

        /// <summary>
        /// 初期化処理
        /// </summary>
        public void Initialize()
        {
            RenameItems.Clear();
            
            foreach (var fileItem in TargetFiles)
            {
                var renameItem = new BatchRenameItem { OriginalItem = fileItem };
                renameItem.PropertyChanged += RenameItem_PropertyChanged;
                RenameItems.Add(renameItem);
            }
            
            UpdateProperties();
            
            // データテンプレート内のTextBlockにも日本語設定
            Loaded += (s, e) => 
            {
                // ChangedCountのテキストをバインディング
                txtChangedCount.SetBinding(System.Windows.Controls.TextBlock.TextProperty, 
                    new System.Windows.Data.Binding("ChangedFilesCount") 
                    { 
                        StringFormat = "変更対象: {0} 件",
                        Source = this
                    });
            };
        }

        /// <summary>
        /// リネームアイテムのプロパティ変更イベント
        /// </summary>
        private void RenameItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(BatchRenameItem.IsChanged) || 
                e.PropertyName == nameof(BatchRenameItem.HasError))
            {
                UpdateProperties();
            }
        }

        /// <summary>
        /// プロパティを更新
        /// </summary>
        private void UpdateProperties()
        {
            OnPropertyChanged(nameof(ChangedFilesCount));
            OnPropertyChanged(nameof(StatusMessage));
            OnPropertyChanged(nameof(IsValid));
        }

        /// <summary>
        /// 全てリセットボタンクリック
        /// </summary>
        private void BtnResetAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in RenameItems)
            {
                item.NewFileName = item.CurrentFileName;
            }
        }

        /// <summary>
        /// 前置詞追加ボタンクリック
        /// </summary>
        private void BtnAddPrefix_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new PrefixSuffixDialog("前置詞を入力してください：", "前置詞追加");
            if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(dialog.InputText))
            {
                foreach (var item in RenameItems)
                {
                    item.NewFileName = dialog.InputText + item.CurrentFileName;
                }
            }
        }

        /// <summary>
        /// 後置詞追加ボタンクリック
        /// </summary>
        private void BtnAddSuffix_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new PrefixSuffixDialog("後置詞を入力してください：", "後置詞追加");
            if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(dialog.InputText))
            {
                foreach (var item in RenameItems)
                {
                    item.NewFileName = item.CurrentFileName + dialog.InputText;
                }
            }
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
            var helpMessage = @"?? 一括リネーム機能の使い方

【基本的な使い方】
1. ファイル一覧の「新しいファイル名」列で、各ファイルの名前を手動で編集してください
2. プレビュー列で変更後のファイル名を確認してください
3. 「変更」ボタンで変更を適用してください

【便利な機能】

?? 一括操作ツール
・「全てリセット」：全てのファイル名を元に戻します
・「前置詞追加」：全てのファイル名の前に文字を追加します
・「後置詞追加」：全てのファイル名の後に文字を追加します

?? 表示について
・変更されたファイルは行が青色になります
・エラーがある場合は赤色で表示されます
・状態列で各ファイルの変更状況を確認できます

【ファイル名変更について】
・Windowsエクスプローラーでのファイル名も変更されます
・変更後は元に戻すことができません
・ファイルが他のアプリケーションで使用中の場合は変更できません

【ファイル名の制限】
・使用できない文字: \ / : * ? "" < > |
・予約語は使用できません（CON, PRN, AUX等）
・既存のファイルと同じ名前は使用できません
・200文字以内で入力してください

【注意事項】
・拡張子は自動的に付加されるため、入力不要です
・変更前にファイルのバックアップを推奨します
・重要なファイルは慎重に操作してください";

            System.Windows.MessageBox.Show(helpMessage, "一括リネーム機能のヘルプ", 
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// データテンプレート内のテキストを設定
        /// </summary>
        private void SetDataTemplateTexts()
        {
            // DataGridの行を巡回してTextBlockを設定
            for (int i = 0; i < dgFiles.Items.Count; i++)
            {
                var row = dgFiles.ItemContainerGenerator.ContainerFromIndex(i) as System.Windows.Controls.DataGridRow;
                if (row != null)
                {
                    var changedTextBlock = FindVisualChild<System.Windows.Controls.TextBlock>(row, "txtChanged");
                    var errorTextBlock = FindVisualChild<System.Windows.Controls.TextBlock>(row, "txtError");
                    
                    if (changedTextBlock != null)
                        changedTextBlock.Text = "?変更";
                    if (errorTextBlock != null)
                        errorTextBlock.Text = "?エラー";
                }
            }
        }
        
        /// <summary>
        /// 子要素を名前で検索
        /// </summary>
        private T FindVisualChild<T>(System.Windows.DependencyObject parent, string name) where T : System.Windows.DependencyObject
        {
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T && ((System.Windows.FrameworkElement)child).Name == name)
                {
                    return (T)child;
                }
                
                var childOfChild = FindVisualChild<T>(child, name);
                if (childOfChild != null)
                    return childOfChild;
            }
            return null;
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
    /// 前置詞・後置詞入力ダイアログ
    /// </summary>
    public class PrefixSuffixDialog : Window
    {
        private readonly System.Windows.Controls.TextBox textBox;

        public string InputText => textBox.Text;

        public PrefixSuffixDialog(string message, string title)
        {
            Title = title;
            Width = 350;
            Height = 150;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;

            var grid = new System.Windows.Controls.Grid();
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });

            var label = new System.Windows.Controls.Label { Content = message, Margin = new Thickness(10) };
            System.Windows.Controls.Grid.SetRow(label, 0);
            grid.Children.Add(label);

            textBox = new System.Windows.Controls.TextBox { Margin = new Thickness(10, 0, 10, 10) };
            System.Windows.Controls.Grid.SetRow(textBox, 1);
            grid.Children.Add(textBox);

            var buttonPanel = new System.Windows.Controls.StackPanel 
            { 
                Orientation = System.Windows.Controls.Orientation.Horizontal, 
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                Margin = new Thickness(10)
            };

            var okButton = new System.Windows.Controls.Button 
            { 
                Content = "OK", 
                Width = 75, 
                Height = 25, 
                Margin = new Thickness(0, 0, 5, 0),
                IsDefault = true
            };
            okButton.Click += (s, e) => { DialogResult = true; Close(); };

            var cancelButton = new System.Windows.Controls.Button 
            { 
                Content = "キャンセル", 
                Width = 75, 
                Height = 25,
                IsCancel = true
            };
            cancelButton.Click += (s, e) => { DialogResult = false; Close(); };

            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);

            System.Windows.Controls.Grid.SetRow(buttonPanel, 2);
            grid.Children.Add(buttonPanel);

            Content = grid;

            Loaded += (s, e) => textBox.Focus();
        }
    }
}