using System;
using System.Windows;
using System.Windows.Controls;
using Nico2PDF.Models;

namespace Nico2PDF.Views
{
    /// <summary>
    /// 位置設定ダイアログ
    /// </summary>
    public partial class PositionSettingsDialog : Window
    {
        /// <summary>
        /// ダイアログの結果（プロパティは削除し、WPFのDialogResultを使用）
        /// </summary>

        /// <summary>
        /// ページ番号のフォントサイズ
        /// </summary>
        public float PageNumberFontSize 
        { 
            get 
            { 
                if (float.TryParse(txtPageNumberFontSize.Text, out float value))
                    return value;
                return 10.0f;
            } 
        }

        /// <summary>
        /// ページ振りの位置（0:右上, 1:右下, 2:左上, 3:左下）
        /// </summary>
        public int PageNumberPosition => cmbPageNumberPosition.SelectedIndex;

        /// <summary>
        /// ページ振りのX軸オフセット
        /// </summary>
        public float PageNumberOffsetX 
        { 
            get 
            { 
                if (float.TryParse(txtPageNumberOffsetX.Text, out float value))
                    return value;
                return 20.0f;
            } 
        }

        /// <summary>
        /// ページ振りのY軸オフセット
        /// </summary>
        public float PageNumberOffsetY 
        { 
            get 
            { 
                if (float.TryParse(txtPageNumberOffsetY.Text, out float value))
                    return value;
                return 20.0f;
            } 
        }

        /// <summary>
        /// ヘッダの位置（0:左, 1:中央, 2:右）
        /// </summary>
        public int HeaderPosition => cmbHeaderPosition.SelectedIndex;

        /// <summary>
        /// ヘッダのX軸オフセット
        /// </summary>
        public float HeaderOffsetX 
        { 
            get 
            { 
                if (float.TryParse(txtHeaderOffsetX.Text, out float value))
                    return value;
                return 20.0f;
            } 
        }

        /// <summary>
        /// ヘッダのY軸オフセット
        /// </summary>
        public float HeaderOffsetY 
        { 
            get 
            { 
                if (float.TryParse(txtHeaderOffsetY.Text, out float value))
                    return value;
                return 20.0f;
            } 
        }

        /// <summary>
        /// ヘッダのフォントサイズ
        /// </summary>
        public float HeaderFontSize 
        { 
            get 
            { 
                if (float.TryParse(txtHeaderFontSize.Text, out float value))
                    return value;
                return 10.0f;
            } 
        }

        /// <summary>
        /// フッタの位置（0:左, 1:中央, 2:右）
        /// </summary>
        public int FooterPosition => cmbFooterPosition.SelectedIndex;

        /// <summary>
        /// フッタのX軸オフセット
        /// </summary>
        public float FooterOffsetX 
        { 
            get 
            { 
                if (float.TryParse(txtFooterOffsetX.Text, out float value))
                    return value;
                return 20.0f;
            } 
        }

        /// <summary>
        /// フッタのY軸オフセット
        /// </summary>
        public float FooterOffsetY 
        { 
            get 
            { 
                if (float.TryParse(txtFooterOffsetY.Text, out float value))
                    return value;
                return 20.0f;
            } 
        }

        /// <summary>
        /// フッタのフォントサイズ
        /// </summary>
        public float FooterFontSize 
        { 
            get 
            { 
                if (float.TryParse(txtFooterFontSize.Text, out float value))
                    return value;
                return 10.0f;
            } 
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public PositionSettingsDialog()
        {
            InitializeComponent();
            InitializeDefaults();
        }

        /// <summary>
        /// プロジェクトデータから初期化するコンストラクタ
        /// </summary>
        /// <param name="project">プロジェクトデータ</param>
        public PositionSettingsDialog(ProjectData project) : this()
        {
            if (project != null)
            {
                LoadFromProject(project);
            }
        }

        /// <summary>
        /// デフォルト値を設定
        /// </summary>
        private void InitializeDefaults()
        {
            cmbPageNumberPosition.SelectedIndex = 0; // 右上
            txtPageNumberOffsetX.Text = "20";
            txtPageNumberOffsetY.Text = "20";
            txtPageNumberFontSize.Text = "10";
            
            cmbHeaderPosition.SelectedIndex = 0; // 左
            txtHeaderOffsetX.Text = "20";
            txtHeaderOffsetY.Text = "20";
            txtHeaderFontSize.Text = "10";
            
            cmbFooterPosition.SelectedIndex = 2; // 右
            txtFooterOffsetX.Text = "20";
            txtFooterOffsetY.Text = "20";
            txtFooterFontSize.Text = "10";
        }

        /// <summary>
        /// プロジェクトデータから設定を読み込み
        /// </summary>
        /// <param name="project">プロジェクトデータ</param>
        private void LoadFromProject(ProjectData project)
        {
            cmbPageNumberPosition.SelectedIndex = project.PageNumberPosition;
            txtPageNumberOffsetX.Text = project.PageNumberOffsetX.ToString("0.0");
            txtPageNumberOffsetY.Text = project.PageNumberOffsetY.ToString("0.0");
            
            cmbHeaderPosition.SelectedIndex = project.HeaderPosition;
            txtHeaderOffsetX.Text = project.HeaderOffsetX.ToString("0.0");
            txtHeaderOffsetY.Text = project.HeaderOffsetY.ToString("0.0");
            
            cmbFooterPosition.SelectedIndex = project.FooterPosition;
            txtFooterOffsetX.Text = project.FooterOffsetX.ToString("0.0");
            txtFooterOffsetY.Text = project.FooterOffsetY.ToString("0.0");
        }

        /// <summary>
        /// プロジェクトデータに設定を保存
        /// </summary>
        /// <param name="project">プロジェクトデータ</param>
        public void SaveToProject(ProjectData project)
        {
            if (project != null)
            {
                project.PageNumberPosition = PageNumberPosition;
                project.PageNumberOffsetX = PageNumberOffsetX;
                project.PageNumberOffsetY = PageNumberOffsetY;
                
                project.HeaderPosition = HeaderPosition;
                project.HeaderOffsetX = HeaderOffsetX;
                project.HeaderOffsetY = HeaderOffsetY;
                
                project.FooterPosition = FooterPosition;
                project.FooterOffsetX = FooterOffsetX;
                project.FooterOffsetY = FooterOffsetY;
            }
        }

        /// <summary>
        /// OKボタンクリック
        /// </summary>
        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        /// <summary>
        /// キャンセルボタンクリック
        /// </summary>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }

    }
}