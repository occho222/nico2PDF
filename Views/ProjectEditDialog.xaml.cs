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
    /// �v���W�F�N�g�ҏW�_�C�A���O
    /// </summary>
    public partial class ProjectEditDialog : Window
    {
        #region �v���p�e�B
        /// <summary>
        /// �v���W�F�N�g��
        /// </summary>
        public string ProjectName { get; set; } = "";

        /// <summary>
        /// �v���W�F�N�g�J�e�S��
        /// </summary>
        public string Category { get; set; } = "";

        /// <summary>
        /// �t�H���_�p�X
        /// </summary>
        public string FolderPath { get; set; } = "";

        /// <summary>
        /// �T�u�t�H���_���܂ނ��ǂ���
        /// </summary>
        public bool IncludeSubfolders { get; set; } = false;
        public int SubfolderDepth { get; set; } = 1;

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X���g�p���邩�ǂ���
        /// </summary>
        public bool UseCustomPdfPath { get; set; } = false;

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X
        /// </summary>
        public string CustomPdfPath { get; set; } = "";

        /// <summary>
        /// ���p�\�ȃJ�e�S�����X�g
        /// </summary>
        private List<string> availableCategories = new List<string>();
        #endregion

        #region �R���X�g���N�^
        public ProjectEditDialog()
        {
            InitializeComponent();
            LoadAvailableCategories();
        }
        #endregion

        #region ������
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
        /// ���p�\�ȃJ�e�S����ǂݍ���
        /// </summary>
        private void LoadAvailableCategories()
        {
            var allProjects = ProjectManager.LoadProjects();
            availableCategories = ProjectManager.GetAvailableCategories(allProjects);
            
            // �悭�g����J�e�S����ǉ�
            var defaultCategories = new List<string> { "�Ɩ�", "�l", "�J��", "����", "�A�[�J�C�u" };
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

        #region �C�x���g�n���h��
        /// <summary>
        /// �E�B���h�E�ǂݍ��ݎ�
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

            // �J�X�^��PDF�p�X�̗L��/������ݒ�
            UpdateCustomPdfPathEnabled();
            
            // �q���g�e�L�X�g�̕\������
            UpdateDropHints();
        }

        /// <summary>
        /// �h���b�O&�h���b�v�q���g�̕\�����X�V
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
        /// �t�H���_�I���{�^���N���b�N��
        /// </summary>
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "�v���W�F�N�g�t�H���_��I�����Ă�������";
                if (!string.IsNullOrEmpty(txtFolderPath.Text))
                {
                    dialog.SelectedPath = txtFolderPath.Text;
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtFolderPath.Text = dialog.SelectedPath;
                    
                    // �v���W�F�N�g������̏ꍇ�̓t�H���_����ݒ�
                    if (string.IsNullOrEmpty(txtProjectName.Text))
                    {
                        txtProjectName.Text = Path.GetFileName(dialog.SelectedPath);
                    }
                    
                    // �q���g�\�����X�V
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X�I���{�^���N���b�N��
        /// </summary>
        private void BtnSelectCustomPdfPath_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "PDF�ۑ��t�H���_��I�����Ă��������i�t�H���_�p�X�݂̂��ݒ肳��܂��j";
                if (!string.IsNullOrEmpty(txtCustomPdfPath.Text))
                {
                    // �����̃p�X���t�@�C�������܂ޏꍇ�́A�f�B���N�g���p�X�݂̂��擾
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
                        // �p�X�̐e�f�B���N�g�������݂��邩�`�F�b�N
                        var parentDir = Path.GetDirectoryName(existingPath);
                        if (!string.IsNullOrEmpty(parentDir) && Directory.Exists(parentDir))
                        {
                            dialog.SelectedPath = parentDir;
                        }
                    }
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // �t�H���_�p�X�݂̂�ݒ�i�t�@�C�����͊܂߂Ȃ��j
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
                    
                    // �q���g�\�����X�V
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// �T�u�t�H���_�ǂݍ��݃`�F�b�N��
        /// </summary>
        private void ChkIncludeSubfolders_Checked(object sender, RoutedEventArgs e)
        {
            // �T�u�t�H���_���܂ޏꍇ�A�J�X�^��PDF�p�X��K�{�ɂ���
            chkUseCustomPdfPath.IsChecked = true;
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// �T�u�t�H���_�ǂݍ��݃`�F�b�N������
        /// </summary>
        private void ChkIncludeSubfolders_Unchecked(object sender, RoutedEventArgs e)
        {
            // �T�u�t�H���_���܂܂Ȃ��ꍇ�͔C��
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X�g�p�`�F�b�N��
        /// </summary>
        private void ChkUseCustomPdfPath_Checked(object sender, RoutedEventArgs e)
        {
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X�g�p�`�F�b�N������
        /// </summary>
        private void ChkUseCustomPdfPath_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X���͗��̗L��/�������X�V
        /// </summary>
        private void UpdateCustomPdfPathEnabled()
        {
            var includeSubfolders = chkIncludeSubfolders.IsChecked == true;
            var useCustomPdfPath = chkUseCustomPdfPath.IsChecked == true;
            
            // �T�u�t�H���_���܂ޏꍇ�́A�J�X�^��PDF�p�X�������I�ɗL���ɂ���
            if (includeSubfolders)
            {
                chkUseCustomPdfPath.IsChecked = true;
                chkUseCustomPdfPath.IsEnabled = false; // �`�F�b�N�{�b�N�X�𖳌����i�K�{�j
                gridCustomPdfPath.IsEnabled = true;
            }
            else
            {
                chkUseCustomPdfPath.IsEnabled = true; // �`�F�b�N�{�b�N�X��L�����i�C�Ӂj
                gridCustomPdfPath.IsEnabled = useCustomPdfPath;
            }
        }

        /// <summary>
        /// OK�{�^���N���b�N��
        /// </summary>
        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtProjectName.Text))
            {
                MessageBox.Show("�v���W�F�N�g������͂��Ă��������B", "�G���[", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtFolderPath.Text))
            {
                MessageBox.Show("�t�H���_��I�����Ă��������B", "�G���[", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!Directory.Exists(txtFolderPath.Text))
            {
                MessageBox.Show("�I�����ꂽ�t�H���_�����݂��܂���B", "�G���[", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (chkIncludeSubfolders.IsChecked == true && chkUseCustomPdfPath.IsChecked != true)
            {
                MessageBox.Show("�T�u�t�H���_���܂ސݒ�̏ꍇ�A�J�X�^��PDF�ۑ��p�X�̐ݒ肪�K�{�ł��B", "�G���[", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (chkUseCustomPdfPath.IsChecked == true)
            {
                if (string.IsNullOrWhiteSpace(txtCustomPdfPath.Text))
                {
                    MessageBox.Show("�J�X�^��PDF�ۑ��p�X��I�����Ă��������B", "�G���[", 
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
                    var result = MessageBox.Show("�w�肳�ꂽPDF�ۑ��t�H���_�����݂��܂���B�쐬���܂����H", "�m�F", 
                        MessageBoxButton.YesNo, MessageBoxImage.Question);
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        try
                        {
                            Directory.CreateDirectory(txtCustomPdfPath.Text);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"�t�H���_�̍쐬�Ɏ��s���܂���: {ex.Message}", "�G���[", 
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
        /// �L�����Z���{�^���N���b�N��
        /// </summary>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        #endregion

        #region �h���b�O&�h���b�v����
        /// <summary>
        /// �t�H���_�p�X�p�h���b�O&�h���b�v�G���A��DragEnter
        /// </summary>
        private void FolderDropArea_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // �h���b�O�I�[�o�[���̎��o�I�t�B�[�h�o�b�N
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
        /// �t�H���_�p�X�p�h���b�O&�h���b�v�G���A��DragOver
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
        /// �t�H���_�p�X�p�h���b�O&�h���b�v�G���A��DragLeave
        /// </summary>
        private void FolderDropArea_DragLeave(object sender, DragEventArgs e)
        {
            // �h���b�O���[�u���̎��o�I�t�B�[�h�o�b�N�����ɖ߂�
            if (sender is Border border)
            {
                border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250));
                border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 122, 204));
            }
        }

        /// <summary>
        /// �t�H���_�p�X�p�h���b�O&�h���b�v�G���A��Drop
        /// </summary>
        private void FolderDropArea_Drop(object sender, DragEventArgs e)
        {
            // �h���b�O�I�[�o�[���̎��o�I�t�B�[�h�o�b�N�����ɖ߂�
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
                    
                    // �t�H���_���t�@�C�����𔻒�
                    if (Directory.Exists(droppedPath))
                    {
                        txtFolderPath.Text = droppedPath;
                        
                        // �v���W�F�N�g������̏ꍇ�̓t�H���_����ݒ�
                        if (string.IsNullOrEmpty(txtProjectName.Text))
                        {
                            txtProjectName.Text = Path.GetFileName(droppedPath);
                        }
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // �t�@�C���̏ꍇ�͐e�t�H���_���g�p
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            txtFolderPath.Text = parentFolder;
                            
                            // �v���W�F�N�g������̏ꍇ�̓t�H���_����ݒ�
                            if (string.IsNullOrEmpty(txtProjectName.Text))
                            {
                                txtProjectName.Text = Path.GetFileName(parentFolder);
                            }
                        }
                    }
                    
                    // �q���g�\�����X�V
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// PDF�ۑ��p�X�p�h���b�O&�h���b�v�G���A��DragEnter
        /// </summary>
        private void PdfDropArea_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // �h���b�O�I�[�o�[���̎��o�I�t�B�[�h�o�b�N
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
        /// PDF�ۑ��p�X�p�h���b�O&�h���b�v�G���A��DragOver
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
        /// PDF�ۑ��p�X�p�h���b�O&�h���b�v�G���A��DragLeave
        /// </summary>
        private void PdfDropArea_DragLeave(object sender, DragEventArgs e)
        {
            // �h���b�O���[�u���̎��o�I�t�B�[�h�o�b�N�����ɖ߂�
            if (sender is Border border)
            {
                border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250));
                border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 122, 204));
            }
        }

        /// <summary>
        /// PDF�ۑ��p�X�p�h���b�O&�h���b�v�G���A��Drop
        /// </summary>
        private void PdfDropArea_Drop(object sender, DragEventArgs e)
        {
            // �h���b�O�I�[�o�[���̎��o�I�t�B�[�h�o�b�N�����ɖ߂�
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
                    
                    // �t�H���_���t�@�C�����𔻒�
                    if (Directory.Exists(droppedPath))
                    {
                        // �t�H���_�p�X�݂̂�ݒ�i�t�@�C�����͊܂߂Ȃ��j
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
                        // �t�@�C���̏ꍇ�͐e�t�H���_���g�p
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
                    
                    // �q���g�\�����X�V
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// �t�H���_�p�X�e�L�X�g�{�b�N�X�̃h���b�O�G���^�[�i���ł̌݊����ێ��j
        /// </summary>
        private void TxtFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // �h���b�O�I�[�o�[���̎��o�I�t�B�[�h�o�b�N
                txtFolderPath.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightBlue);
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// �t�H���_�p�X�e�L�X�g�{�b�N�X�̃h���b�O�I�[�o�[�i���ł̌݊����ێ��j
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
        /// �t�H���_�p�X�e�L�X�g�{�b�N�X�̃h���b�O���[�u�i���ł̌݊����ێ��j
        /// </summary>
        private void TxtFolderPath_DragLeave(object sender, DragEventArgs e)
        {
            // �h���b�O���[�u���̎��o�I�t�B�[�h�o�b�N�����ɖ߂�
            txtFolderPath.Background = System.Windows.Media.Brushes.White;
        }

        /// <summary>
        /// �t�H���_�p�X�e�L�X�g�{�b�N�X�̃h���b�v�i���ł̌݊����ێ��j
        /// </summary>
        private void TxtFolderPath_Drop(object sender, DragEventArgs e)
        {
            // �h���b�O�I�[�o�[���̎��o�I�t�B�[�h�o�b�N�����ɖ߂�
            txtFolderPath.Background = System.Windows.Media.Brushes.White;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // �t�H���_���t�@�C�����𔻒�
                    if (Directory.Exists(droppedPath))
                    {
                        txtFolderPath.Text = droppedPath;
                        
                        // �v���W�F�N�g������̏ꍇ�̓t�H���_����ݒ�
                        if (string.IsNullOrEmpty(txtProjectName.Text))
                        {
                            txtProjectName.Text = Path.GetFileName(droppedPath);
                        }
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // �t�@�C���̏ꍇ�͐e�t�H���_���g�p
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            txtFolderPath.Text = parentFolder;
                            
                            // �v���W�F�N�g������̏ꍇ�̓t�H���_����ݒ�
                            if (string.IsNullOrEmpty(txtProjectName.Text))
                            {
                                txtProjectName.Text = Path.GetFileName(parentFolder);
                            }
                        }
                    }
                    
                    // �q���g�\�����X�V
                    UpdateDropHints();
                }
            }
        }

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X�e�L�X�g�{�b�N�X�̃h���b�O�G���^�[�i���ł̌݊����ێ��j
        /// </summary>
        private void TxtCustomPdfPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // �h���b�O�I�[�o�[���̎��o�I�t�B�[�h�o�b�N
                txtCustomPdfPath.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightBlue);
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X�e�L�X�g�{�b�N�X�̃h���b�O�I�[�o�[�i���ł̌݊����ێ��j
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
        /// �J�X�^��PDF�ۑ��p�X�e�L�X�g�{�b�N�X�̃h���b�O���[�u�i���ł̌݊����ێ��j
        /// </summary>
        private void TxtCustomPdfPath_DragLeave(object sender, DragEventArgs e)
        {
            // �h���b�O���[�u���̎��o�I�t�B�[�h�o�b�N�����ɖ߂�
            txtCustomPdfPath.Background = System.Windows.Media.Brushes.White;
        }

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X�e�L�X�g�{�b�N�X�̃h���b�v�i���ł̌݊����ێ��j
        /// </summary>
        private void TxtCustomPdfPath_Drop(object sender, DragEventArgs e)
        {
            // �h���b�O�I�[�o�[���̎��o�I�t�B�[�h�o�b�N�����ɖ߂�
            txtCustomPdfPath.Background = System.Windows.Media.Brushes.White;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // �t�H���_���t�@�C�����𔻒�
                    if (Directory.Exists(droppedPath))
                    {
                        // �t�H���_�p�X�݂̂�ݒ�i�t�@�C�����͊܂߂Ȃ��j
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
                        // �t�@�C���̏ꍇ�͐e�t�H���_���g�p
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
                    
                    // �q���g�\�����X�V
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