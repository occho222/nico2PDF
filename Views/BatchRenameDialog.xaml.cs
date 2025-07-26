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
    /// �ꊇ���l�[���_�C�A���O
    /// </summary>
    public partial class BatchRenameDialog : Window, INotifyPropertyChanged
    {
        public BatchRenameDialog()
        {
            InitializeComponent();
            DataContext = this;
            RenameItems = new ObservableCollection<BatchRenameItem>();
            
            // ���{��e�L�X�g��ݒ�
            Title = "�ꊇ���l�[��";
            txtTitle.Text = "�ꊇ���l�[��";
            btnHelp.Content = "�w���v";
            btnHelp.ToolTip = "�ꊇ���l�[�����@�̏ڍא�����\��";
            btnOK.Content = "�ύX";
            btnCancel.Content = "�L�����Z��";
            
            // �{�^���̓��{��ݒ�
            btnResetAll.Content = "�S�ă��Z�b�g";
            btnResetAll.ToolTip = "�S�Ẵt�@�C���������ɖ߂��܂�";
            btnAddPrefix.Content = "�O�u���ǉ�";
            btnAddPrefix.ToolTip = "�S�Ẵt�@�C�����ɑO�u����ǉ�";
            btnAddSuffix.Content = "��u���ǉ�";
            btnAddSuffix.ToolTip = "�S�Ẵt�@�C�����Ɍ�u����ǉ�";
            
            // �J�����w�b�_�[�̓��{��ݒ�
            colNo.Header = "No";
            colFolder.Header = "�t�H���_";
            colCurrentName.Header = "���݂̃t�@�C����";
            colNewName.Header = "�V�����t�@�C����";
            colPreview.Header = "�v���r���[";
            colExtension.Header = "�g���q";
            colStatus.Header = "���";
            
            // �x�����b�Z�[�W
            txtWarningTitle.Text = "���ӎ���";
            txtWarning1.Text = "�E�t�@�C������ύX����ƁA���̃t�@�C�����ɖ߂����Ƃ͂ł��܂���";
            txtWarning2.Text = "�E���̃A�v���P�[�V�����Ŏg�p���̃t�@�C���͕ύX�ł��܂���";
            txtWarning3.Text = "�E�ύX���PDF�t�@�C���̍Đ������K�v�ɂȂ�ꍇ������܂�";

            // �f�[�^�e���v���[�g����TextBlock�ɂ����{��ݒ�
            Loaded += (s, e) => 
            {
                // ChangedCount�̃e�L�X�g���o�C���f�B���O
                txtChangedCount.SetBinding(System.Windows.Controls.TextBlock.TextProperty, 
                    new System.Windows.Data.Binding("ChangedFilesCount") 
                    { 
                        StringFormat = "�ύX�Ώ�: {0} ��",
                        Source = this
                    });
                    
                // �f�[�^�e���v���[�g�̒��ɂ��� TextBlock �������Đݒ�
                SetDataTemplateTexts();
            };
        }

        /// <summary>
        /// �Ώۃt�@�C�����X�g
        /// </summary>
        public List<FileItem> TargetFiles { get; set; } = new List<FileItem>();

        /// <summary>
        /// �ꊇ���l�[���A�C�e��
        /// </summary>
        public ObservableCollection<BatchRenameItem> RenameItems { get; set; }

        /// <summary>
        /// �Ώۃt�@�C�����
        /// </summary>
        public string TargetFilesInfo => $"�Ώۃt�@�C����: {TargetFiles.Count}��";

        /// <summary>
        /// �ύX���ꂽ�t�@�C����
        /// </summary>
        public int ChangedFilesCount => RenameItems.Count(item => item.IsChanged && !item.HasError);

        /// <summary>
        /// �X�e�[�^�X���b�Z�[�W
        /// </summary>
        public string StatusMessage
        {
            get
            {
                var errorCount = RenameItems.Count(item => item.HasError);
                if (errorCount > 0)
                    return $"�G���[: {errorCount} ��";
                
                var changedCount = ChangedFilesCount;
                if (changedCount > 0)
                    return $"�ύX��������: {changedCount} ��";
                
                return "�ύX����t�@�C������ҏW���Ă�������";
            }
        }

        /// <summary>
        /// ���͒l���L�����ǂ���
        /// </summary>
        public bool IsValid => RenameItems.Any(item => item.IsChanged) && !RenameItems.Any(item => item.HasError);

        /// <summary>
        /// ����������
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
            
            // �f�[�^�e���v���[�g����TextBlock�ɂ����{��ݒ�
            Loaded += (s, e) => 
            {
                // ChangedCount�̃e�L�X�g���o�C���f�B���O
                txtChangedCount.SetBinding(System.Windows.Controls.TextBlock.TextProperty, 
                    new System.Windows.Data.Binding("ChangedFilesCount") 
                    { 
                        StringFormat = "�ύX�Ώ�: {0} ��",
                        Source = this
                    });
            };
        }

        /// <summary>
        /// ���l�[���A�C�e���̃v���p�e�B�ύX�C�x���g
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
        /// �v���p�e�B���X�V
        /// </summary>
        private void UpdateProperties()
        {
            OnPropertyChanged(nameof(ChangedFilesCount));
            OnPropertyChanged(nameof(StatusMessage));
            OnPropertyChanged(nameof(IsValid));
        }

        /// <summary>
        /// �S�ă��Z�b�g�{�^���N���b�N
        /// </summary>
        private void BtnResetAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in RenameItems)
            {
                item.NewFileName = item.CurrentFileName;
            }
        }

        /// <summary>
        /// �O�u���ǉ��{�^���N���b�N
        /// </summary>
        private void BtnAddPrefix_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new PrefixSuffixDialog("�O�u������͂��Ă��������F", "�O�u���ǉ�");
            if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(dialog.InputText))
            {
                foreach (var item in RenameItems)
                {
                    item.NewFileName = dialog.InputText + item.CurrentFileName;
                }
            }
        }

        /// <summary>
        /// ��u���ǉ��{�^���N���b�N
        /// </summary>
        private void BtnAddSuffix_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new PrefixSuffixDialog("��u������͂��Ă��������F", "��u���ǉ�");
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
            var helpMessage = @"?? �ꊇ���l�[���@�\�̎g����

�y��{�I�Ȏg�����z
1. �t�@�C���ꗗ�́u�V�����t�@�C�����v��ŁA�e�t�@�C���̖��O���蓮�ŕҏW���Ă�������
2. �v���r���[��ŕύX��̃t�@�C�������m�F���Ă�������
3. �u�ύX�v�{�^���ŕύX��K�p���Ă�������

�y�֗��ȋ@�\�z

?? �ꊇ����c�[��
�E�u�S�ă��Z�b�g�v�F�S�Ẵt�@�C���������ɖ߂��܂�
�E�u�O�u���ǉ��v�F�S�Ẵt�@�C�����̑O�ɕ�����ǉ����܂�
�E�u��u���ǉ��v�F�S�Ẵt�@�C�����̌�ɕ�����ǉ����܂�

?? �\���ɂ���
�E�ύX���ꂽ�t�@�C���͍s���F�ɂȂ�܂�
�E�G���[������ꍇ�͐ԐF�ŕ\������܂�
�E��ԗ�Ŋe�t�@�C���̕ύX�󋵂��m�F�ł��܂�

�y�t�@�C�����ύX�ɂ��āz
�EWindows�G�N�X�v���[���[�ł̃t�@�C�������ύX����܂�
�E�ύX��͌��ɖ߂����Ƃ��ł��܂���
�E�t�@�C�������̃A�v���P�[�V�����Ŏg�p���̏ꍇ�͕ύX�ł��܂���

�y�t�@�C�����̐����z
�E�g�p�ł��Ȃ�����: \ / : * ? "" < > |
�E�\���͎g�p�ł��܂���iCON, PRN, AUX���j
�E�����̃t�@�C���Ɠ������O�͎g�p�ł��܂���
�E200�����ȓ��œ��͂��Ă�������

�y���ӎ����z
�E�g���q�͎����I�ɕt������邽�߁A���͕s�v�ł�
�E�ύX�O�Ƀt�@�C���̃o�b�N�A�b�v�𐄏����܂�
�E�d�v�ȃt�@�C���͐T�d�ɑ��삵�Ă�������";

            System.Windows.MessageBox.Show(helpMessage, "�ꊇ���l�[���@�\�̃w���v", 
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// �f�[�^�e���v���[�g���̃e�L�X�g��ݒ�
        /// </summary>
        private void SetDataTemplateTexts()
        {
            // DataGrid�̍s�����񂵂�TextBlock��ݒ�
            for (int i = 0; i < dgFiles.Items.Count; i++)
            {
                var row = dgFiles.ItemContainerGenerator.ContainerFromIndex(i) as System.Windows.Controls.DataGridRow;
                if (row != null)
                {
                    var changedTextBlock = FindVisualChild<System.Windows.Controls.TextBlock>(row, "txtChanged");
                    var errorTextBlock = FindVisualChild<System.Windows.Controls.TextBlock>(row, "txtError");
                    
                    if (changedTextBlock != null)
                        changedTextBlock.Text = "?�ύX";
                    if (errorTextBlock != null)
                        errorTextBlock.Text = "?�G���[";
                }
            }
        }
        
        /// <summary>
        /// �q�v�f�𖼑O�Ō���
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
        /// �v���p�e�B�ύX�C�x���g
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// �v���p�e�B�ύX�ʒm
        /// </summary>
        /// <param name="propertyName">�v���p�e�B��</param>
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// �O�u���E��u�����̓_�C�A���O
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
                Content = "�L�����Z��", 
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