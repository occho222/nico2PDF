using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.Json.Serialization;

namespace Nico2PDF.Models
{
    /// <summary>
    /// �v���W�F�N�g�f�[�^���f��
    /// </summary>
    public class ProjectData : INotifyPropertyChanged
    {
        private string _name = "";
        private bool _isActive = false;
        private bool _includeSubfolders = false;
        private int _subfolderDepth = 1;
        private bool _useCustomPdfPath = false;
        private string _customPdfPath = "";

        /// <summary>
        /// �v���W�F�N�gID
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// �v���W�F�N�g��
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
        /// �v���W�F�N�g�̊K�w���x��
        /// </summary>
        public int Level { get; set; } = 0;

        /// <summary>
        /// �A�N�e�B�u���
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
        /// �v���W�F�N�g�t�H���_�̃p�X
        /// </summary>
        public string FolderPath { get; set; } = "";

        /// <summary>
        /// �T�u�t�H���_���܂ނ��ǂ���
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
        /// サブフォルダの読み込み階層数（1-5、デフォルト1）
        /// </summary>
        public int SubfolderDepth
        {
            get => _subfolderDepth;
            set
            {
                // 1-5の範囲でクランプ
                _subfolderDepth = Math.Max(1, Math.Min(5, value));
                OnPropertyChanged(nameof(SubfolderDepth));
            }
        }

        /// <summary>
        /// �J�X�^��PDF�ۑ��p�X���g�p���邩�ǂ���
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
        /// �J�X�^��PDF�ۑ��p�X
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
        /// PDF�o�̓t�H���_�̃p�X
        /// </summary>
        public string PdfOutputFolder 
        { 
            get 
            {
                if (UseCustomPdfPath && !string.IsNullOrEmpty(CustomPdfPath))
                {
                    // �J�X�^���p�X���g�p����ꍇ�APDF�t�H���_���쐬
                    return Path.Combine(CustomPdfPath, "PDF");
                }
                return Path.Combine(FolderPath, "PDF");
            }
            set
            {
                // ����݊����̂��߁Asetter �͕ێ�
                if (!UseCustomPdfPath)
                {
                    // �]���ʂ�̓���
                }
            }
        }

        /// <summary>
        /// ����PDF�ۑ��t�H���_�̃p�X
        /// </summary>
        [JsonIgnore]
        public string MergePdfFolder
        {
            get
            {
                if (UseCustomPdfPath && !string.IsNullOrEmpty(CustomPdfPath))
                {
                    // �J�X�^���p�X���g�p����ꍇ�A�J�X�^���p�X����mergePDF�t�H���_���쐬
                    return Path.Combine(CustomPdfPath, "mergePDF");
                }
                // �ʏ�̓v���W�F�N�g�t�H���_����mergePDF�t�H���_���쐬
                return Path.Combine(FolderPath, "mergePDF");
            }
        }

        /// <summary>
        /// ����PDF�t�@�C����
        /// </summary>
        public string MergeFileName { get; set; } = "結合PDF";

        /// <summary>
        /// �y�[�W�ԍ��ǉ��t���O
        /// </summary>
        public bool AddPageNumber { get; set; } = false;

        /// <summary>
        /// ������ǉ��t���O
        /// </summary>
        public bool AddBookmarks { get; set; } = true;

        /// <summary>
        /// �t�H���_�ʃO���[�v���t���O
        /// </summary>
        public bool GroupByFolder { get; set; } = false;

        /// <summary>
        /// �w�b�_�E�t�b�^�ǉ��t���O
        /// </summary>
        public bool AddHeaderFooter { get; set; } = false;

        /// <summary>
        /// �w�b�_�E�t�b�^�e�L�X�g
        /// </summary>
        public string HeaderFooterText { get; set; } = "";

        /// <summary>
        /// �w�b�_�E�t�b�^�t�H���g�T�C�Y
        /// </summary>
        public float HeaderFooterFontSize { get; set; } = 10.0f;

        /// <summary>
        /// ページ振りの位置（0:右上, 1:右下, 2:左上, 3:左下）
        /// </summary>
        public int PageNumberPosition { get; set; } = 0;

        /// <summary>
        /// ページ振りのX軸オフセット
        /// </summary>
        public float PageNumberOffsetX { get; set; } = 20.0f;

        /// <summary>
        /// ページ振りのY軸オフセット
        /// </summary>
        public float PageNumberOffsetY { get; set; } = 20.0f;

        /// <summary>
        /// ヘッダの位置（0:左, 1:中央, 2:右）
        /// </summary>
        public int HeaderPosition { get; set; } = 0;

        /// <summary>
        /// ヘッダのX軸オフセット
        /// </summary>
        public float HeaderOffsetX { get; set; } = 20.0f;

        /// <summary>
        /// ヘッダのY軸オフセット
        /// </summary>
        public float HeaderOffsetY { get; set; } = 20.0f;

        /// <summary>
        /// フッタの位置（0:左, 1:中央, 2:右）
        /// </summary>
        public int FooterPosition { get; set; } = 2;

        /// <summary>
        /// フッタのX軸オフセット
        /// </summary>
        public float FooterOffsetX { get; set; } = 20.0f;

        /// <summary>
        /// フッタのY軸オフセット
        /// </summary>
        public float FooterOffsetY { get; set; } = 20.0f;

        /// <summary>
        /// �ŐV�̌���PDF�t�@�C���p�X
        /// </summary>
        public string LatestMergedPdfPath { get; set; } = "";

        /// <summary>
        /// �쐬����
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;

        /// <summary>
        /// �ŏI�A�N�Z�X����
        /// </summary>
        public DateTime LastAccessDate { get; set; } = DateTime.Now;

        /// <summary>
        /// �t�@�C���A�C�e�����X�g
        /// </summary>
        public List<FileItemData> FileItems { get; set; } = new List<FileItemData>();

        /// <summary>
        /// �v���W�F�N�g�J�e�S���i�t�H���_�����p�j
        /// </summary>
        public string Category { get; set; } = "";

        /// <summary>
        /// �J�e�S���̐���
        /// </summary>
        public string CategoryDescription { get; set; } = "";

        /// <summary>
        /// �J�e�S���̐F�i�\���p�j
        /// </summary>
        public string CategoryColor { get; set; } = "#E9ECEF";

        /// <summary>
        /// �J�e�S���A�C�R���i�\���p�j
        /// </summary>
        public string CategoryIcon { get; set; } = "??";

        /// <summary>
        /// �\�����iJSON��Ώہj
        /// </summary>
        [JsonIgnore]
        public string DisplayName => string.IsNullOrEmpty(Category) ? Name : $"{Name}";

        /// <summary>
        /// �J�e�S���t���\�����iJSON��Ώہj
        /// </summary>
        [JsonIgnore]
        public string CategoryDisplayName => string.IsNullOrEmpty(Category) ? $"?? {Name}" : $"{CategoryIcon} {Name}";

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
}