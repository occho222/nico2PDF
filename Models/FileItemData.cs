using System;

namespace Nico2PDF.Models
{
    /// <summary>
    /// �t�@�C���A�C�e���f�[�^�i�ۑ��p�j
    /// </summary>
    public class FileItemData
    {
        /// <summary>
        /// �I�����
        /// </summary>
        public bool IsSelected { get; set; }

        /// <summary>
        /// �Ώۃy�[�W
        /// </summary>
        public string TargetPages { get; set; } = "";

        /// <summary>
        /// �t�@�C���p�X
        /// </summary>
        public string FilePath { get; set; } = "";

        /// <summary>
        /// �ŏI�X�V����
        /// </summary>
        public DateTime LastModified { get; set; }

        /// <summary>
        /// �\������
        /// </summary>
        public int DisplayOrder { get; set; } = 0;

        /// <summary>
        /// ���΃p�X�i�T�u�t�H���_�ǂݍ��ݗp�j
        /// </summary>
        public string RelativePath { get; set; } = "";

        /// <summary>
        /// �\�����i���[�U�[���ύX�\�j
        /// </summary>
        public string DisplayName { get; set; } = "";

        /// <summary>
        /// ���̃t�@�C�����i�ύX�O�̖��O��ێ��j
        /// </summary>
        public string OriginalFileName { get; set; } = "";
    }
}