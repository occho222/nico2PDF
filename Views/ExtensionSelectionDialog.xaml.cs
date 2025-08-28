using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Nico2PDF.Models;

namespace Nico2PDF.Views
{
    /// <summary>
    /// 拡張子別ファイル選択ダイアログ
    /// </summary>
    public partial class ExtensionSelectionDialog : Window
    {
        /// <summary>
        /// ファイル選択を適用するためのコールバック
        /// </summary>
        public Action<Func<string, bool>, string>? SelectFilesByExtension { get; set; }

        /// <summary>
        /// ファイル一覧
        /// </summary>
        private List<FileItem> _fileItems;

        public ExtensionSelectionDialog()
        {
            InitializeComponent();
            _fileItems = new List<FileItem>();
        }

        /// <summary>
        /// ファイル一覧を設定
        /// </summary>
        public void SetFileItems(List<FileItem> fileItems)
        {
            _fileItems = fileItems;
            UpdateFileCounts();
        }

        /// <summary>
        /// 各拡張子のファイル数を更新
        /// </summary>
        private void UpdateFileCounts()
        {
            var excelCount = _fileItems.Count(f => IsExcelFile(f.Extension));
            var wordCount = _fileItems.Count(f => IsWordFile(f.Extension));
            var powerPointCount = _fileItems.Count(f => IsPowerPointFile(f.Extension));
            var pdfCount = _fileItems.Count(f => IsPdfFile(f.Extension));
            var officeCount = _fileItems.Count(f => IsOfficeFile(f.Extension));

            txtExcelCount.Text = $"対象ファイル: {excelCount}個";
            txtWordCount.Text = $"対象ファイル: {wordCount}個";
            txtPowerPointCount.Text = $"対象ファイル: {powerPointCount}個";
            txtPdfCount.Text = $"対象ファイル: {pdfCount}個";
            txtOfficeCount.Text = $"対象ファイル: {officeCount}個";

            // ファイル数が0の場合はボタンを無効化
            btnSelectExcel.IsEnabled = excelCount > 0;
            btnSelectWord.IsEnabled = wordCount > 0;
            btnSelectPowerPoint.IsEnabled = powerPointCount > 0;
            btnSelectPDF.IsEnabled = pdfCount > 0;
            btnSelectOffice.IsEnabled = officeCount > 0;
        }

        private void BtnSelectExcel_Click(object sender, RoutedEventArgs e)
        {
            SelectFilesByExtension?.Invoke(IsExcelFile, "Excel");
            Close();
        }

        private void BtnSelectWord_Click(object sender, RoutedEventArgs e)
        {
            SelectFilesByExtension?.Invoke(IsWordFile, "Word");
            Close();
        }

        private void BtnSelectPowerPoint_Click(object sender, RoutedEventArgs e)
        {
            SelectFilesByExtension?.Invoke(IsPowerPointFile, "PowerPoint");
            Close();
        }

        private void BtnSelectPDF_Click(object sender, RoutedEventArgs e)
        {
            SelectFilesByExtension?.Invoke(IsPdfFile, "PDF");
            Close();
        }

        private void BtnSelectOffice_Click(object sender, RoutedEventArgs e)
        {
            SelectFilesByExtension?.Invoke(IsOfficeFile, "Office文書");
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #region ファイル拡張子判定メソッド

        private static bool IsExcelFile(string extension)
        {
            return extension.ToUpper() is "XLS" or "XLSX" or "XLSM";
        }

        private static bool IsWordFile(string extension)
        {
            return extension.ToUpper() is "DOC" or "DOCX";
        }

        private static bool IsPowerPointFile(string extension)
        {
            return extension.ToUpper() is "PPT" or "PPTX";
        }

        private static bool IsPdfFile(string extension)
        {
            return extension.ToUpper() is "PDF";
        }

        private static bool IsOfficeFile(string extension)
        {
            return IsExcelFile(extension) || IsWordFile(extension) || IsPowerPointFile(extension);
        }

        #endregion
    }
}