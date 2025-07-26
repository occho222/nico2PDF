using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Nico2PDF.Models;

namespace Nico2PDF.Services
{
    /// <summary>
    /// PDF�����T�[�r�X�i�y�[�W���o�A�����A���̑��̑���j
    /// </summary>
    public class PdfProcessingService
    {
        /// <summary>
        /// PDF�t�@�C������w�肳�ꂽ�y�[�W�݂̂𒊏o
        /// </summary>
        /// <param name="inputPdfPath">����PDF�t�@�C���p�X</param>
        /// <param name="outputPdfPath">�o��PDF�t�@�C���p�X</param>
        /// <param name="pageRange">�y�[�W�͈͕�����i��: "1,3,5-7"�j</param>
        public static void ExtractPages(string inputPdfPath, string outputPdfPath, string pageRange)
        {
            if (string.IsNullOrWhiteSpace(pageRange))
            {
                throw new ArgumentException("�y�[�W�͈͂��w�肳��Ă��܂���B", nameof(pageRange));
            }

            var pageNumbers = ParsePageRange(pageRange);
            if (!pageNumbers.Any())
            {
                throw new ArgumentException("�L���ȃy�[�W�͈͂��w�肳��Ă��܂���B", nameof(pageRange));
            }

            using (var inputReader = new PdfReader(inputPdfPath))
            {
                var totalPages = inputReader.NumberOfPages;

                // ���݂��Ȃ��y�[�W�̃`�F�b�N
                var invalidPages = pageNumbers.Where(p => p > totalPages).ToList();
                if (invalidPages.Any())
                {
                    throw new ArgumentException($"���݂��Ȃ��y�[�W�ԍ����w�肳��Ă��܂�: {string.Join(", ", invalidPages)} (���y�[�W��: {totalPages})");
                }

                // �w�肳�ꂽ�y�[�W�݂̂𒊏o����PDF���쐬
                ExtractPdfPages(inputPdfPath, outputPdfPath, pageNumbers);
            }
        }

        /// <summary>
        /// PDF�t�@�C�����e�y�[�W�ʂɕ���
        /// </summary>
        /// <param name="inputPdfPath">����PDF�t�@�C���p�X</param>
        /// <param name="outputFolder">�o�̓t�H���_</param>
        /// <param name="fileNamePattern">�t�@�C�����p�^�[���i��: "page_{0}.pdf"�j</param>
        /// <returns>�쐬���ꂽ�t�@�C���p�X�̃��X�g</returns>
        public static List<string> SplitPdfByPage(string inputPdfPath, string outputFolder, string fileNamePattern = "page_{0}.pdf")
        {
            var outputFiles = new List<string>();

            using (var inputReader = new PdfReader(inputPdfPath))
            {
                var totalPages = inputReader.NumberOfPages;

                for (int i = 1; i <= totalPages; i++)
                {
                    var outputFileName = string.Format(fileNamePattern, i);
                    var outputPath = Path.Combine(outputFolder, outputFileName);
                    
                    // �e�y�[�W���ʂ�PDF�t�@�C���Ƃ��ĕۑ�
                    ExtractPdfPages(inputPdfPath, outputPath, new List<int> { i });
                    outputFiles.Add(outputPath);
                }
            }

            return outputFiles;
        }

        /// <summary>
        /// PDF�t�@�C������w��͈͂̃y�[�W���폜
        /// </summary>
        /// <param name="inputPdfPath">����PDF�t�@�C���p�X</param>
        /// <param name="outputPdfPath">�o��PDF�t�@�C���p�X</param>
        /// <param name="pageRangeToRemove">�폜����y�[�W�͈͕�����i��: "1,3,5-7"�j</param>
        public static void RemovePages(string inputPdfPath, string outputPdfPath, string pageRangeToRemove)
        {
            if (string.IsNullOrWhiteSpace(pageRangeToRemove))
            {
                throw new ArgumentException("�폜����y�[�W�͈͂��w�肳��Ă��܂���B", nameof(pageRangeToRemove));
            }

            var pagesToRemove = ParsePageRange(pageRangeToRemove);
            if (!pagesToRemove.Any())
            {
                throw new ArgumentException("�L���ȃy�[�W�͈͂��w�肳��Ă��܂���B", nameof(pageRangeToRemove));
            }

            using (var inputReader = new PdfReader(inputPdfPath))
            {
                var totalPages = inputReader.NumberOfPages;

                // ���݂��Ȃ��y�[�W�̃`�F�b�N
                var invalidPages = pagesToRemove.Where(p => p > totalPages).ToList();
                if (invalidPages.Any())
                {
                    throw new ArgumentException($"���݂��Ȃ��y�[�W�ԍ����w�肳��Ă��܂�: {string.Join(", ", invalidPages)} (���y�[�W��: {totalPages})");
                }

                // �폜�ΏۈȊO�̃y�[�W���擾
                var pagesToKeep = Enumerable.Range(1, totalPages).Except(pagesToRemove).ToList();
                
                if (!pagesToKeep.Any())
                {
                    throw new ArgumentException("���ׂẴy�[�W���폜���邱�Ƃ͂ł��܂���B");
                }

                // �c���y�[�W�݂̂𒊏o����PDF���쐬
                ExtractPdfPages(inputPdfPath, outputPdfPath, pagesToKeep);
            }
        }

        /// <summary>
        /// PDF�t�@�C���̏����擾
        /// </summary>
        /// <param name="pdfPath">PDF�t�@�C���p�X</param>
        /// <returns>PDF�t�@�C�����</returns>
        public static PdfFileInfo GetPdfInfo(string pdfPath)
        {
            using (var reader = new PdfReader(pdfPath))
            {
                return new PdfFileInfo
                {
                    FilePath = pdfPath,
                    TotalPages = reader.NumberOfPages,
                    Title = reader.Info.ContainsKey("Title") ? reader.Info["Title"] : "",
                    Author = reader.Info.ContainsKey("Author") ? reader.Info["Author"] : "",
                    Subject = reader.Info.ContainsKey("Subject") ? reader.Info["Subject"] : "",
                    Creator = reader.Info.ContainsKey("Creator") ? reader.Info["Creator"] : "",
                    Producer = reader.Info.ContainsKey("Producer") ? reader.Info["Producer"] : "",
                    CreationDate = reader.Info.ContainsKey("CreationDate") ? reader.Info["CreationDate"] : "",
                    ModificationDate = reader.Info.ContainsKey("ModDate") ? reader.Info["ModDate"] : ""
                };
            }
        }

        /// <summary>
        /// �y�[�W�͈͂����
        /// </summary>
        /// <param name="pageRange">�y�[�W�͈͕�����</param>
        /// <returns>�y�[�W�ԍ����X�g</returns>
        private static List<int> ParsePageRange(string pageRange)
        {
            var pages = new List<int>();
            if (string.IsNullOrWhiteSpace(pageRange))
                return pages;

            try
            {
                var parts = pageRange.Split(',');
                foreach (var part in parts)
                {
                    var trimmed = part.Trim();
                    if (trimmed.Contains('-'))
                    {
                        var range = trimmed.Split('-');
                        if (range.Length == 2 && int.TryParse(range[0], out int start) && int.TryParse(range[1], out int end))
                        {
                            for (int i = start; i <= end; i++)
                            {
                                if (i > 0 && !pages.Contains(i))
                                    pages.Add(i);
                            }
                        }
                    }
                    else if (int.TryParse(trimmed, out int page))
                    {
                        if (page > 0 && !pages.Contains(page))
                            pages.Add(page);
                    }
                }
            }
            catch
            {
                throw new ArgumentException($"�����ȃy�[�W�͈͎w��: {pageRange}");
            }

            return pages.OrderBy(p => p).ToList();
        }

        /// <summary>
        /// PDF����w�肳�ꂽ�y�[�W�݂̂𒊏o
        /// </summary>
        /// <param name="inputPdfPath">����PDF�t�@�C���p�X</param>
        /// <param name="outputPdfPath">�o��PDF�t�@�C���p�X</param>
        /// <param name="pageNumbers">�y�[�W�ԍ����X�g</param>
        private static void ExtractPdfPages(string inputPdfPath, string outputPdfPath, List<int> pageNumbers)
        {
            using (var inputReader = new PdfReader(inputPdfPath))
            using (var outputDocument = new Document())
            using (var outputWriter = new PdfCopy(outputDocument, new FileStream(outputPdfPath, FileMode.Create)))
            {
                outputDocument.Open();

                foreach (var pageNumber in pageNumbers.OrderBy(p => p))
                {
                    if (pageNumber <= inputReader.NumberOfPages)
                    {
                        var page = outputWriter.GetImportedPage(inputReader, pageNumber);
                        outputWriter.AddPage(page);
                    }
                }
            }
        }
    }

    /// <summary>
    /// PDF�t�@�C�����
    /// </summary>
    public class PdfFileInfo
    {
        public string FilePath { get; set; } = "";
        public int TotalPages { get; set; }
        public string Title { get; set; } = "";
        public string Author { get; set; } = "";
        public string Subject { get; set; } = "";
        public string Creator { get; set; } = "";
        public string Producer { get; set; } = "";
        public string CreationDate { get; set; } = "";
        public string ModificationDate { get; set; } = "";
    }
}