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
    /// PDF処理サービス（ページ抽出、分割、その他の操作）
    /// </summary>
    public class PdfProcessingService
    {
        /// <summary>
        /// PDFファイルから指定されたページのみを抽出
        /// </summary>
        /// <param name="inputPdfPath">入力PDFファイルパス</param>
        /// <param name="outputPdfPath">出力PDFファイルパス</param>
        /// <param name="pageRange">ページ範囲文字列（例: "1,3,5-7"）</param>
        public static void ExtractPages(string inputPdfPath, string outputPdfPath, string pageRange)
        {
            if (string.IsNullOrWhiteSpace(pageRange))
            {
                throw new ArgumentException("ページ範囲が指定されていません。", nameof(pageRange));
            }

            var pageNumbers = ParsePageRange(pageRange);
            if (!pageNumbers.Any())
            {
                throw new ArgumentException("有効なページ範囲が指定されていません。", nameof(pageRange));
            }

            using (var inputReader = new PdfReader(inputPdfPath))
            {
                var totalPages = inputReader.NumberOfPages;

                // 存在しないページのチェック
                var invalidPages = pageNumbers.Where(p => p > totalPages).ToList();
                if (invalidPages.Any())
                {
                    throw new ArgumentException($"存在しないページ番号が指定されています: {string.Join(", ", invalidPages)} (総ページ数: {totalPages})");
                }

                // 指定されたページのみを抽出してPDFを作成
                ExtractPdfPages(inputPdfPath, outputPdfPath, pageNumbers);
            }
        }

        /// <summary>
        /// PDFファイルを各ページ別に分割
        /// </summary>
        /// <param name="inputPdfPath">入力PDFファイルパス</param>
        /// <param name="outputFolder">出力フォルダ</param>
        /// <param name="fileNamePattern">ファイル名パターン（例: "page_{0}.pdf"）</param>
        /// <returns>作成されたファイルパスのリスト</returns>
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
                    
                    // 各ページを個別のPDFファイルとして保存
                    ExtractPdfPages(inputPdfPath, outputPath, new List<int> { i });
                    outputFiles.Add(outputPath);
                }
            }

            return outputFiles;
        }

        /// <summary>
        /// PDFファイルから指定範囲のページを削除
        /// </summary>
        /// <param name="inputPdfPath">入力PDFファイルパス</param>
        /// <param name="outputPdfPath">出力PDFファイルパス</param>
        /// <param name="pageRangeToRemove">削除するページ範囲文字列（例: "1,3,5-7"）</param>
        public static void RemovePages(string inputPdfPath, string outputPdfPath, string pageRangeToRemove)
        {
            if (string.IsNullOrWhiteSpace(pageRangeToRemove))
            {
                throw new ArgumentException("削除するページ範囲が指定されていません。", nameof(pageRangeToRemove));
            }

            var pagesToRemove = ParsePageRange(pageRangeToRemove);
            if (!pagesToRemove.Any())
            {
                throw new ArgumentException("有効なページ範囲が指定されていません。", nameof(pageRangeToRemove));
            }

            using (var inputReader = new PdfReader(inputPdfPath))
            {
                var totalPages = inputReader.NumberOfPages;

                // 存在しないページのチェック
                var invalidPages = pagesToRemove.Where(p => p > totalPages).ToList();
                if (invalidPages.Any())
                {
                    throw new ArgumentException($"存在しないページ番号が指定されています: {string.Join(", ", invalidPages)} (総ページ数: {totalPages})");
                }

                // 削除対象以外のページを取得
                var pagesToKeep = Enumerable.Range(1, totalPages).Except(pagesToRemove).ToList();
                
                if (!pagesToKeep.Any())
                {
                    throw new ArgumentException("すべてのページを削除することはできません。");
                }

                // 残すページのみを抽出してPDFを作成
                ExtractPdfPages(inputPdfPath, outputPdfPath, pagesToKeep);
            }
        }

        /// <summary>
        /// PDFファイルの情報を取得
        /// </summary>
        /// <param name="pdfPath">PDFファイルパス</param>
        /// <returns>PDFファイル情報</returns>
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
        /// ページ範囲を解析
        /// </summary>
        /// <param name="pageRange">ページ範囲文字列</param>
        /// <returns>ページ番号リスト</returns>
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
                throw new ArgumentException($"無効なページ範囲指定: {pageRange}");
            }

            return pages.OrderBy(p => p).ToList();
        }

        /// <summary>
        /// PDFから指定されたページのみを抽出
        /// </summary>
        /// <param name="inputPdfPath">入力PDFファイルパス</param>
        /// <param name="outputPdfPath">出力PDFファイルパス</param>
        /// <param name="pageNumbers">ページ番号リスト</param>
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
    /// PDFファイル情報
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