using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Threading.Tasks;
using Nico2PDF.Models;

namespace Nico2PDF.Services
{
    /// <summary>
    /// PDF結合サービス
    /// </summary>
    public class PdfMergeService
    {
        /// <summary>
        /// PDFファイルを結合
        /// </summary>
        /// <param name="pdfFilePaths">結合するPDFファイルパスリスト</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="addPageNumber">ページ番号追加フラグ</param>
        /// <param name="addBookmarks">しおり追加フラグ</param>
        /// <param name="fileItems">ファイルアイテムリスト（しおり生成用）</param>
        /// <param name="addHeaderFooter">ヘッダ・フッタ追加フラグ</param>
        /// <param name="headerFooterText">ヘッダ・フッタテキスト</param>
        /// <param name="headerFooterFontSize">ヘッダ・フッタフォントサイズ</param>
        public static void MergePdfFiles(List<string> pdfFilePaths, string outputPath, bool addPageNumber = false, bool addBookmarks = false, List<FileItem> fileItems = null, bool addHeaderFooter = false, string headerFooterText = "", float headerFooterFontSize = 10.0f)
        {
            using (var document = new Document())
            using (var copy = new PdfCopy(document, new FileStream(outputPath, FileMode.Create)))
            {
                document.Open();

                // しおり用の情報を記録
                var bookmarks = new List<Dictionary<string, object>>();
                int currentPage = 1;

                for (int fileIndex = 0; fileIndex < pdfFilePaths.Count; fileIndex++)
                {
                    var pdfPath = pdfFilePaths[fileIndex];
                    var startPage = currentPage;

                    using (var reader = new PdfReader(pdfPath))
                    {
                        // PDFのページを追加
                        for (int i = 1; i <= reader.NumberOfPages; i++)
                        {
                            var page = copy.GetImportedPage(reader, i);
                            copy.AddPage(page);
                        }

                        // しおり情報を追加
                        if (addBookmarks && fileItems != null && fileIndex < fileItems.Count)
                        {
                            var fileItem = fileItems[fileIndex];
                            var bookmark = new Dictionary<string, object>
                            {
                                ["Title"] = GetBookmarkTitle(fileItem),
                                ["Action"] = "GoTo",
                                ["Page"] = $"{startPage} Fit"
                            };
                            bookmarks.Add(bookmark);
                        }

                        currentPage += reader.NumberOfPages;
                    }
                }

                // しおりを設定
                if (addBookmarks && bookmarks.Any())
                {
                    copy.Outlines = bookmarks;
                }
            }

            // ページ番号またはヘッダ・フッタ追加（オプション）
            if (addPageNumber || addHeaderFooter)
            {
                AddPageNumbersAndHeaderFooter(outputPath, addPageNumber, addHeaderFooter, headerFooterText, headerFooterFontSize);
            }
        }

        /// <summary>
        /// しおりのタイトルを生成
        /// </summary>
        /// <param name="fileItem">ファイルアイテム</param>
        /// <returns>しおりタイトル</returns>
        private static string GetBookmarkTitle(FileItem fileItem)
        {
            // 表示名（リネームされている場合はそれを使用）を拡張子なしで取得
            var displayName = fileItem.DisplayName;
            var titleWithoutExtension = Path.GetFileNameWithoutExtension(displayName);
            
            // 番号を追加してより構造化されたしおりタイトルにする
            return $"{fileItem.Number:D3}_{titleWithoutExtension}";
        }

        /// <summary>
        /// PDFにページ番号とヘッダ・フッタを追加
        /// </summary>
        /// <param name="pdfPath">PDFファイルパス</param>
        /// <param name="addPageNumber">ページ番号追加フラグ</param>
        /// <param name="addHeaderFooter">ヘッダ・フッタ追加フラグ</param>
        /// <param name="headerFooterText">ヘッダ・フッタテキスト</param>
        /// <param name="headerFooterFontSize">ヘッダ・フッタフォントサイズ</param>
        private static void AddPageNumbersAndHeaderFooter(string pdfPath, bool addPageNumber, bool addHeaderFooter, string headerFooterText, float headerFooterFontSize)
        {
            var tempPath = pdfPath + ".tmp";

            using (var reader = new PdfReader(pdfPath))
            using (var stamper = new PdfStamper(reader, new FileStream(tempPath, FileMode.Create)))
            {
                var totalPages = reader.NumberOfPages;

                for (int i = 1; i <= totalPages; i++)
                {
                    var cb = stamper.GetOverContent(i);
                    var pageSize = reader.GetPageSize(i);

                    // ページ番号を追加（右上）
                    if (addPageNumber)
                    {
                        cb.BeginText();
                        cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
                        // ページ番号を右上に配置（x座標：右端から20pt、y座標：上端から20pt）
                        cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT,
                            $"{i} / {totalPages}",
                            pageSize.Width - 20, pageSize.Height - 20, 0);
                        cb.EndText();
                    }

                    // ヘッダ・フッタを追加
                    if (addHeaderFooter && !string.IsNullOrEmpty(headerFooterText))
                    {
                        var font = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
                        
                        // ヘッダ（左上）
                        cb.BeginText();
                        cb.SetFontAndSize(font, headerFooterFontSize);
                        // ヘッダを左上に配置
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT,
                            headerFooterText,
                            20, pageSize.Height - 20, 0);
                        cb.EndText();

                        // フッタ（右下）
                        cb.BeginText();
                        cb.SetFontAndSize(font, headerFooterFontSize);
                        // フッタを右下に配置
                        cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT,
                            headerFooterText,
                            pageSize.Width - 20, 20, 0);
                        cb.EndText();
                    }
                }
            }

            File.Delete(pdfPath);
            File.Move(tempPath, pdfPath);
        }

        /// <summary>
        /// 高度なしおり機能：階層構造でしおりを作成
        /// </summary>
        /// <param name="pdfFilePaths">結合するPDFファイルパスリスト</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="fileItems">ファイルアイテムリスト</param>
        /// <param name="addPageNumber">ページ番号追加フラグ</param>
        /// <param name="groupByFolder">フォルダ別にグループ化するかどうか</param>
        /// <param name="addHeaderFooter">ヘッダ・フッタ追加フラグ</param>
        /// <param name="headerFooterText">ヘッダ・フッタテキスト</param>
        /// <param name="headerFooterFontSize">ヘッダ・フッタフォントサイズ</param>
        public static void MergePdfFilesWithAdvancedBookmarks(List<string> pdfFilePaths, string outputPath, List<FileItem> fileItems, bool addPageNumber = false, bool groupByFolder = false, bool addHeaderFooter = false, string headerFooterText = "", float headerFooterFontSize = 10.0f)
        {
            using (var document = new Document())
            using (var copy = new PdfCopy(document, new FileStream(outputPath, FileMode.Create)))
            {
                document.Open();

                var bookmarks = new List<Dictionary<string, object>>();
                int currentPage = 1;

                if (groupByFolder)
                {
                    // フォルダ別にグループ化したしおり構造を作成
                    CreateFolderGroupedBookmarks(pdfFilePaths, fileItems, copy, bookmarks, ref currentPage);
                }
                else
                {
                    // 通常のフラットなしおり構造を作成
                    CreateFlatBookmarks(pdfFilePaths, fileItems, copy, bookmarks, ref currentPage);
                }

                // しおりを設定
                if (bookmarks.Any())
                {
                    copy.Outlines = bookmarks;
                }
            }

            // ページ番号またはヘッダ・フッタ追加（オプション）
            if (addPageNumber || addHeaderFooter)
            {
                AddPageNumbersAndHeaderFooter(outputPath, addPageNumber, addHeaderFooter, headerFooterText, headerFooterFontSize);
            }
        }

        /// <summary>
        /// フラットなしおり構造を作成
        /// </summary>
        private static void CreateFlatBookmarks(List<string> pdfFilePaths, List<FileItem> fileItems, PdfCopy copy, List<Dictionary<string, object>> bookmarks, ref int currentPage)
        {
            for (int fileIndex = 0; fileIndex < pdfFilePaths.Count; fileIndex++)
            {
                var pdfPath = pdfFilePaths[fileIndex];
                var startPage = currentPage;

                using (var reader = new PdfReader(pdfPath))
                {
                    // PDFのページを追加
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        var page = copy.GetImportedPage(reader, i);
                        copy.AddPage(page);
                    }

                    // しおり情報を追加
                    if (fileItems != null && fileIndex < fileItems.Count)
                    {
                        var fileItem = fileItems[fileIndex];
                        var bookmark = new Dictionary<string, object>
                        {
                            ["Title"] = GetBookmarkTitle(fileItem),
                            ["Action"] = "GoTo",
                            ["Page"] = $"{startPage} Fit"
                        };
                        bookmarks.Add(bookmark);
                    }

                    currentPage += reader.NumberOfPages;
                }
            }
        }

        /// <summary>
        /// フォルダ別にグループ化したしおり構造を作成
        /// </summary>
        private static void CreateFolderGroupedBookmarks(List<string> pdfFilePaths, List<FileItem> fileItems, PdfCopy copy, List<Dictionary<string, object>> bookmarks, ref int currentPage)
        {
            // フォルダ別にファイルをグループ化
            var folderGroups = fileItems?
                .Select((item, index) => new { Item = item, Index = index, PdfPath = pdfFilePaths[index] })
                .GroupBy(x => x.Item.FolderName ?? "ルート")
                .OrderBy(g => g.Key)
                .ToList();

            if (folderGroups == null || !folderGroups.Any())
            {
                // フォルダ情報がない場合はフラット構造にフォールバック
                CreateFlatBookmarks(pdfFilePaths, fileItems, copy, bookmarks, ref currentPage);
                return;
            }

            foreach (var folderGroup in folderGroups)
            {
                var folderStartPage = currentPage;
                var childBookmarks = new List<Dictionary<string, object>>();

                // フォルダ内の各ファイルを処理
                foreach (var fileInfo in folderGroup.OrderBy(x => x.Item.DisplayOrder))
                {
                    var startPage = currentPage;

                    using (var reader = new PdfReader(fileInfo.PdfPath))
                    {
                        // PDFのページを追加
                        for (int i = 1; i <= reader.NumberOfPages; i++)
                        {
                            var page = copy.GetImportedPage(reader, i);
                            copy.AddPage(page);
                        }

                        // 子しおり情報を追加
                        var childBookmark = new Dictionary<string, object>
                        {
                            ["Title"] = GetBookmarkTitle(fileInfo.Item),
                            ["Action"] = "GoTo",
                            ["Page"] = $"{startPage} Fit"
                        };
                        childBookmarks.Add(childBookmark);

                        currentPage += reader.NumberOfPages;
                    }
                }

                // フォルダレベルのしおりを作成
                var folderBookmark = new Dictionary<string, object>
                {
                    ["Title"] = folderGroup.Key,
                    ["Action"] = "GoTo",
                    ["Page"] = $"{folderStartPage} Fit",
                    ["Kids"] = childBookmarks
                };
                bookmarks.Add(folderBookmark);
            }
        }
    }
}