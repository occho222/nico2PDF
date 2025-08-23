using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Threading.Tasks;
using Nico2PDF.Models;
using System.Diagnostics;
using System.Windows;

namespace Nico2PDF.Services
{
    /// <summary>
    /// PDF�����T�[�r�X
    /// </summary>
    public class PdfMergeService
    {
        // 日本語フォントのキャッシュ
        private static BaseFont? _cachedJapaneseFont = null;
        private static string? _cachedFontName = null;
        private static readonly object _fontLock = new object();
        /// <summary>
        /// PDF�t�@�C��������
        /// </summary>
        /// <param name="pdfFilePaths">��������PDF�t�@�C���p�X���X�g</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="addPageNumber">�y�[�W�ԍ��ǉ��t���O</param>
        /// <param name="addBookmarks">������ǉ��t���O</param>
        /// <param name="fileItems">�t�@�C���A�C�e�����X�g�i�����萶���p�j</param>
        /// <param name="addHeaderFooter">�w�b�_�E�t�b�^�ǉ��t���O</param>
        /// <param name="headerFooterText">�w�b�_�E�t�b�^�e�L�X�g</param>
        /// <param name="headerFooterFontSize">�w�b�_�E�t�b�^�t�H���g�T�C�Y</param>
        /// <param name="pageNumberPosition">ページ振りの位置（0:右上, 1:右下, 2:左上, 3:左下）</param>
        /// <param name="pageNumberOffsetX">ページ振りのX軸オフセット</param>
        /// <param name="pageNumberOffsetY">ページ振りのY軸オフセット</param>
        /// <param name="headerPosition">ヘッダの位置（0:左, 1:中央, 2:右）</param>
        /// <param name="headerOffsetX">ヘッダのX軸オフセット</param>
        /// <param name="headerOffsetY">ヘッダのY軸オフセット</param>
        /// <param name="footerPosition">フッタの位置（0:左, 1:中央, 2:右）</param>
        /// <param name="footerOffsetX">フッタのX軸オフセット</param>
        /// <param name="footerOffsetY">フッタのY軸オフセット</param>
        /// <param name="pageNumberFontSize">ページ番号のフォントサイズ</param>
        /// <param name="headerFontSize">ヘッダのフォントサイズ</param>
        /// <param name="footerFontSize">フッタのフォントサイズ</param>
        public static void MergePdfFiles(List<string> pdfFilePaths, string outputPath, bool addPageNumber = false, bool addBookmarks = false, List<FileItem> fileItems = null, bool addHeaderFooter = false, string headerFooterText = "", float headerFooterFontSize = 10.0f,
            int pageNumberPosition = 0, float pageNumberOffsetX = 20.0f, float pageNumberOffsetY = 20.0f, float pageNumberFontSize = 10.0f,
            int headerPosition = 0, float headerOffsetX = 20.0f, float headerOffsetY = 20.0f, float headerFontSize = 10.0f,
            int footerPosition = 2, float footerOffsetX = 20.0f, float footerOffsetY = 20.0f, float footerFontSize = 10.0f,
            bool addHeader = false, bool addFooter = false, string headerText = "", string footerText = "", string headerFooterFont = "MS Gothic")
        {
            using (var document = new Document())
            using (var copy = new PdfCopy(document, new FileStream(outputPath, FileMode.Create)))
            {
                document.Open();

                // ������p�̏����L�^
                var bookmarks = new List<Dictionary<string, object>>();
                int currentPage = 1;

                for (int fileIndex = 0; fileIndex < pdfFilePaths.Count; fileIndex++)
                {
                    var pdfPath = pdfFilePaths[fileIndex];
                    var startPage = currentPage;

                    using (var reader = new PdfReader(pdfPath))
                    {
                        // PDF�̃y�[�W��ǉ�
                        for (int i = 1; i <= reader.NumberOfPages; i++)
                        {
                            var page = copy.GetImportedPage(reader, i);
                            copy.AddPage(page);
                        }

                        // ���������ǉ�
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

                // �������ݒ�
                if (addBookmarks && bookmarks.Any())
                {
                    copy.Outlines = bookmarks;
                }
            }

            // �y�[�W�ԍ��܂��̓w�b�_�E�t�b�^�ǉ��i�I�v�V�����j
            if (addPageNumber || addHeaderFooter || addHeader || addFooter)
            {
                AddPageNumbersAndHeaderFooter(outputPath, addPageNumber, addHeaderFooter, headerFooterText, headerFooterFontSize,
                    pageNumberPosition, pageNumberOffsetX, pageNumberOffsetY, pageNumberFontSize,
                    headerPosition, headerOffsetX, headerOffsetY, headerFontSize,
                    footerPosition, footerOffsetX, footerOffsetY, footerFontSize,
                    addHeader, addFooter, headerText, footerText, headerFooterFont);
            }
        }

        /// <summary>
        /// ������̃^�C�g���𐶐�
        /// </summary>
        /// <param name="fileItem">�t�@�C���A�C�e��</param>
        /// <returns>������^�C�g��</returns>
        private static string GetBookmarkTitle(FileItem fileItem)
        {
            // �\�����i���l�[������Ă���ꍇ�͂�����g�p�j���g���q�Ȃ��Ŏ擾
            var displayName = fileItem.DisplayName;
            var titleWithoutExtension = Path.GetFileNameWithoutExtension(displayName);
            
            // �ԍ���ǉ����Ă��\�������ꂽ������^�C�g���ɂ���
            return $"{fileItem.Number:D3}_{titleWithoutExtension}";
        }

        /// <summary>
        /// PDF�Ƀy�[�W�ԍ��ƃw�b�_�E�t�b�^��ǉ�
        /// </summary>
        /// <param name="pdfPath">PDF�t�@�C���p�X</param>
        /// <param name="addPageNumber">�y�[�W�ԍ��ǉ��t���O</param>
        /// <param name="addHeaderFooter">�w�b�_�E�t�b�^�ǉ��t���O</param>
        /// <param name="headerFooterText">�w�b�_�E�t�b�^�e�L�X�g</param>
        /// <param name="headerFooterFontSize">�w�b�_�E�t�b�^�t�H���g�T�C�Y</param>
        /// <param name="pageNumberPosition">ページ振りの位置（0:右上, 1:右下, 2:左上, 3:左下）</param>
        /// <param name="pageNumberOffsetX">ページ振りのX軸オフセット</param>
        /// <param name="pageNumberOffsetY">ページ振りのY軸オフセット</param>
        /// <param name="headerPosition">ヘッダの位置（0:左, 1:中央, 2:右）</param>
        /// <param name="headerOffsetX">ヘッダのX軸オフセット</param>
        /// <param name="headerOffsetY">ヘッダのY軸オフセット</param>
        /// <param name="footerPosition">フッタの位置（0:左, 1:中央, 2:右）</param>
        /// <param name="footerOffsetX">フッタのX軸オフセット</param>
        /// <param name="footerOffsetY">フッタのY軸オフセット</param>
        private static void AddPageNumbersAndHeaderFooter(string pdfPath, bool addPageNumber, bool addHeaderFooter, string headerFooterText, float headerFooterFontSize, 
            int pageNumberPosition = 0, float pageNumberOffsetX = 20.0f, float pageNumberOffsetY = 20.0f, float pageNumberFontSize = 10.0f,
            int headerPosition = 0, float headerOffsetX = 20.0f, float headerOffsetY = 20.0f, float headerFontSize = 10.0f,
            int footerPosition = 2, float footerOffsetX = 20.0f, float footerOffsetY = 20.0f, float footerFontSize = 10.0f,
            bool addHeader = false, bool addFooter = false, string headerText = "", string footerText = "", string headerFooterFont = "MS Gothic")
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

                    // ページ振りを追加
                    if (addPageNumber)
                    {
                        cb.BeginText();
                        // 日本語対応フォントを使用
                        var pageNumberFont = CreateJapaneseFont(headerFooterFont);
                        cb.SetFontAndSize(pageNumberFont, pageNumberFontSize);
                        
                        float x, y;
                        int alignment;
                        
                        // 位置に応じてX、Y座標とアライメントを設定
                        switch (pageNumberPosition)
                        {
                            case 0: // 右上
                                x = pageSize.Width - pageNumberOffsetX;
                                y = pageSize.Height - pageNumberOffsetY;
                                alignment = PdfContentByte.ALIGN_RIGHT;
                                break;
                            case 1: // 右下
                                x = pageSize.Width - pageNumberOffsetX;
                                y = pageNumberOffsetY;
                                alignment = PdfContentByte.ALIGN_RIGHT;
                                break;
                            case 2: // 左上
                                x = pageNumberOffsetX;
                                y = pageSize.Height - pageNumberOffsetY;
                                alignment = PdfContentByte.ALIGN_LEFT;
                                break;
                            case 3: // 左下
                                x = pageNumberOffsetX;
                                y = pageNumberOffsetY;
                                alignment = PdfContentByte.ALIGN_LEFT;
                                break;
                            default: // デフォルトは右上
                                x = pageSize.Width - pageNumberOffsetX;
                                y = pageSize.Height - pageNumberOffsetY;
                                alignment = PdfContentByte.ALIGN_RIGHT;
                                break;
                        }
                        
                        cb.ShowTextAligned(alignment, $"{i} / {totalPages}", x, y, 0);
                        cb.EndText();
                    }

                    // 日本語対応フォントを使用
                    var font = CreateJapaneseFont(headerFooterFont);

                    // ヘッダを追加
                    if ((addHeaderFooter && !string.IsNullOrEmpty(headerFooterText)) || (addHeader && !string.IsNullOrEmpty(headerText)))
                    {
                        cb.BeginText();
                        cb.SetFontAndSize(font, headerFontSize);
                        
                        float headerX;
                        int headerAlignment;
                        
                        // ヘッダの位置に応じてX座標とアライメントを設定
                        switch (headerPosition)
                        {
                            case 0: // 左
                                headerX = headerOffsetX;
                                headerAlignment = PdfContentByte.ALIGN_LEFT;
                                break;
                            case 1: // 中央
                                headerX = pageSize.Width / 2;
                                headerAlignment = PdfContentByte.ALIGN_CENTER;
                                break;
                            case 2: // 右
                                headerX = pageSize.Width - headerOffsetX;
                                headerAlignment = PdfContentByte.ALIGN_RIGHT;
                                break;
                            default: // デフォルトは左
                                headerX = headerOffsetX;
                                headerAlignment = PdfContentByte.ALIGN_LEFT;
                                break;
                        }
                        
                        var displayHeaderText = addHeader && !string.IsNullOrEmpty(headerText) ? headerText : headerFooterText;
                        cb.ShowTextAligned(headerAlignment, displayHeaderText, headerX, pageSize.Height - headerOffsetY, 0);
                        cb.EndText();
                    }

                    // フッタを追加
                    if ((addHeaderFooter && !string.IsNullOrEmpty(headerFooterText)) || (addFooter && !string.IsNullOrEmpty(footerText)))
                    {
                        cb.BeginText();
                        cb.SetFontAndSize(font, footerFontSize);
                        
                        float footerX;
                        int footerAlignment;
                        
                        // フッタの位置に応じてX座標とアライメントを設定
                        switch (footerPosition)
                        {
                            case 0: // 左
                                footerX = footerOffsetX;
                                footerAlignment = PdfContentByte.ALIGN_LEFT;
                                break;
                            case 1: // 中央
                                footerX = pageSize.Width / 2;
                                footerAlignment = PdfContentByte.ALIGN_CENTER;
                                break;
                            case 2: // 右
                                footerX = pageSize.Width - footerOffsetX;
                                footerAlignment = PdfContentByte.ALIGN_RIGHT;
                                break;
                            default: // デフォルトは右
                                footerX = pageSize.Width - footerOffsetX;
                                footerAlignment = PdfContentByte.ALIGN_RIGHT;
                                break;
                        }
                        
                        var displayFooterText = addFooter && !string.IsNullOrEmpty(footerText) ? footerText : headerFooterText;
                        cb.ShowTextAligned(footerAlignment, displayFooterText, footerX, footerOffsetY, 0);
                        cb.EndText();
                    }
                }
            }

            File.Delete(pdfPath);
            File.Move(tempPath, pdfPath);
        }

        /// <summary>
        /// 日本語対応フォントを作成（キャッシュ機能付き）
        /// </summary>
        /// <returns>日本語対応BaseFont</returns>
        private static BaseFont CreateJapaneseFont(string preferredFont = "MS Gothic")
        {
            lock (_fontLock)
            {
                // キャッシュされたフォントがあり、同じフォント名の場合は返す
                if (_cachedJapaneseFont != null && _cachedFontName == preferredFont)
                {
                    return _cachedJapaneseFont;
                }

                // 全フォント候補をリスト化
                var allFontCandidates = new[]
                {
                    new { Path = @"C:\Windows\Fonts\NotoSansJP-VF.ttf", Name = "Noto Sans JP", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\YuGothR.ttc,0", Name = "Yu Gothic", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\YuGothM.ttc,0", Name = "Yu Gothic UI", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\BIZ-UDPGothicR.ttc,0", Name = "BIZ UDPGothic", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\BIZ-UDGothicR.ttc,0", Name = "BIZ UDGothic", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\msgothic.ttc,0", Name = "MS Gothic", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\meiryo.ttc,0", Name = "Meiryo", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\meiryo.ttc,1", Name = "Meiryo UI", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\msmincho.ttc,0", Name = "MS Mincho", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\YuMinR.ttc,0", Name = "Yu Mincho", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.NOT_EMBEDDED },
                };

                // 指定されたフォントを最初に試すための候補リストを作成
                var fontCandidates = new List<dynamic>();
                
                // 指定されたフォントを最初に追加
                var preferredCandidate = allFontCandidates.FirstOrDefault(f => f.Name == preferredFont);
                if (preferredCandidate != null)
                {
                    fontCandidates.Add(preferredCandidate);
                }
                
                // 残りのフォントを追加（ゴシック系優先順）
                foreach (var candidate in allFontCandidates)
                {
                    if (candidate.Name != preferredFont)
                    {
                        fontCandidates.Add(candidate);
                    }
                }

                foreach (var candidate in fontCandidates)
                {
                    try
                    {
                        
                        // ファイルパスの場合は存在確認
                        if (candidate.Path.Contains(@"\"))
                        {
                            var filePath = candidate.Path.Contains(",") ? candidate.Path.Split(',')[0] : candidate.Path;
                            if (!File.Exists(filePath))
                            {
                                continue;
                            }
                        }
                        
                        var font = BaseFont.CreateFont(candidate.Path, candidate.Encoding, candidate.Embedded);
                        
                        // 日本語文字のテストを簡略化して、フォントが作成できたら直ちに使用する
                        // ConvertToBytesでのテストは省略し、フォント作成成功で日本語対応とみなす
                        _cachedJapaneseFont = font;  // キャッシュに保存
                        _cachedFontName = preferredFont;  // キャッシュしたフォント名を保存
                        return font;
                    }
                    catch (Exception ex)
                    {
                    }
                }

                // 全て失敗した場合の最終手段：より確実な方法でフォントを作成
                
                // 最終手段のゴシック系フォント（異なるエンコーディングで試行）
                var finalCandidates = new[]
                {
                    // 1. Shift_JISエンコーディングでのゴシックフォント
                    new { Path = @"C:\Windows\Fonts\YuGothR.ttc,0", Name = "Yu Gothic (90ms-RKSJ-H)", Encoding = "90ms-RKSJ-H", Embedded = BaseFont.NOT_EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\msgothic.ttc,0", Name = "MS Gothic (90ms-RKSJ-H)", Encoding = "90ms-RKSJ-H", Embedded = BaseFont.NOT_EMBEDDED },
                    
                    // 2. 埋め込み版ゴシックフォント
                    new { Path = @"C:\Windows\Fonts\YuGothR.ttc,0", Name = "Yu Gothic (Embedded)", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\msgothic.ttc,0", Name = "MS Gothic (Embedded)", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.EMBEDDED },
                    new { Path = @"C:\Windows\Fonts\BIZ-UDGothicR.ttc,0", Name = "BIZ UD Gothic (Embedded)", Encoding = BaseFont.IDENTITY_H, Embedded = BaseFont.EMBEDDED },
                    
                    // 3. 縦書きエンコーディングでのゴシックフォント
                    new { Path = @"C:\Windows\Fonts\YuGothR.ttc,0", Name = "Yu Gothic (IDENTITY_V)", Encoding = BaseFont.IDENTITY_V, Embedded = BaseFont.NOT_EMBEDDED },
                    
                    // 4. UTF-16エンコーディングでのゴシックフォント
                    new { Path = @"C:\Windows\Fonts\YuGothR.ttc,0", Name = "Yu Gothic (UniJIS)", Encoding = "UniJIS-UTF16-H", Embedded = BaseFont.NOT_EMBEDDED },
                    
                    // 5. 最もシンプルなゴシックフォント
                    new { Path = @"C:\Windows\Fonts\NotoSansJP-VF.ttf", Name = "Noto Sans JP (Simple)", Encoding = "CP1252", Embedded = BaseFont.NOT_EMBEDDED },
                };
                
                foreach (var candidate in finalCandidates)
                {
                    try
                    {
                        
                        var font = BaseFont.CreateFont(candidate.Path, candidate.Encoding, candidate.Embedded);
                        _cachedJapaneseFont = font;
                        _cachedFontName = preferredFont;
                        return font;
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }
                }
                
                // 完全に失敗した場合はHelveticaを使用
                var helveticaFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
                _cachedJapaneseFont = helveticaFont;
                _cachedFontName = preferredFont;
                return helveticaFont;
            }
        }

        /// <summary>
        /// ���x�Ȃ�����@�\�F�K�w�\���ł�������쐬
        /// </summary>
        /// <param name="pdfFilePaths">��������PDF�t�@�C���p�X���X�g</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="fileItems">�t�@�C���A�C�e�����X�g</param>
        /// <param name="addPageNumber">�y�[�W�ԍ��ǉ��t���O</param>
        /// <param name="groupByFolder">�t�H���_�ʂɃO���[�v�����邩�ǂ���</param>
        /// <param name="addHeaderFooter">�w�b�_�E�t�b�^�ǉ��t���O</param>
        /// <param name="headerFooterText">�w�b�_�E�t�b�^�e�L�X�g</param>
        /// <param name="headerFooterFontSize">�w�b�_�E�t�b�^�t�H���g�T�C�Y</param>
        /// <param name="pageNumberPosition">ページ振りの位置（0:右上, 1:右下, 2:左上, 3:左下）</param>
        /// <param name="pageNumberOffsetX">ページ振りのX軸オフセット</param>
        /// <param name="pageNumberOffsetY">ページ振りのY軸オフセット</param>
        /// <param name="headerPosition">ヘッダの位置（0:左, 1:中央, 2:右）</param>
        /// <param name="headerOffsetX">ヘッダのX軸オフセット</param>
        /// <param name="headerOffsetY">ヘッダのY軸オフセット</param>
        /// <param name="footerPosition">フッタの位置（0:左, 1:中央, 2:右）</param>
        /// <param name="footerOffsetX">フッタのX軸オフセット</param>
        /// <param name="footerOffsetY">フッタのY軸オフセット</param>
        /// <param name="pageNumberFontSize">ページ番号のフォントサイズ</param>
        /// <param name="headerFontSize">ヘッダのフォントサイズ</param>
        /// <param name="footerFontSize">フッタのフォントサイズ</param>
        public static void MergePdfFilesWithAdvancedBookmarks(List<string> pdfFilePaths, string outputPath, List<FileItem> fileItems, bool addPageNumber = false, bool groupByFolder = false, bool addHeaderFooter = false, string headerFooterText = "", float headerFooterFontSize = 10.0f,
            int pageNumberPosition = 0, float pageNumberOffsetX = 20.0f, float pageNumberOffsetY = 20.0f, float pageNumberFontSize = 10.0f,
            int headerPosition = 0, float headerOffsetX = 20.0f, float headerOffsetY = 20.0f, float headerFontSize = 10.0f,
            int footerPosition = 2, float footerOffsetX = 20.0f, float footerOffsetY = 20.0f, float footerFontSize = 10.0f,
            bool addHeader = false, bool addFooter = false, string headerText = "", string footerText = "", string headerFooterFont = "MS Gothic")
        {
            using (var document = new Document())
            using (var copy = new PdfCopy(document, new FileStream(outputPath, FileMode.Create)))
            {
                document.Open();

                var bookmarks = new List<Dictionary<string, object>>();
                int currentPage = 1;

                if (groupByFolder)
                {
                    // �t�H���_�ʂɃO���[�v������������\�����쐬
                    CreateFolderGroupedBookmarks(pdfFilePaths, fileItems, copy, bookmarks, ref currentPage);
                }
                else
                {
                    // �ʏ�̃t���b�g�Ȃ�����\�����쐬
                    CreateFlatBookmarks(pdfFilePaths, fileItems, copy, bookmarks, ref currentPage);
                }

                // �������ݒ�
                if (bookmarks.Any())
                {
                    copy.Outlines = bookmarks;
                }
            }

            // �y�[�W�ԍ��܂��̓w�b�_�E�t�b�^�ǉ��i�I�v�V�����j
            if (addPageNumber || addHeaderFooter || addHeader || addFooter)
            {
                AddPageNumbersAndHeaderFooter(outputPath, addPageNumber, addHeaderFooter, headerFooterText, headerFooterFontSize,
                    pageNumberPosition, pageNumberOffsetX, pageNumberOffsetY, pageNumberFontSize,
                    headerPosition, headerOffsetX, headerOffsetY, headerFontSize,
                    footerPosition, footerOffsetX, footerOffsetY, footerFontSize,
                    addHeader, addFooter, headerText, footerText, headerFooterFont);
            }
        }

        /// <summary>
        /// �t���b�g�Ȃ�����\�����쐬
        /// </summary>
        private static void CreateFlatBookmarks(List<string> pdfFilePaths, List<FileItem> fileItems, PdfCopy copy, List<Dictionary<string, object>> bookmarks, ref int currentPage)
        {
            for (int fileIndex = 0; fileIndex < pdfFilePaths.Count; fileIndex++)
            {
                var pdfPath = pdfFilePaths[fileIndex];
                var startPage = currentPage;

                using (var reader = new PdfReader(pdfPath))
                {
                    // PDF�̃y�[�W��ǉ�
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        var page = copy.GetImportedPage(reader, i);
                        copy.AddPage(page);
                    }

                    // ���������ǉ�
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
        /// �t�H���_�ʂɃO���[�v������������\�����쐬
        /// </summary>
        private static void CreateFolderGroupedBookmarks(List<string> pdfFilePaths, List<FileItem> fileItems, PdfCopy copy, List<Dictionary<string, object>> bookmarks, ref int currentPage)
        {
            // �t�H���_�ʂɃt�@�C�����O���[�v��
            var folderGroups = fileItems?
                .Select((item, index) => new { Item = item, Index = index, PdfPath = pdfFilePaths[index] })
                .GroupBy(x => x.Item.FolderName ?? "���[�g")
                .OrderBy(g => g.Key)
                .ToList();

            if (folderGroups == null || !folderGroups.Any())
            {
                // �t�H���_��񂪂Ȃ��ꍇ�̓t���b�g�\���Ƀt�H�[���o�b�N
                CreateFlatBookmarks(pdfFilePaths, fileItems, copy, bookmarks, ref currentPage);
                return;
            }

            foreach (var folderGroup in folderGroups)
            {
                var folderStartPage = currentPage;
                var childBookmarks = new List<Dictionary<string, object>>();

                // �t�H���_���̊e�t�@�C��������
                foreach (var fileInfo in folderGroup.OrderBy(x => x.Item.DisplayOrder))
                {
                    var startPage = currentPage;

                    using (var reader = new PdfReader(fileInfo.PdfPath))
                    {
                        // PDF�̃y�[�W��ǉ�
                        for (int i = 1; i <= reader.NumberOfPages; i++)
                        {
                            var page = copy.GetImportedPage(reader, i);
                            copy.AddPage(page);
                        }

                        // �q���������ǉ�
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

                // �t�H���_���x���̂�������쐬
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