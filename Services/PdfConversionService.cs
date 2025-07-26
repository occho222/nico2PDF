using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Threading.Tasks;
using System.Diagnostics;
using Nico2PDF.Models;
using System.Runtime.InteropServices;

namespace Nico2PDF.Services
{
    /// <summary>
    /// PDF変換サービス
    /// </summary>
    public class PdfConversionService
    {
        // Windows API宣言（PowerPointの完全な非表示化用）
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // ShowWindow用の定数
        private const int SW_HIDE = 0;
        private const int SW_MINIMIZE = 6;

        // SetWindowPos用の定数
        private const uint SWP_HIDEWINDOW = 0x0080;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;

        /// <summary>
        /// Office文書をPDFに変換
        /// </summary>
        /// <param name="filePath">変換元ファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="targetPages">対象ページ</param>
        /// <param name="baseFolderPath">基準フォルダパス（サブフォルダ構造維持用）</param>
        /// <param name="maintainSubfolderStructure">サブフォルダ構造を維持するかどうか</param>
        public static void ConvertToPdf(string filePath, string pdfOutputFolder, string targetPages = "", 
            string baseFolderPath = "", bool maintainSubfolderStructure = false)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            
            string outputPath;
            if (maintainSubfolderStructure && !string.IsNullOrEmpty(baseFolderPath))
            {
                // サブフォルダ構造を維持
                var fileInfo = new FileInfo(filePath);
                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                
                outputPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
            }
            else
            {
                // 従来通り、すべて同じフォルダに出力
                outputPath = Path.Combine(pdfOutputFolder, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
            }

            switch (extension)
            {
                case ".xls":
                case ".xlsx":
                case ".xlsm":
                    ConvertExcelToPdf(filePath, outputPath, targetPages);
                    break;
                case ".doc":
                case ".docx":
                    ConvertWordToPdf(filePath, outputPath, targetPages);
                    break;
                case ".ppt":
                case ".pptx":
                    ConvertPowerPointToPdf(filePath, outputPath, targetPages);
                    break;
                case ".pdf":
                    // PDFファイルの場合、ページ指定があれば抽出、なければコピー
                    ProcessPdfFile(filePath, outputPath, targetPages);
                    break;
                default:
                    throw new NotSupportedException($"対応していないファイル形式: {extension}");
            }
        }

        /// <summary>
        /// Office文書をPDFに変換（従来の互換性メソッド）
        /// </summary>
        /// <param name="filePath">変換元ファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="targetPages">対象ページ</param>
        public static void ConvertToPdf(string filePath, string pdfOutputFolder, string targetPages = "")
        {
            ConvertToPdf(filePath, pdfOutputFolder, targetPages, "", false);
        }

        /// <summary>
        /// Excel→PDF変換
        /// </summary>
        /// <param name="inputPath">入力ファイルパス</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="targetPages">対象ページ</param>
        private static void ConvertExcelToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? excelApp = null;
            dynamic? workbook = null;

            try
            {
                // 既存のExcelプロセスを強制終了
                KillExistingExcelProcesses();

                // Excelアプリケーションを動的に作成
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel Applicationが見つかりません。");
                }

                excelApp = Activator.CreateInstance(excelType);
                if (excelApp == null)
                {
                    throw new InvalidOperationException("Excel Applicationの起動ができませんでした。");
                }

                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;
                
                workbook = excelApp.Workbooks.Open(inputPath);

                var targetSheets = ParsePageRange(targetPages);

                if (targetSheets.Any())
                {
                    // 指定シートのみ変換
                    var totalSheets = workbook.Worksheets.Count;

                    // 存在しないシートのチェック
                    var invalidSheets = targetSheets.Where(s => s > totalSheets).ToList();
                    if (invalidSheets.Any())
                    {
                        throw new ArgumentException($"存在しないシート番号が指定されています: {string.Join(", ", invalidSheets)} (総シート数: {totalSheets})");
                    }

                    // 指定シートを選択
                    for (int i = 0; i < targetSheets.Count; i++)
                    {
                        var sheet = workbook.Worksheets[targetSheets[i]];
                        
                        if (i == 0)
                        {
                            sheet.Select();
                        }
                        else
                        {
                            sheet.Select(false); // 追加選択
                        }
                    }

                    // 選択されたシートをPDF変換 (xlTypePDF = 0)
                    excelApp.ActiveSheet.ExportAsFixedFormat(0, outputPath);
                }
                else
                {
                    // 全シート変換 (xlTypePDF = 0)
                    workbook.ExportAsFixedFormat(0, outputPath);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Excel変換中にエラーが発生しました: {ex.Message}", ex);
            }
            finally
            {
                try
                {
                    workbook?.Close(false);
                    excelApp?.Quit();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Excel終了処理でエラー: {ex.Message}");
                }

                if (workbook != null) ReleaseComObject(workbook);
                if (excelApp != null) ReleaseComObject(excelApp);

                // 強制ガベージコレクション
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // 残存プロセスを強制終了
                KillExistingExcelProcesses();
            }
        }

        /// <summary>
        /// Word→PDF変換
        /// </summary>
        /// <param name="inputPath">入力ファイルパス</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="targetPages">対象ページ</param>
        private static void ConvertWordToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? wordApp = null;
            dynamic? document = null;

            try
            {
                // 既存のWordプロセスをクリーンアップ
                KillExistingWordProcesses();

                // Wordアプリケーションを動的に作成
                var wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    throw new InvalidOperationException("Word Applicationが見つかりません。");
                }

                wordApp = Activator.CreateInstance(wordType);
                if (wordApp == null)
                {
                    throw new InvalidOperationException("Word Applicationの起動ができませんでした。");
                }

                wordApp.Visible = false;
                wordApp.DisplayAlerts = 0; // wdAlertsNone = 0
                document = wordApp.Documents.Open(inputPath);

                var targetPageList = ParsePageRange(targetPages);

                if (targetPageList.Any())
                {
                    // 指定ページのみ変換
                    var totalPages = document.ComputeStatistics(4); // wdStatisticPages = 4

                    // 存在しないページのチェック
                    var invalidPages = targetPageList.Where(p => p > totalPages).ToList();
                    if (invalidPages.Any())
                    {
                        throw new ArgumentException($"存在しないページ番号が指定されています: {string.Join(", ", invalidPages)} (総ページ数: {totalPages})");
                    }

                    // ページ範囲指定でPDF出力
                    // wdExportFormatPDF = 17, wdExportFromTo = 3
                    document.ExportAsFixedFormat(outputPath, 17, Range: 3, From: targetPageList.Min(), To: targetPageList.Max());
                }
                else
                {
                    // 全ページ変換
                    // wdExportFormatPDF = 17
                    document.ExportAsFixedFormat(outputPath, 17);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Word変換中にエラーが発生しました: {ex.Message}", ex);
            }
            finally
            {
                try
                {
                    document?.Close(false);
                    wordApp?.Quit();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Word終了処理でエラー: {ex.Message}");
                }

                if (document != null) ReleaseComObject(document);
                if (wordApp != null) ReleaseComObject(wordApp);

                // 強制ガベージコレクション
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // 残存プロセスを強制終了
                KillExistingWordProcesses();
            }
        }

        /// <summary>
        /// PDFファイルの処理（ページ指定があれば抽出、なければコピー）
        /// </summary>
        /// <param name="inputPath">入力PDFファイルパス</param>
        /// <param name="outputPath">出力PDFファイルパス</param>
        /// <param name="targetPages">対象ページ</param>
        private static void ProcessPdfFile(string inputPath, string outputPath, string targetPages = "")
        {
            var targetPageList = ParsePageRange(targetPages);

            if (targetPageList.Any())
            {
                // 指定ページのみ抽出
                using (var inputReader = new PdfReader(inputPath))
                {
                    var totalPages = inputReader.NumberOfPages;

                    // 存在しないページのチェック
                    var invalidPages = targetPageList.Where(p => p > totalPages).ToList();
                    if (invalidPages.Any())
                    {
                        throw new ArgumentException($"存在しないページ番号が指定されています: {string.Join(", ", invalidPages)} (総ページ数: {totalPages})");
                    }

                    // 指定されたページのみを抽出してPDFを作成
                    ExtractPdfPages(inputPath, outputPath, targetPageList);
                }
            }
            else
            {
                // 全ページコピー（従来の動作）
                File.Copy(inputPath, outputPath, overwrite: true); // 必ず上書き
            }
        }

        /// <summary>
        /// PowerPoint→PDF変換
        /// </summary>
        /// <param name="inputPath">入力ファイルパス</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="targetPages">対象ページ</param>
        private static void ConvertPowerPointToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? pptApp = null;
            dynamic? presentation = null;
            string? tempPdfPath = null;

            try
            {
                // 既存のPowerPointプロセスを強制終了
                KillExistingPowerPointProcesses();

                var pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    throw new InvalidOperationException("PowerPointアプリケーションが見つかりません。");
                }
                
                pptApp = Activator.CreateInstance(pptType);
                if (pptApp == null)
                {
                    throw new InvalidOperationException("PowerPointアプリケーションの起動ができませんでした。");
                }

                // PowerPoint用のバックグラウンド処理設定（安全な方法）
                SetPowerPointBackgroundMode(pptApp);
                
                presentation = pptApp.Presentations.Open(inputPath);

                // プレゼンテーション開いた後に再度非表示化を実行
                System.Threading.Thread.Sleep(200); // 少し待機
                HidePowerPointWindows(pptApp);

                var targetSlides = ParsePageRange(targetPages);
                var totalSlides = presentation.Slides.Count;

                if (targetSlides.Any())
                {
                    // 存在しないスライドのチェック
                    var invalidSlides = targetSlides.Where(s => s > totalSlides).ToList();
                    if (invalidSlides.Any())
                    {
                        throw new ArgumentException($"存在しないスライド番号が指定されています: {string.Join(", ", invalidSlides)} (総スライド数: {totalSlides})");
                    }

                    // 一時的に全スライドをPDFに変換
                    tempPdfPath = Path.Combine(Path.GetTempPath(), $"temp_ppt_{Guid.NewGuid()}.pdf");
                    
                    // PDF変換前に再度非表示化
                    HidePowerPointWindows(pptApp);
                    
                    presentation.SaveAs(tempPdfPath, 32); // 32 = ppSaveAsPDF

                    // PowerPointを閉じる
                    presentation.Close();
                    pptApp.Quit();
                    ReleaseComObject(presentation);
                    ReleaseComObject(pptApp);
                    presentation = null;
                    pptApp = null;

                    // 一時的なGC実行
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // 指定されたページのみを抽出してPDFを作成
                    ExtractPdfPages(tempPdfPath, outputPath, targetSlides);
                }
                else
                {
                    // 全スライド変換前に再度非表示化
                    HidePowerPointWindows(pptApp);
                    
                    // 全スライド変換
                    presentation.SaveAs(outputPath, 32); // 32 = ppSaveAsPDF
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"PowerPoint変換中にエラーが発生しました: {ex.Message}", ex);
            }
            finally
            {
                try 
                { 
                    presentation?.Close();
                    pptApp?.Quit(); 
                } 
                catch (Exception ex)
                {
                    Debug.WriteLine($"PowerPoint終了処理でエラー: {ex.Message}");
                }
                
                if (presentation != null) ReleaseComObject(presentation);
                if (pptApp != null) ReleaseComObject(pptApp);

                // 一時ファイルを削除
                if (!string.IsNullOrEmpty(tempPdfPath) && File.Exists(tempPdfPath))
                {
                    try
                    {
                        File.Delete(tempPdfPath);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"一時ファイル削除に失敗: {ex.Message}");
                    }
                }

                // 強制ガベージコレクション（ExcelやWordと同様）
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // 残存プロセスを強制終了
                KillExistingPowerPointProcesses();
            }
        }

        /// <summary>
        /// PowerPointをバックグラウンドモードに設定
        /// </summary>
        /// <param name="pptApp">PowerPointアプリケーションオブジェクト</param>
        private static void SetPowerPointBackgroundMode(dynamic pptApp)
        {
            try
            {
                // 1. 基本的なプロパティ設定
                try 
                { 
                    // DisplayAlertsを最初に設定（重要）
                    pptApp.DisplayAlerts = 2; // ppAlertsNone = 2 
                }
                catch (Exception ex) 
                { 
                    Debug.WriteLine($"DisplayAlerts設定エラー: {ex.Message}"); 
                }

                try 
                { 
                    // ScreenUpdatingを無効化
                    pptApp.ScreenUpdating = false; 
                }
                catch (Exception ex) 
                { 
                    Debug.WriteLine($"ScreenUpdating設定エラー: {ex.Message}"); 
                }

                try 
                { 
                    // Visibleプロパティを設定
                    pptApp.Visible = false; 
                }
                catch (Exception ex) 
                { 
                    Debug.WriteLine($"Visible設定エラー: {ex.Message}"); 
                }

                // 2. Windows APIを使用した安全な非表示化
                HidePowerPointWindows(pptApp);

                // 3. 少し待機してから再度非表示化を試行
                System.Threading.Thread.Sleep(100);
                HidePowerPointWindows(pptApp);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"PowerPointバックグラウンド設定の一般的なエラー: {ex.Message}");
                // PowerPointの背景設定でエラーが発生しても処理を続行
            }
        }

        /// <summary>
        /// Windows APIを使用してPowerPointウィンドウを安全に非表示にする
        /// </summary>
        /// <param name="pptApp">PowerPointアプリケーションオブジェクト</param>
        private static void HidePowerPointWindows(dynamic pptApp)
        {
            try
            {
                // PowerPointプロセスのみを対象とした安全な非表示化
                var processes = Process.GetProcessesByName("POWERPNT");
                foreach (var process in processes)
                {
                    try
                    {
                        // メインウィンドウハンドルを取得
                        var mainWindowHandle = process.MainWindowHandle;
                        if (mainWindowHandle != IntPtr.Zero)
                        {
                            // ウィンドウを非表示にする
                            ShowWindow(mainWindowHandle, SW_HIDE);
                            
                            // さらにウィンドウを非アクティブな位置に移動
                            SetWindowPos(mainWindowHandle, IntPtr.Zero, -32000, -32000, 0, 0, 
                                SWP_HIDEWINDOW | SWP_NOACTIVATE | SWP_NOSIZE);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"PowerPointプロセス処理エラー: {ex.Message}");
                    }
                }

                // PowerPointの特定のウィンドウクラスのみを対象とした非表示化
                try
                {
                    // PowerPointのメインウィンドウクラスを検索
                    var pptMainWindow = FindWindow("PP12FrameClass", null);
                    if (pptMainWindow != IntPtr.Zero)
                    {
                        ShowWindow(pptMainWindow, SW_HIDE);
                    }

                    // スライドショーウィンドウを検索
                    var pptSlideWindow = FindWindow("PPSlideShowClass", null);
                    if (pptSlideWindow != IntPtr.Zero)
                    {
                        ShowWindow(pptSlideWindow, SW_HIDE);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"PowerPoint特定ウィンドウ非表示エラー: {ex.Message}");
                }

                // COM経由でのウィンドウ制御（安全な方法）
                try
                {
                    if (pptApp.Windows != null && pptApp.Windows.Count > 0)
                    {
                        for (int i = 1; i <= pptApp.Windows.Count; i++)
                        {
                            try
                            {
                                var window = pptApp.Windows[i];
                                window.Visible = false;
                                window.WindowState = 2; // ppWindowMinimized
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"PowerPointウィンドウ{i}非表示エラー: {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"PowerPoint COM ウィンドウ制御エラー: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"PowerPointウィンドウ非表示処理の一般的なエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 既存のExcelプロセスを強制終了
        /// </summary>
        private static void KillExistingExcelProcesses()
        {
            try
            {
                var processes = Process.GetProcessesByName("EXCEL");
                foreach (var proc in processes)
                {
                    try
                    {
                        if (!proc.HasExited)
                        {
                            proc.Kill();
                            proc.WaitForExit(3000);
                        }
                    }
                    catch { }
                    finally
                    {
                        proc.Dispose();
                    }
                }
                
                // 少し待機
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Excelプロセス終了処理でエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 既存のWordプロセスを強制終了
        /// </summary>
        private static void KillExistingWordProcesses()
        {
            try
            {
                var processes = Process.GetProcessesByName("WINWORD");
                foreach (var proc in processes)
                {
                    try
                    {
                        if (!proc.HasExited)
                        {
                            proc.Kill();
                            proc.WaitForExit(3000);
                        }
                    }
                    catch { }
                    finally
                    {
                        proc.Dispose();
                    }
                }
                
                // 少し待機
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Wordプロセス終了処理でエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 既存のPowerPointプロセスを強制終了
        /// </summary>
        private static void KillExistingPowerPointProcesses()
        {
            try
            {
                var processes = Process.GetProcessesByName("POWERPNT");
                foreach (var proc in processes)
                {
                    try
                    {
                        if (!proc.HasExited)
                        {
                            proc.Kill();
                            proc.WaitForExit(3000);
                        }
                    }
                    catch { }
                    finally
                    {
                        proc.Dispose();
                    }
                }
                
                // 少し待機
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"PowerPointプロセス終了処理でエラー: {ex.Message}");
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

        /// <summary>
        /// COMオブジェクトを解放
        /// </summary>
        /// <param name="obj">解放対象オブジェクト</param>
        private static void ReleaseComObject(object? obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"COMオブジェクト解放エラー: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// PDFファイルから指定されたページのみを抽出（公開メソッド）
        /// </summary>
        /// <param name="inputPdfPath">入力PDFファイルパス</param>
        /// <param name="outputPdfPath">出力PDFファイルパス</param>
        /// <param name="pageRange">ページ範囲文字列（例: "1,3,5-7"）</param>
        public static void ExtractPdfPagesFromRange(string inputPdfPath, string outputPdfPath, string pageRange)
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
        /// 相対パスを取得
        /// </summary>
        private static string GetRelativePath(string basePath, string fullPath)
        {
            var baseUri = new Uri(basePath.EndsWith(Path.DirectorySeparatorChar.ToString()) ? basePath : basePath + Path.DirectorySeparatorChar);
            var fullUri = new Uri(fullPath);
            
            if (baseUri.Scheme != fullUri.Scheme)
            {
                return fullPath;
            }

            var relativeUri = baseUri.MakeRelativeUri(fullUri);
            var relativePath = Uri.UnescapeDataString(relativeUri.ToString());
            
            return relativePath.Replace('/', Path.DirectorySeparatorChar);
        }
    }
}