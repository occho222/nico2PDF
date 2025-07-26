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
    /// PDF�ϊ��T�[�r�X
    /// </summary>
    public class PdfConversionService
    {
        // Windows API�錾�iPowerPoint�̊��S�Ȕ�\�����p�j
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // ShowWindow�p�̒萔
        private const int SW_HIDE = 0;
        private const int SW_MINIMIZE = 6;

        // SetWindowPos�p�̒萔
        private const uint SWP_HIDEWINDOW = 0x0080;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;

        /// <summary>
        /// Office������PDF�ɕϊ�
        /// </summary>
        /// <param name="filePath">�ϊ����t�@�C���p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        /// <param name="baseFolderPath">��t�H���_�p�X�i�T�u�t�H���_�\���ێ��p�j</param>
        /// <param name="maintainSubfolderStructure">�T�u�t�H���_�\�����ێ����邩�ǂ���</param>
        public static void ConvertToPdf(string filePath, string pdfOutputFolder, string targetPages = "", 
            string baseFolderPath = "", bool maintainSubfolderStructure = false)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            
            string outputPath;
            if (maintainSubfolderStructure && !string.IsNullOrEmpty(baseFolderPath))
            {
                // �T�u�t�H���_�\�����ێ�
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
                // �]���ʂ�A���ׂē����t�H���_�ɏo��
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
                    // PDF�t�@�C���̏ꍇ�A�y�[�W�w�肪����Β��o�A�Ȃ���΃R�s�[
                    ProcessPdfFile(filePath, outputPath, targetPages);
                    break;
                default:
                    throw new NotSupportedException($"�Ή����Ă��Ȃ��t�@�C���`��: {extension}");
            }
        }

        /// <summary>
        /// Office������PDF�ɕϊ��i�]���̌݊������\�b�h�j
        /// </summary>
        /// <param name="filePath">�ϊ����t�@�C���p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        public static void ConvertToPdf(string filePath, string pdfOutputFolder, string targetPages = "")
        {
            ConvertToPdf(filePath, pdfOutputFolder, targetPages, "", false);
        }

        /// <summary>
        /// Excel��PDF�ϊ�
        /// </summary>
        /// <param name="inputPath">���̓t�@�C���p�X</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        private static void ConvertExcelToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? excelApp = null;
            dynamic? workbook = null;

            try
            {
                // ������Excel�v���Z�X�������I��
                KillExistingExcelProcesses();

                // Excel�A�v���P�[�V�����𓮓I�ɍ쐬
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel Application��������܂���B");
                }

                excelApp = Activator.CreateInstance(excelType);
                if (excelApp == null)
                {
                    throw new InvalidOperationException("Excel Application�̋N�����ł��܂���ł����B");
                }

                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;
                
                workbook = excelApp.Workbooks.Open(inputPath);

                var targetSheets = ParsePageRange(targetPages);

                if (targetSheets.Any())
                {
                    // 指定シートのみ変換（表示順序で処理）
                    var totalSheets = workbook.Sheets.Count;

                    // 存在しないシートのチェック
                    var invalidSheets = targetSheets.Where(s => s > totalSheets).ToList();
                    if (invalidSheets.Any())
                    {
                        throw new ArgumentException($"存在しないシート番号が指定されています: {string.Join(", ", invalidSheets)} (総シート数: {totalSheets})");
                    }

                    // 指定シートを表示順序で選択
                    for (int i = 0; i < targetSheets.Count; i++)
                    {
                        var sheetIndex = targetSheets[i];
                        var sheet = workbook.Sheets[sheetIndex]; // 表示順序でシートを取得
                        
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
                    // �S�V�[�g�ϊ� (xlTypePDF = 0)
                    workbook.ExportAsFixedFormat(0, outputPath);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Excel�ϊ����ɃG���[���������܂���: {ex.Message}", ex);
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
                    Debug.WriteLine($"Excel�I�������ŃG���[: {ex.Message}");
                }

                if (workbook != null) ReleaseComObject(workbook);
                if (excelApp != null) ReleaseComObject(excelApp);

                // �����K�x�[�W�R���N�V����
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // �c���v���Z�X�������I��
                KillExistingExcelProcesses();
            }
        }

        /// <summary>
        /// Word��PDF�ϊ�
        /// </summary>
        /// <param name="inputPath">���̓t�@�C���p�X</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        private static void ConvertWordToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? wordApp = null;
            dynamic? document = null;

            try
            {
                // ������Word�v���Z�X���N���[���A�b�v
                KillExistingWordProcesses();

                // Word�A�v���P�[�V�����𓮓I�ɍ쐬
                var wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    throw new InvalidOperationException("Word Application��������܂���B");
                }

                wordApp = Activator.CreateInstance(wordType);
                if (wordApp == null)
                {
                    throw new InvalidOperationException("Word Application�̋N�����ł��܂���ł����B");
                }

                wordApp.Visible = false;
                wordApp.DisplayAlerts = 0; // wdAlertsNone = 0
                document = wordApp.Documents.Open(inputPath);

                var targetPageList = ParsePageRange(targetPages);

                if (targetPageList.Any())
                {
                    // �w��y�[�W�̂ݕϊ�
                    var totalPages = document.ComputeStatistics(4); // wdStatisticPages = 4

                    // ���݂��Ȃ��y�[�W�̃`�F�b�N
                    var invalidPages = targetPageList.Where(p => p > totalPages).ToList();
                    if (invalidPages.Any())
                    {
                        throw new ArgumentException($"���݂��Ȃ��y�[�W�ԍ����w�肳��Ă��܂�: {string.Join(", ", invalidPages)} (���y�[�W��: {totalPages})");
                    }

                    // �y�[�W�͈͎w���PDF�o��
                    // wdExportFormatPDF = 17, wdExportFromTo = 3
                    document.ExportAsFixedFormat(outputPath, 17, Range: 3, From: targetPageList.Min(), To: targetPageList.Max());
                }
                else
                {
                    // �S�y�[�W�ϊ�
                    // wdExportFormatPDF = 17
                    document.ExportAsFixedFormat(outputPath, 17);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Word�ϊ����ɃG���[���������܂���: {ex.Message}", ex);
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
                    Debug.WriteLine($"Word�I�������ŃG���[: {ex.Message}");
                }

                if (document != null) ReleaseComObject(document);
                if (wordApp != null) ReleaseComObject(wordApp);

                // �����K�x�[�W�R���N�V����
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // �c���v���Z�X�������I��
                KillExistingWordProcesses();
            }
        }

        /// <summary>
        /// PDF�t�@�C���̏����i�y�[�W�w�肪����Β��o�A�Ȃ���΃R�s�[�j
        /// </summary>
        /// <param name="inputPath">����PDF�t�@�C���p�X</param>
        /// <param name="outputPath">�o��PDF�t�@�C���p�X</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        private static void ProcessPdfFile(string inputPath, string outputPath, string targetPages = "")
        {
            var targetPageList = ParsePageRange(targetPages);

            if (targetPageList.Any())
            {
                // �w��y�[�W�̂ݒ��o
                using (var inputReader = new PdfReader(inputPath))
                {
                    var totalPages = inputReader.NumberOfPages;

                    // ���݂��Ȃ��y�[�W�̃`�F�b�N
                    var invalidPages = targetPageList.Where(p => p > totalPages).ToList();
                    if (invalidPages.Any())
                    {
                        throw new ArgumentException($"���݂��Ȃ��y�[�W�ԍ����w�肳��Ă��܂�: {string.Join(", ", invalidPages)} (���y�[�W��: {totalPages})");
                    }

                    // �w�肳�ꂽ�y�[�W�݂̂𒊏o����PDF���쐬
                    ExtractPdfPages(inputPath, outputPath, targetPageList);
                }
            }
            else
            {
                // �S�y�[�W�R�s�[�i�]���̓���j
                File.Copy(inputPath, outputPath, overwrite: true); // �K���㏑��
            }
        }

        /// <summary>
        /// PowerPoint��PDF�ϊ�
        /// </summary>
        /// <param name="inputPath">���̓t�@�C���p�X</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        private static void ConvertPowerPointToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? pptApp = null;
            dynamic? presentation = null;
            string? tempPdfPath = null;

            try
            {
                // ������PowerPoint�v���Z�X�������I��
                KillExistingPowerPointProcesses();

                var pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    throw new InvalidOperationException("PowerPoint�A�v���P�[�V������������܂���B");
                }
                
                pptApp = Activator.CreateInstance(pptType);
                if (pptApp == null)
                {
                    throw new InvalidOperationException("PowerPoint�A�v���P�[�V�����̋N�����ł��܂���ł����B");
                }

                // PowerPoint�p�̃o�b�N�O���E���h�����ݒ�i���S�ȕ��@�j
                SetPowerPointBackgroundMode(pptApp);
                
                presentation = pptApp.Presentations.Open(inputPath);

                // �v���[���e�[�V�����J������ɍēx��\���������s
                System.Threading.Thread.Sleep(200); // �����ҋ@
                HidePowerPointWindows(pptApp);

                var targetSlides = ParsePageRange(targetPages);
                var totalSlides = presentation.Slides.Count;

                if (targetSlides.Any())
                {
                    // ���݂��Ȃ��X���C�h�̃`�F�b�N
                    var invalidSlides = targetSlides.Where(s => s > totalSlides).ToList();
                    if (invalidSlides.Any())
                    {
                        throw new ArgumentException($"���݂��Ȃ��X���C�h�ԍ����w�肳��Ă��܂�: {string.Join(", ", invalidSlides)} (���X���C�h��: {totalSlides})");
                    }

                    // �ꎞ�I�ɑS�X���C�h��PDF�ɕϊ�
                    tempPdfPath = Path.Combine(Path.GetTempPath(), $"temp_ppt_{Guid.NewGuid()}.pdf");
                    
                    // PDF�ϊ��O�ɍēx��\����
                    HidePowerPointWindows(pptApp);
                    
                    presentation.SaveAs(tempPdfPath, 32); // 32 = ppSaveAsPDF

                    // PowerPoint�����
                    presentation.Close();
                    pptApp.Quit();
                    ReleaseComObject(presentation);
                    ReleaseComObject(pptApp);
                    presentation = null;
                    pptApp = null;

                    // �ꎞ�I��GC���s
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // �w�肳�ꂽ�y�[�W�݂̂𒊏o����PDF���쐬
                    ExtractPdfPages(tempPdfPath, outputPath, targetSlides);
                }
                else
                {
                    // �S�X���C�h�ϊ��O�ɍēx��\����
                    HidePowerPointWindows(pptApp);
                    
                    // �S�X���C�h�ϊ�
                    presentation.SaveAs(outputPath, 32); // 32 = ppSaveAsPDF
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"PowerPoint�ϊ����ɃG���[���������܂���: {ex.Message}", ex);
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
                    Debug.WriteLine($"PowerPoint�I�������ŃG���[: {ex.Message}");
                }
                
                if (presentation != null) ReleaseComObject(presentation);
                if (pptApp != null) ReleaseComObject(pptApp);

                // �ꎞ�t�@�C�����폜
                if (!string.IsNullOrEmpty(tempPdfPath) && File.Exists(tempPdfPath))
                {
                    try
                    {
                        File.Delete(tempPdfPath);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"�ꎞ�t�@�C���폜�Ɏ��s: {ex.Message}");
                    }
                }

                // �����K�x�[�W�R���N�V�����iExcel��Word�Ɠ��l�j
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // �c���v���Z�X�������I��
                KillExistingPowerPointProcesses();
            }
        }

        /// <summary>
        /// PowerPoint���o�b�N�O���E���h���[�h�ɐݒ�
        /// </summary>
        /// <param name="pptApp">PowerPoint�A�v���P�[�V�����I�u�W�F�N�g</param>
        private static void SetPowerPointBackgroundMode(dynamic pptApp)
        {
            try
            {
                // 1. ��{�I�ȃv���p�e�B�ݒ�
                try 
                { 
                    // DisplayAlerts���ŏ��ɐݒ�i�d�v�j
                    pptApp.DisplayAlerts = 2; // ppAlertsNone = 2 
                }
                catch (Exception ex) 
                { 
                    Debug.WriteLine($"DisplayAlerts�ݒ�G���[: {ex.Message}"); 
                }

                try 
                { 
                    // ScreenUpdating�𖳌���
                    pptApp.ScreenUpdating = false; 
                }
                catch (Exception ex) 
                { 
                    Debug.WriteLine($"ScreenUpdating�ݒ�G���[: {ex.Message}"); 
                }

                try 
                { 
                    // Visible�v���p�e�B��ݒ�
                    pptApp.Visible = false; 
                }
                catch (Exception ex) 
                { 
                    Debug.WriteLine($"Visible�ݒ�G���[: {ex.Message}"); 
                }

                // 2. Windows API���g�p�������S�Ȕ�\����
                HidePowerPointWindows(pptApp);

                // 3. �����ҋ@���Ă���ēx��\���������s
                System.Threading.Thread.Sleep(100);
                HidePowerPointWindows(pptApp);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"PowerPoint�o�b�N�O���E���h�ݒ�̈�ʓI�ȃG���[: {ex.Message}");
                // PowerPoint�̔w�i�ݒ�ŃG���[���������Ă������𑱍s
            }
        }

        /// <summary>
        /// Windows API���g�p����PowerPoint�E�B���h�E�����S�ɔ�\���ɂ���
        /// </summary>
        /// <param name="pptApp">PowerPoint�A�v���P�[�V�����I�u�W�F�N�g</param>
        private static void HidePowerPointWindows(dynamic pptApp)
        {
            try
            {
                // PowerPoint�v���Z�X�݂̂�ΏۂƂ������S�Ȕ�\����
                var processes = Process.GetProcessesByName("POWERPNT");
                foreach (var process in processes)
                {
                    try
                    {
                        // ���C���E�B���h�E�n���h�����擾
                        var mainWindowHandle = process.MainWindowHandle;
                        if (mainWindowHandle != IntPtr.Zero)
                        {
                            // �E�B���h�E���\���ɂ���
                            ShowWindow(mainWindowHandle, SW_HIDE);
                            
                            // ����ɃE�B���h�E���A�N�e�B�u�Ȉʒu�Ɉړ�
                            SetWindowPos(mainWindowHandle, IntPtr.Zero, -32000, -32000, 0, 0, 
                                SWP_HIDEWINDOW | SWP_NOACTIVATE | SWP_NOSIZE);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"PowerPoint�v���Z�X�����G���[: {ex.Message}");
                    }
                }

                // PowerPoint�̓���̃E�B���h�E�N���X�݂̂�ΏۂƂ�����\����
                try
                {
                    // PowerPoint�̃��C���E�B���h�E�N���X������
                    var pptMainWindow = FindWindow("PP12FrameClass", null);
                    if (pptMainWindow != IntPtr.Zero)
                    {
                        ShowWindow(pptMainWindow, SW_HIDE);
                    }

                    // �X���C�h�V���[�E�B���h�E������
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

                // COM�o�R�ł̃E�B���h�E����i���S�ȕ��@�j
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
                    Debug.WriteLine($"PowerPoint COM �E�B���h�E����G���[: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"PowerPointウィンドウ非表示処理の一般的なエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// ������Excel�v���Z�X�������I��
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
                
                // �����ҋ@
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Excel�v���Z�X�I�������ŃG���[: {ex.Message}");
            }
        }

        /// <summary>
        /// ������Word�v���Z�X�������I��
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
                
                // �����ҋ@
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Word�v���Z�X�I�������ŃG���[: {ex.Message}");
            }
        }

        /// <summary>
        /// ������PowerPoint�v���Z�X�������I��
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
                
                // �����ҋ@
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"PowerPoint�v���Z�X�I�������ŃG���[: {ex.Message}");
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

        /// <summary>
        /// COM�I�u�W�F�N�g�����
        /// </summary>
        /// <param name="obj">����ΏۃI�u�W�F�N�g</param>
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
                    Debug.WriteLine($"COM�I�u�W�F�N�g����G���[: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// PDF�t�@�C������w�肳�ꂽ�y�[�W�݂̂𒊏o�i���J���\�b�h�j
        /// </summary>
        /// <param name="inputPdfPath">����PDF�t�@�C���p�X</param>
        /// <param name="outputPdfPath">�o��PDF�t�@�C���p�X</param>
        /// <param name="pageRange">�y�[�W�͈͕�����i��: "1,3,5-7"�j</param>
        public static void ExtractPdfPagesFromRange(string inputPdfPath, string outputPdfPath, string pageRange)
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
        /// ���΃p�X���擾
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