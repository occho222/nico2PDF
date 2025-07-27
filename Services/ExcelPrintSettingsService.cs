using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Nico2PDF.Models;

namespace Nico2PDF.Services
{
    /// <summary>
    /// Excel印刷設定サービス
    /// </summary>
    public static class ExcelPrintSettingsService
    {

        /// <summary>
        /// Excelファイルに印刷設定を適用
        /// </summary>
        /// <param name="filePath">Excelファイルパス</param>
        /// <param name="settings">印刷設定</param>
        public static void ApplyPrintSettings(string filePath, PrintSettingsItem settings)
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

                // ワークブックを開く
                workbook = excelApp.Workbooks.Open(filePath);

                // 全てのワークシートに印刷設定を適用
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    var worksheet = workbook.Worksheets[i];
                    ApplyPrintSettingsToWorksheet(worksheet, settings);
                }

                // ワークブックを保存
                workbook.Save();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Excel印刷設定の適用でエラーが発生しました: {ex.Message}", ex);
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
                    Debug.WriteLine($"Excel終了でエラー: {ex.Message}");
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
        /// ワークシートに印刷設定を適用
        /// </summary>
        /// <param name="worksheet">ワークシート</param>
        /// <param name="settings">印刷設定</param>
        private static void ApplyPrintSettingsToWorksheet(dynamic worksheet, PrintSettingsItem settings)
        {
            try
            {
                var pageSetup = worksheet.PageSetup;
                var appliedSettings = new System.Collections.Generic.List<string>();

                // 変更された設定のみを適用
                if (settings.IsPaperSizeChanged && settings.PaperSize.HasValue)
                {
                    pageSetup.PaperSize = GetExcelPaperSize(settings.PaperSize.Value);
                    appliedSettings.Add("用紙サイズ");
                }

                if (settings.IsOrientationChanged && settings.Orientation.HasValue)
                {
                    pageSetup.Orientation = GetExcelOrientation(settings.Orientation.Value);
                    appliedSettings.Add("用紙の向き");
                }

                if (settings.IsFitToPageOptionChanged && settings.FitToPageOption.HasValue)
                {
                    ApplyFitToPageSettings(pageSetup, settings.FitToPageOption.Value);
                    appliedSettings.Add("印刷範囲");
                }

                if (appliedSettings.Count > 0)
                {
                    Debug.WriteLine($"印刷設定適用: {worksheet.Name} - {string.Join(", ", appliedSettings)}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ワークシート印刷設定エラー: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// ページに合わせる設定を適用
        /// </summary>
        /// <param name="pageSetup">PageSetupオブジェクト</param>
        /// <param name="fitToPageOption">ページ設定オプション</param>
        private static void ApplyFitToPageSettings(dynamic pageSetup, FitToPageOption fitToPageOption)
        {
            switch (fitToPageOption)
            {
                case FitToPageOption.FitSheetOnOnePage:
                    // シートを1ページに印刷
                    pageSetup.Zoom = false;
                    pageSetup.FitToPagesWide = 1;
                    pageSetup.FitToPagesTall = 1;
                    break;

                case FitToPageOption.FitAllColumnsOnOnePage:
                    // 全ての列を1ページに印刷
                    pageSetup.Zoom = false;
                    pageSetup.FitToPagesWide = 1;
                    pageSetup.FitToPagesTall = false; // 行は制限しない
                    break;

                case FitToPageOption.FitAllRowsOnOnePage:
                    // 全ての行を1ページに印刷
                    pageSetup.Zoom = false;
                    pageSetup.FitToPagesWide = false; // 列は制限しない
                    pageSetup.FitToPagesTall = 1;
                    break;

                case FitToPageOption.None:
                default:
                    // 標準（100%表示）
                    pageSetup.Zoom = 100;
                    pageSetup.FitToPagesWide = false;
                    pageSetup.FitToPagesTall = false;
                    break;
            }
        }

        /// <summary>
        /// 用紙サイズをExcel定数に変換
        /// </summary>
        /// <param name="paperSize">用紙サイズ</param>
        /// <returns>Excel用紙サイズ定数</returns>
        private static int GetExcelPaperSize(PaperSize paperSize)
        {
            return paperSize switch
            {
                PaperSize.A3 => 8,   // xlPaperA3
                PaperSize.A4 => 9,   // xlPaperA4
                _ => 9               // デフォルトはA4
            };
        }

        /// <summary>
        /// 用紙の向きをExcel定数に変換
        /// </summary>
        /// <param name="orientation">用紙の向き</param>
        /// <returns>Excel用紙の向き定数</returns>
        private static int GetExcelOrientation(Models.Orientation orientation)
        {
            return orientation switch
            {
                Models.Orientation.Landscape => 2,  // xlLandscape
                Models.Orientation.Portrait => 1,   // xlPortrait
                _ => 1                               // デフォルトは縦
            };
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

                // 待機時間
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Excelプロセス終了でエラー: {ex.Message}");
            }
        }


        /// <summary>
        /// COMオブジェクト解放
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
    }
}