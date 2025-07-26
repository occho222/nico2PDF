using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Nico2PDF.Models;

namespace Nico2PDF.Services
{
    /// <summary>
    /// ファイル管理サービス
    /// </summary>
    public class FileManagementService
    {
        /// <summary>
        /// 指定フォルダからファイルを読み込み
        /// </summary>
        /// <param name="folderPath">フォルダパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>ファイルアイテムリスト</returns>
        public static List<FileItem> LoadFilesFromFolder(string folderPath, string pdfOutputFolder, bool includeSubfolders = false)
        {
            var fileItems = new List<FileItem>();
            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };

            var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            foreach (var ext in extensions)
            {
                var files = Directory.GetFiles(folderPath, ext, searchOption);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    string extensionUpper = fileInfo.Extension.TrimStart('.').ToUpper();
                    
                    var item = new FileItem
                    {
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        Extension = extensionUpper,
                        LastModified = fileInfo.LastWriteTime,
                        IsSelected = true,
                        PdfStatus = CheckPdfExists(fileInfo, pdfOutputFolder, folderPath, includeSubfolders) ? "変換済" : "未変換",
                        TargetPages = GetDefaultTargetPages(extensionUpper),
                        RelativePath = GetRelativePath(folderPath, fileInfo.FullName),
                        OriginalFileName = fileInfo.Name
                    };
                    fileItems.Add(item);
                }
            }

            return fileItems.OrderBy(f => f.RelativePath).ThenBy(f => f.FileName).ToList();
        }

        /// <summary>
        /// 指定フォルダからファイルを読み込み（従来の互換性メソッド）
        /// </summary>
        /// <param name="folderPath">フォルダパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <returns>ファイルアイテムリスト</returns>
        public static List<FileItem> LoadFilesFromFolder(string folderPath, string pdfOutputFolder)
        {
            return LoadFilesFromFolder(folderPath, pdfOutputFolder, false);
        }

        /// <summary>
        /// ファイルの更新をチェック
        /// </summary>
        /// <param name="folderPath">フォルダパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="currentFileItems">現在のファイルアイテムリスト</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>更新されたファイルアイテムリスト</returns>
        public static (List<FileItem> UpdatedItems, List<string> ChangedFiles, List<string> AddedFiles, List<string> DeletedFiles) 
            UpdateFiles(string folderPath, string pdfOutputFolder, List<FileItem> currentFileItems, bool includeSubfolders = false)
        {
            var previousFiles = currentFileItems.ToDictionary(f => f.FilePath, f => f);
            var newFileItems = new List<FileItem>();
            var changedFiles = new List<string>();
            var addedFiles = new List<string>();
            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };

            var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            // 現在のファイルシステムからファイルを読み込み
            foreach (var ext in extensions)
            {
                var files = Directory.GetFiles(folderPath, ext, searchOption);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    string extensionUpper = fileInfo.Extension.TrimStart('.').ToUpper();

                    bool isSelected = true;
                    string targetPages = GetDefaultTargetPages(extensionUpper);
                    int displayOrder = 0;
                    string displayName = "";
                    string originalFileName = fileInfo.Name;

                    // 既存ファイルの場合は更新日時をチェック
                    if (previousFiles.TryGetValue(file, out var existingFile))
                    {
                        if (existingFile.LastModified != fileInfo.LastWriteTime)
                        {
                            // 更新日時が変更された場合
                            changedFiles.Add(fileInfo.Name);
                            isSelected = true;
                        }
                        else
                        {
                            // 変更されていない場合は前の選択状態を保持
                            isSelected = existingFile.IsSelected;
                            targetPages = existingFile.TargetPages;
                            displayOrder = existingFile.DisplayOrder;
                            displayName = existingFile.DisplayName;
                            originalFileName = existingFile.OriginalFileName;
                        }
                        
                        // 処理済みファイルを辞書から削除
                        previousFiles.Remove(file);
                    }
                    else
                    {
                        // 新規ファイルかもしれないが、ファイル名が変更された可能性もチェック
                        var matchingItem = currentFileItems.FirstOrDefault(item => 
                            Path.GetDirectoryName(item.FilePath) == Path.GetDirectoryName(file) &&
                            (item.OriginalFileName == fileInfo.Name || 
                             Path.GetFileNameWithoutExtension(item.OriginalFileName) == Path.GetFileNameWithoutExtension(fileInfo.Name)));

                        if (matchingItem != null)
                        {
                            // 既存アイテムの設定を継承（名前変更されたファイル）
                            isSelected = matchingItem.IsSelected;
                            targetPages = matchingItem.TargetPages;
                            displayOrder = matchingItem.DisplayOrder;
                            displayName = matchingItem.DisplayName;
                            originalFileName = matchingItem.OriginalFileName;
                            
                            // 古いPDFファイルを削除（名前変更されたファイルの場合）
                            if (matchingItem.FilePath != file && matchingItem.Extension.ToUpper() != "PDF")
                            {
                                var oldPdfPath = GetPdfPath(matchingItem.FilePath, pdfOutputFolder, folderPath, includeSubfolders);
                                if (File.Exists(oldPdfPath))
                                {
                                    try
                                    {
                                        File.Delete(oldPdfPath);
                                        System.Diagnostics.Debug.WriteLine($"古いPDFファイルを削除: {oldPdfPath}");
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"古いPDFファイル削除エラー: {ex.Message}");
                                    }
                                }
                            }
                            
                            // 元のアイテムを辞書から削除（重複を避けるため）
                            previousFiles.Remove(matchingItem.FilePath);
                        }
                        else
                        {
                            // 真の新規ファイルの場合
                            addedFiles.Add(fileInfo.Name);
                            isSelected = true;
                            displayOrder = currentFileItems.Count + addedFiles.Count - 1;
                            originalFileName = fileInfo.Name;
                        }
                    }

                    // PDFステータスを確認
                    bool pdfExists = CheckPdfExists(fileInfo, pdfOutputFolder, folderPath, includeSubfolders);
                    string pdfStatus = pdfExists ? "変換済" : "未変換";

                    // PDFが存在しない場合は選択状態にする（PDFファイル自体は除く）
                    if (!pdfExists && fileInfo.Extension.ToLower() != ".pdf")
                    {
                        isSelected = true;
                    }

                    var item = new FileItem
                    {
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        Extension = extensionUpper,
                        LastModified = fileInfo.LastWriteTime,
                        IsSelected = isSelected,
                        PdfStatus = pdfStatus,
                        TargetPages = targetPages,
                        DisplayOrder = displayOrder,
                        RelativePath = GetRelativePath(folderPath, fileInfo.FullName),
                        DisplayName = displayName,
                        OriginalFileName = originalFileName
                    };
                    newFileItems.Add(item);
                }
            }

            // 削除されたファイルを検出
            var deletedFiles = new List<string>();

            foreach (var previousFile in previousFiles.Values)
            {
                // ファイルが名前変更されていないかチェック
                var directory = Path.GetDirectoryName(previousFile.FilePath);
                var matchingNewFile = newFileItems.FirstOrDefault(item => 
                    Path.GetDirectoryName(item.FilePath) == directory &&
                    (item.OriginalFileName == previousFile.OriginalFileName ||
                     Path.GetFileNameWithoutExtension(previousFile.OriginalFileName) == Path.GetFileNameWithoutExtension(previousFile.OriginalFileName)));

                if (matchingNewFile == null)
                {
                    // 真に削除されたファイル
                    deletedFiles.Add(previousFile.FileName);

                    // 対応するPDFファイルを削除
                    if (previousFile.Extension.ToUpper() != "PDF")
                    {
                        var pdfPath = GetPdfPath(previousFile.FilePath, pdfOutputFolder, folderPath, includeSubfolders);
                        if (File.Exists(pdfPath))
                        {
                            try
                            {
                                File.Delete(pdfPath);
                                System.Diagnostics.Debug.WriteLine($"削除されたファイルのPDFを削除: {pdfPath}");
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"PDFファイル削除エラー: {ex.Message}");
                            }
                        }
                    }
                }
            }

            // 表示順序で並び替え
            var orderedItems = newFileItems.OrderBy(f => f.DisplayOrder).ThenBy(f => f.RelativePath).ThenBy(f => f.FileName).ToList();
            
            // 番号を再設定
            for (int i = 0; i < orderedItems.Count; i++)
            {
                orderedItems[i].Number = i + 1;
                orderedItems[i].DisplayOrder = i;
            }

            return (orderedItems, changedFiles, addedFiles, deletedFiles);
        }

        /// <summary>
        /// ファイルの更新をチェック（従来の互換性メソッド）
        /// </summary>
        /// <param name="folderPath">フォルダパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="currentFileItems">現在のファイルアイテムリスト</param>
        /// <returns>更新されたファイルアイテムリスト</returns>
        public static (List<FileItem> UpdatedItems, List<string> ChangedFiles, List<string> AddedFiles, List<string> DeletedFiles) 
            UpdateFiles(string folderPath, string pdfOutputFolder, List<FileItem> currentFileItems)
        {
            return UpdateFiles(folderPath, pdfOutputFolder, currentFileItems, false);
        }

        /// <summary>
        /// ファイル名を変更
        /// </summary>
        /// <param name="filePath">元のファイルパス</param>
        /// <param name="newFileName">新しいファイル名（拡張子なし）</param>
        /// <returns>新しいファイルパス</returns>
        public static string RenamePhysicalFile(string filePath, string newFileName)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("ファイルが見つかりません。", filePath);

            var directory = Path.GetDirectoryName(filePath);
            var extension = Path.GetExtension(filePath);
            var newFilePath = Path.Combine(directory!, newFileName + extension);

            if (File.Exists(newFilePath) && !string.Equals(newFilePath, filePath, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("同じ名前のファイルが既に存在します。");

            File.Move(filePath, newFilePath);
            return newFilePath;
        }

        /// <summary>
        /// ファイル名変更時に古いPDFファイルを削除し、新しいPDFファイルの存在を確認
        /// </summary>
        /// <param name="oldFilePath">変更前のファイルパス</param>
        /// <param name="newFilePath">変更後のファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="baseFolderPath">基準フォルダパス</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>新しいPDFファイルが存在するかどうか</returns>
        public static bool HandlePdfFileAfterRename(string oldFilePath, string newFilePath, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            try
            {
                var oldFileInfo = new FileInfo(oldFilePath);
                var newFileInfo = new FileInfo(newFilePath);
                
                // PDFファイル自体の場合は処理しない
                if (oldFileInfo.Extension.ToUpper() == "PDF")
                {
                    return true;
                }

                // 古いPDFファイルのパスを取得
                var oldPdfPath = GetPdfPath(oldFilePath, pdfOutputFolder, baseFolderPath, includeSubfolders);
                
                // 古いPDFファイルが存在する場合は削除
                if (File.Exists(oldPdfPath))
                {
                    try
                    {
                        File.Delete(oldPdfPath);
                        System.Diagnostics.Debug.WriteLine($"名前変更に伴う古いPDFファイル削除: {oldPdfPath}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"古いPDFファイル削除エラー: {ex.Message}");
                    }
                }

                // 新しいPDFファイルの存在を確認
                var newPdfPath = GetPdfPath(newFilePath, pdfOutputFolder, baseFolderPath, includeSubfolders);
                return File.Exists(newPdfPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PDF処理エラー: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 複数ファイルの名前を手動で一括変更
        /// </summary>
        /// <param name="renameItems">リネームアイテムリスト</param>
        /// <returns>変更結果</returns>
        public static (int SuccessCount, int FailCount, List<string> Errors) BatchRenameFiles(List<BatchRenameItem> renameItems)
        {
            var successCount = 0;
            var failCount = 0;
            var errors = new List<string>();

            // 変更されたファイルのみを処理
            var changedItems = renameItems.Where(item => item.IsChanged && !item.HasError).ToList();

            foreach (var renameItem in changedItems)
            {
                try
                {
                    var originalItem = renameItem.OriginalItem;
                    var oldFilePath = originalItem.FilePath;
                    var newFilePath = RenamePhysicalFile(originalItem.FilePath, renameItem.NewFileName);
                    
                    // ファイルアイテムを更新
                    originalItem.FilePath = newFilePath;
                    originalItem.FileName = Path.GetFileName(newFilePath);
                    originalItem.DisplayName = renameItem.PreviewFileName;
                    
                    successCount++;
                }
                catch (Exception ex)
                {
                    errors.Add($"{renameItem.CurrentFileName}: {ex.Message}");
                    failCount++;
                }
            }

            return (successCount, failCount, errors);
        }

        /// <summary>
        /// 複数ファイルの名前を一括変更（旧版・互換性のため残す）
        /// </summary>
        /// <param name="fileItems">ファイルアイテムリスト</param>
        /// <param name="pattern">変更パターン（例：新しい名前_{0}）</param>
        /// <param name="physicalRename">物理ファイル名も変更するかどうか</param>
        /// <returns>変更結果</returns>
        public static (int SuccessCount, int FailCount, List<string> Errors) BatchRenameFiles(List<FileItem> fileItems, string pattern, bool physicalRename = false)
        {
            var successCount = 0;
            var failCount = 0;
            var errors = new List<string>();

            for (int i = 0; i < fileItems.Count; i++)
            {
                var item = fileItems[i];
                try
                {
                    var newName = string.Format(pattern, i + 1, item.FileName, Path.GetFileNameWithoutExtension(item.FileName));
                    
                    if (physicalRename)
                    {
                        var newFilePath = RenamePhysicalFile(item.FilePath, newName);
                        item.FilePath = newFilePath;
                        item.FileName = Path.GetFileName(newFilePath);
                    }
                    
                    item.DisplayName = newName + Path.GetExtension(item.FileName);
                    successCount++;
                }
                catch (Exception ex)
                {
                    errors.Add($"{item.FileName}: {ex.Message}");
                    failCount++;
                }
            }

            return (successCount, failCount, errors);
        }

        /// <summary>
        /// PDFファイルの存在を確認
        /// </summary>
        /// <param name="fileInfo">ファイル情報</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="baseFolderPath">基準フォルダパス</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>PDFファイルが存在するかどうか</returns>
        private static bool CheckPdfExists(FileInfo fileInfo, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            if (fileInfo.Extension.ToLower() == ".pdf") return true;

            var pdfPath = GetPdfPath(fileInfo.FullName, pdfOutputFolder, baseFolderPath, includeSubfolders);
            return File.Exists(pdfPath);
        }

        /// <summary>
        /// PDFファイルのパスを取得
        /// </summary>
        /// <param name="originalFilePath">元のファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="baseFolderPath">基準フォルダパス</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>PDFファイルのパス</returns>
        private static string GetPdfPath(string originalFilePath, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            var fileInfo = new FileInfo(originalFilePath);
            var fileName = Path.GetFileNameWithoutExtension(fileInfo.Name) + ".pdf";

            if (includeSubfolders)
            {
                // サブフォルダ構造を維持
                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                return Path.Combine(outputDir, fileName);
            }
            else
            {
                // すべてのファイルを同じフォルダに出力
                return Path.Combine(pdfOutputFolder, fileName);
            }
        }

        /// <summary>
        /// 相対パスを取得
        /// </summary>
        /// <param name="basePath">基準パス</param>
        /// <param name="fullPath">完全パス</param>
        /// <returns>相対パス</returns>
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

        /// <summary>
        /// 拡張子に基づいてデフォルトの対象ページを取得
        /// </summary>
        /// <param name="extension">拡張子</param>
        /// <returns>デフォルトの対象ページ</returns>
        private static string GetDefaultTargetPages(string extension)
        {
            return extension switch
            {
                "XLS" or "XLSX" or "XLSM" => "1-1",
                _ => ""
            };
        }

        /// <summary>
        /// サブフォルダ用のPDF出力ディレクトリを作成
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="baseFolderPath">基準フォルダパス</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        public static void EnsurePdfOutputDirectory(string filePath, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            if (includeSubfolders)
            {
                var fileInfo = new FileInfo(filePath);
                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            else
            {
                if (!Directory.Exists(pdfOutputFolder))
                {
                    Directory.CreateDirectory(pdfOutputFolder);
                }
            }
        }

        /// <summary>
        /// ファイル一覧をテキストファイルに書き出し
        /// </summary>
        /// <param name="fileItems">ファイルアイテムリスト</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="includeHeaders">ヘッダーを含めるかどうか</param>
        /// <param name="includeDetails">詳細情報を含めるかどうか</param>
        /// <param name="selectedOnly">選択されたファイルのみを出力するかどうか</param>
        public static void ExportFileList(List<FileItem> fileItems, string outputPath, bool includeHeaders = true, bool includeDetails = true, bool selectedOnly = false)
        {
            var exportItems = selectedOnly ? fileItems.Where(f => f.IsSelected).ToList() : fileItems;
            var sb = new StringBuilder();

            if (includeHeaders)
            {
                if (includeDetails)
                {
                    sb.AppendLine("番号\t表示名\tファイル名\t拡張子\t最終更新日時\tPDFステータス\t対象ページ\tフォルダ\tファイルパス");
                }
                else
                {
                    sb.AppendLine("番号\t表示名\tフォルダ");
                }
            }

            foreach (var item in exportItems.OrderBy(f => f.DisplayOrder))
            {
                if (includeDetails)
                {
                    sb.AppendLine($"{item.Number}\t{item.DisplayName}\t{item.FileName}\t{item.Extension}\t{item.LastModified:yyyy/MM/dd HH:mm:ss}\t{item.PdfStatus}\t{item.TargetPages}\t{item.FolderName}\t{item.FilePath}");
                }
                else
                {
                    sb.AppendLine($"{item.Number}\t{item.DisplayName}\t{item.FolderName}");
                }
            }

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// ファイル一覧をCSVファイルに書き出し
        /// </summary>
        /// <param name="fileItems">ファイルアイテムリスト</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="includeHeaders">ヘッダーを含めるかどうか</param>
        /// <param name="includeDetails">詳細情報を含めるかどうか</param>
        /// <param name="selectedOnly">選択されたファイルのみを出力するかどうか</param>
        public static void ExportFileListToCsv(List<FileItem> fileItems, string outputPath, bool includeHeaders = true, bool includeDetails = true, bool selectedOnly = false)
        {
            var exportItems = selectedOnly ? fileItems.Where(f => f.IsSelected).ToList() : fileItems;
            var sb = new StringBuilder();

            if (includeHeaders)
            {
                if (includeDetails)
                {
                    sb.AppendLine("番号,表示名,ファイル名,拡張子,最終更新日時,PDFステータス,対象ページ,フォルダ,ファイルパス");
                }
                else
                {
                    sb.AppendLine("番号,表示名,フォルダ");
                }
            }

            foreach (var item in exportItems.OrderBy(f => f.DisplayOrder))
            {
                if (includeDetails)
                {
                    sb.AppendLine($"{item.Number},\"{EscapeCsvField(item.DisplayName)}\",\"{EscapeCsvField(item.FileName)}\",{item.Extension},{item.LastModified:yyyy/MM/dd HH:mm:ss},{item.PdfStatus},\"{EscapeCsvField(item.TargetPages)}\",\"{EscapeCsvField(item.FolderName)}\",\"{EscapeCsvField(item.FilePath)}\"");
                }
                else
                {
                    sb.AppendLine($"{item.Number},\"{EscapeCsvField(item.DisplayName)}\",\"{EscapeCsvField(item.FolderName)}\"");
                }
            }

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// ファイル一覧をクリップボードにコピー
        /// </summary>
        /// <param name="fileItems">ファイルアイテムリスト</param>
        /// <param name="includeHeaders">ヘッダーを含めるかどうか</param>
        /// <param name="includeDetails">詳細情報を含めるかどうか</param>
        /// <param name="selectedOnly">選択されたファイルのみを出力するかどうか</param>
        public static void CopyFileListToClipboard(List<FileItem> fileItems, bool includeHeaders = true, bool includeDetails = true, bool selectedOnly = false)
        {
            var exportItems = selectedOnly ? fileItems.Where(f => f.IsSelected).ToList() : fileItems;
            var sb = new StringBuilder();

            if (includeHeaders)
            {
                if (includeDetails)
                {
                    sb.AppendLine("番号\t表示名\tファイル名\t拡張子\t最終更新日時\tPDFステータス\t対象ページ\tフォルダ\tファイルパス");
                }
                else
                {
                    sb.AppendLine("番号\t表示名\tフォルダ");
                }
            }

            foreach (var item in exportItems.OrderBy(f => f.DisplayOrder))
            {
                if (includeDetails)
                {
                    sb.AppendLine($"{item.Number}\t{item.DisplayName}\t{item.FileName}\t{item.Extension}\t{item.LastModified:yyyy/MM/dd HH:mm:ss}\t{item.PdfStatus}\t{item.TargetPages}\t{item.FolderName}\t{item.FilePath}");
                }
                else
                {
                    sb.AppendLine($"{item.Number}\t{item.DisplayName}\t{item.FolderName}");
                }
            }

            System.Windows.Clipboard.SetText(sb.ToString());
        }

        /// <summary>
        /// ファイル名のみをクリップボードにコピー（改行区切り）
        /// </summary>
        /// <param name="fileItems">ファイルアイテムリスト</param>
        /// <param name="selectedOnly">選択されたファイルのみを出力するかどうか</param>
        public static void CopyFileNamesToClipboard(List<FileItem> fileItems, bool selectedOnly = false)
        {
            var exportItems = selectedOnly ? fileItems.Where(f => f.IsSelected).ToList() : fileItems;
            var sb = new StringBuilder();

            foreach (var item in exportItems.OrderBy(f => f.DisplayOrder))
            {
                sb.AppendLine(item.FileName);
            }

            System.Windows.Clipboard.SetText(sb.ToString());
        }

        /// <summary>
        /// CSVフィールドをエスケープ
        /// </summary>
        /// <param name="field">フィールド値</param>
        /// <returns>エスケープされたフィールド値</returns>
        private static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";
            
            // ダブルクォートをエスケープ
            return field.Replace("\"", "\"\"");
        }
    }
}