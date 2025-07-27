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
    /// �t�@�C���Ǘ��T�[�r�X
    /// </summary>
    public class FileManagementService
    {
        /// <summary>
        /// �w��t�H���_����t�@�C����ǂݍ���
        /// </summary>
        /// <param name="folderPath">�t�H���_�p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
        /// <returns>�t�@�C���A�C�e�����X�g</returns>
        public static List<FileItem> LoadFilesFromFolder(string folderPath, string pdfOutputFolder, bool includeSubfolders = false, int subfolderDepth = 2)
        {
            var fileItems = new List<FileItem>();
            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };

            foreach (var ext in extensions)
            {
                var files = includeSubfolders 
                    ? GetFilesWithDepthLimit(folderPath, ext, subfolderDepth)
                    : Directory.GetFiles(folderPath, ext, SearchOption.TopDirectoryOnly);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    
                    // ~付きの中間ファイルを除外
                    if (fileInfo.Name.StartsWith("~"))
                        continue;
                    
                    string extensionUpper = fileInfo.Extension.TrimStart('.').ToUpper();
                    
                    var item = new FileItem
                    {
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        Extension = extensionUpper,
                        LastModified = fileInfo.LastWriteTime,
                        IsSelected = true,
                        PdfStatus = CheckPdfExists(fileInfo, pdfOutputFolder, folderPath, includeSubfolders) ? "�ϊ���" : "���ϊ�",
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
        /// �w��t�H���_����t�@�C����ǂݍ��݁i�]���̌݊������\�b�h�j
        /// </summary>
        /// <param name="folderPath">�t�H���_�p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <returns>�t�@�C���A�C�e�����X�g</returns>
        public static List<FileItem> LoadFilesFromFolder(string folderPath, string pdfOutputFolder)
        {
            return LoadFilesFromFolder(folderPath, pdfOutputFolder, false);
        }

        /// <summary>
        /// �t�@�C���̍X�V���`�F�b�N
        /// </summary>
        /// <param name="folderPath">�t�H���_�p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="currentFileItems">���݂̃t�@�C���A�C�e�����X�g</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
        /// <returns>�X�V���ꂽ�t�@�C���A�C�e�����X�g</returns>
        public static (List<FileItem> UpdatedItems, List<string> ChangedFiles, List<string> AddedFiles, List<string> DeletedFiles) 
            UpdateFiles(string folderPath, string pdfOutputFolder, List<FileItem> currentFileItems, bool includeSubfolders = false, int subfolderDepth = 1)
        {
            var previousFiles = currentFileItems.ToDictionary(f => f.FilePath, f => f);
            var newFileItems = new List<FileItem>();
            var changedFiles = new List<string>();
            var addedFiles = new List<string>();
            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };

            // ���݂̃t�@�C���V�X�e������t�@�C����ǂݍ���
            foreach (var ext in extensions)
            {
                var files = includeSubfolders ? 
                    GetFilesWithDepthLimit(folderPath, ext, subfolderDepth) :
                    Directory.GetFiles(folderPath, ext, SearchOption.TopDirectoryOnly);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    
                    // ~付きの中間ファイルを除外
                    if (fileInfo.Name.StartsWith("~"))
                        continue;
                    
                    string extensionUpper = fileInfo.Extension.TrimStart('.').ToUpper();

                    bool isSelected = true;
                    string targetPages = GetDefaultTargetPages(extensionUpper);
                    int displayOrder = 0;
                    string displayName = "";
                    string originalFileName = fileInfo.Name;

                    // �����t�@�C���̏ꍇ�͍X�V�������`�F�b�N
                    if (previousFiles.TryGetValue(file, out var existingFile))
                    {
                        if (existingFile.LastModified != fileInfo.LastWriteTime)
                        {
                            // �X�V�������ύX���ꂽ�ꍇ
                            changedFiles.Add(fileInfo.Name);
                            isSelected = true;
                        }
                        else
                        {
                            // �ύX����Ă��Ȃ��ꍇ�͑O�̑I����Ԃ�ێ�
                            isSelected = existingFile.IsSelected;
                            targetPages = existingFile.TargetPages;
                            displayOrder = existingFile.DisplayOrder;
                            displayName = existingFile.DisplayName;
                            originalFileName = existingFile.OriginalFileName;
                        }
                        
                        // �����ς݃t�@�C������������폜
                        previousFiles.Remove(file);
                    }
                    else
                    {
                        // �V�K�t�@�C����������Ȃ����A�t�@�C�������ύX���ꂽ�\�����`�F�b�N
                        var matchingItem = currentFileItems.FirstOrDefault(item => 
                            Path.GetDirectoryName(item.FilePath) == Path.GetDirectoryName(file) &&
                            (item.OriginalFileName == fileInfo.Name || 
                             Path.GetFileNameWithoutExtension(item.OriginalFileName) == Path.GetFileNameWithoutExtension(fileInfo.Name)));

                        if (matchingItem != null)
                        {
                            // �����A�C�e���̐ݒ���p���i���O�ύX���ꂽ�t�@�C���j
                            isSelected = matchingItem.IsSelected;
                            targetPages = matchingItem.TargetPages;
                            displayOrder = matchingItem.DisplayOrder;
                            displayName = matchingItem.DisplayName;
                            originalFileName = matchingItem.OriginalFileName;
                            
                            // �Â�PDF�t�@�C�����폜�i���O�ύX���ꂽ�t�@�C���̏ꍇ�j
                            if (matchingItem.FilePath != file && matchingItem.Extension.ToUpper() != "PDF")
                            {
                                var oldPdfPath = GetPdfPath(matchingItem.FilePath, pdfOutputFolder, folderPath, includeSubfolders);
                                if (File.Exists(oldPdfPath))
                                {
                                    try
                                    {
                                        File.Delete(oldPdfPath);
                                        System.Diagnostics.Debug.WriteLine($"�Â�PDF�t�@�C�����폜: {oldPdfPath}");
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"�Â�PDF�t�@�C���폜�G���[: {ex.Message}");
                                    }
                                }
                            }
                            
                            // ���̃A�C�e������������폜�i�d��������邽�߁j
                            previousFiles.Remove(matchingItem.FilePath);
                        }
                        else
                        {
                            // �^�̐V�K�t�@�C���̏ꍇ
                            addedFiles.Add(fileInfo.Name);
                            isSelected = true;
                            displayOrder = currentFileItems.Count + addedFiles.Count - 1;
                            originalFileName = fileInfo.Name;
                        }
                    }

                    // PDF�X�e�[�^�X���m�F
                    bool pdfExists = CheckPdfExists(fileInfo, pdfOutputFolder, folderPath, includeSubfolders);
                    string pdfStatus = pdfExists ? "�ϊ���" : "���ϊ�";

                    // PDF�����݂��Ȃ��ꍇ�͑I����Ԃɂ���iPDF�t�@�C�����̂͏����j
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

            // �폜���ꂽ�t�@�C�������o
            var deletedFiles = new List<string>();

            foreach (var previousFile in previousFiles.Values)
            {
                // 対応する新しいファイルが存在するかチェック（単純なFilePath比較）
                var matchingNewFile = newFileItems.FirstOrDefault(item => item.FilePath == previousFile.FilePath);

                if (matchingNewFile == null)
                {
                    // ファイルが削除された（または名前変更された）
                    deletedFiles.Add(previousFile.FileName);
                    System.Diagnostics.Debug.WriteLine($"削除されたファイル検出: {previousFile.FilePath}");

                    // 対応するPDFファイルを削除
                    if (previousFile.Extension.ToUpper() != "PDF")
                    {
                        var pdfPath = GetPdfPath(previousFile.FilePath, pdfOutputFolder, folderPath, includeSubfolders);
                        System.Diagnostics.Debug.WriteLine($"PDF削除試行: {pdfPath}");
                        if (File.Exists(pdfPath))
                        {
                            try
                            {
                                File.Delete(pdfPath);
                                System.Diagnostics.Debug.WriteLine($"削除されたファイルのPDF削除成功: {pdfPath}");
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"PDFファイル削除エラー: {ex.Message}");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"PDFファイルが存在しません: {pdfPath}");
                        }
                    }
                }
            }

            // �\�������ŕ��ёւ�
            var orderedItems = newFileItems.OrderBy(f => f.DisplayOrder).ThenBy(f => f.RelativePath).ThenBy(f => f.FileName).ToList();
            
            // �ԍ����Đݒ�
            for (int i = 0; i < orderedItems.Count; i++)
            {
                orderedItems[i].Number = i + 1;
                orderedItems[i].DisplayOrder = i;
            }

            return (orderedItems, changedFiles, addedFiles, deletedFiles);
        }

        /// <summary>
        /// �t�@�C���̍X�V���`�F�b�N�i�]���̌݊������\�b�h�j
        /// </summary>
        /// <param name="folderPath">�t�H���_�p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="currentFileItems">���݂̃t�@�C���A�C�e�����X�g</param>
        /// <returns>�X�V���ꂽ�t�@�C���A�C�e�����X�g</returns>
        public static (List<FileItem> UpdatedItems, List<string> ChangedFiles, List<string> AddedFiles, List<string> DeletedFiles) 
            UpdateFiles(string folderPath, string pdfOutputFolder, List<FileItem> currentFileItems)
        {
            return UpdateFiles(folderPath, pdfOutputFolder, currentFileItems, false);
        }

        /// <summary>
        /// �t�@�C������ύX
        /// </summary>
        /// <param name="filePath">���̃t�@�C���p�X</param>
        /// <param name="newFileName">�V�����t�@�C�����i�g���q�Ȃ��j</param>
        /// <returns>�V�����t�@�C���p�X</returns>
        public static string RenamePhysicalFile(string filePath, string newFileName)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("�t�@�C����������܂���B", filePath);

            var directory = Path.GetDirectoryName(filePath);
            var extension = Path.GetExtension(filePath);
            var newFilePath = Path.Combine(directory!, newFileName + extension);

            if (File.Exists(newFilePath) && !string.Equals(newFilePath, filePath, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("�������O�̃t�@�C�������ɑ��݂��܂��B");

            File.Move(filePath, newFilePath);
            return newFilePath;
        }

        /// <summary>
        /// �t�@�C�����ύX���ɌÂ�PDF�t�@�C�����폜���A�V����PDF�t�@�C���̑��݂��m�F
        /// </summary>
        /// <param name="oldFilePath">�ύX�O�̃t�@�C���p�X</param>
        /// <param name="newFilePath">�ύX��̃t�@�C���p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="baseFolderPath">��t�H���_�p�X</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
        /// <returns>�V����PDF�t�@�C�������݂��邩�ǂ���</returns>
        public static bool HandlePdfFileAfterRename(string oldFilePath, string newFilePath, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            try
            {
                var oldFileInfo = new FileInfo(oldFilePath);
                var newFileInfo = new FileInfo(newFilePath);
                
                // PDF�t�@�C�����̂̏ꍇ�͏������Ȃ�
                if (oldFileInfo.Extension.ToUpper() == "PDF")
                {
                    return true;
                }

                // �Â�PDF�t�@�C���̃p�X���擾
                var oldPdfPath = GetPdfPath(oldFilePath, pdfOutputFolder, baseFolderPath, includeSubfolders);
                
                // �Â�PDF�t�@�C�������݂���ꍇ�͍폜
                if (File.Exists(oldPdfPath))
                {
                    try
                    {
                        File.Delete(oldPdfPath);
                        System.Diagnostics.Debug.WriteLine($"���O�ύX�ɔ����Â�PDF�t�@�C���폜: {oldPdfPath}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"�Â�PDF�t�@�C���폜�G���[: {ex.Message}");
                    }
                }

                // �V����PDF�t�@�C���̑��݂��m�F
                var newPdfPath = GetPdfPath(newFilePath, pdfOutputFolder, baseFolderPath, includeSubfolders);
                return File.Exists(newPdfPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PDF�����G���[: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// �����t�@�C���̖��O���蓮�ňꊇ�ύX
        /// </summary>
        /// <param name="renameItems">���l�[���A�C�e�����X�g</param>
        /// <returns>�ύX����</returns>
        public static (int SuccessCount, int FailCount, List<string> Errors) BatchRenameFiles(List<BatchRenameItem> renameItems)
        {
            var successCount = 0;
            var failCount = 0;
            var errors = new List<string>();

            // �ύX���ꂽ�t�@�C���݂̂�����
            var changedItems = renameItems.Where(item => item.IsChanged && !item.HasError).ToList();

            foreach (var renameItem in changedItems)
            {
                try
                {
                    var originalItem = renameItem.OriginalItem;
                    var oldFilePath = originalItem.FilePath;
                    var newFilePath = RenamePhysicalFile(originalItem.FilePath, renameItem.NewFileName);
                    
                    // �t�@�C���A�C�e�����X�V
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
        /// �����t�@�C���̖��O���ꊇ�ύX�i���ŁE�݊����̂��ߎc���j
        /// </summary>
        /// <param name="fileItems">�t�@�C���A�C�e�����X�g</param>
        /// <param name="pattern">�ύX�p�^�[���i��F�V�������O_{0}�j</param>
        /// <param name="physicalRename">�����t�@�C�������ύX���邩�ǂ���</param>
        /// <returns>�ύX����</returns>
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
        /// PDF�t�@�C���̑��݂��m�F
        /// </summary>
        /// <param name="fileInfo">�t�@�C�����</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="baseFolderPath">��t�H���_�p�X</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
        /// <returns>PDF�t�@�C�������݂��邩�ǂ���</returns>
        private static bool CheckPdfExists(FileInfo fileInfo, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            if (fileInfo.Extension.ToLower() == ".pdf") return true;

            var pdfPath = GetPdfPath(fileInfo.FullName, pdfOutputFolder, baseFolderPath, includeSubfolders);
            return File.Exists(pdfPath);
        }

        /// <summary>
        /// PDF�t�@�C���̃p�X���擾
        /// </summary>
        /// <param name="originalFilePath">���̃t�@�C���p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="baseFolderPath">��t�H���_�p�X</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
        /// <returns>PDF�t�@�C���̃p�X</returns>
        public static string GetPdfPath(string originalFilePath, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            var fileInfo = new FileInfo(originalFilePath);
            var fileName = Path.GetFileNameWithoutExtension(fileInfo.Name) + ".pdf";

            if (includeSubfolders)
            {
                // �T�u�t�H���_�\�����ێ�
                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                return Path.Combine(outputDir, fileName);
            }
            else
            {
                // ���ׂẴt�@�C���𓯂��t�H���_�ɏo��
                return Path.Combine(pdfOutputFolder, fileName);
            }
        }

        /// <summary>
        /// ���΃p�X���擾
        /// </summary>
        /// <param name="basePath">��p�X</param>
        /// <param name="fullPath">���S�p�X</param>
        /// <returns>���΃p�X</returns>
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
        /// �g���q�Ɋ�Â��ăf�t�H���g�̑Ώۃy�[�W���擾
        /// </summary>
        /// <param name="extension">�g���q</param>
        /// <returns>�f�t�H���g�̑Ώۃy�[�W</returns>
        private static string GetDefaultTargetPages(string extension)
        {
            return extension switch
            {
                "XLS" or "XLSX" or "XLSM" => "1-1",
                _ => ""
            };
        }

        /// <summary>
        /// �T�u�t�H���_�p��PDF�o�̓f�B���N�g�����쐬
        /// </summary>
        /// <param name="filePath">�t�@�C���p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="baseFolderPath">��t�H���_�p�X</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
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
        /// �t�@�C���ꗗ���e�L�X�g�t�@�C���ɏ����o��
        /// </summary>
        /// <param name="fileItems">�t�@�C���A�C�e�����X�g</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="includeHeaders">�w�b�_�[���܂߂邩�ǂ���</param>
        /// <param name="includeDetails">�ڍ׏����܂߂邩�ǂ���</param>
        /// <param name="selectedOnly">�I�����ꂽ�t�@�C���݂̂��o�͂��邩�ǂ���</param>
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
        /// �t�@�C���ꗗ��CSV�t�@�C���ɏ����o��
        /// </summary>
        /// <param name="fileItems">�t�@�C���A�C�e�����X�g</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="includeHeaders">�w�b�_�[���܂߂邩�ǂ���</param>
        /// <param name="includeDetails">�ڍ׏����܂߂邩�ǂ���</param>
        /// <param name="selectedOnly">�I�����ꂽ�t�@�C���݂̂��o�͂��邩�ǂ���</param>
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
        /// �t�@�C���ꗗ���N���b�v�{�[�h�ɃR�s�[
        /// </summary>
        /// <param name="fileItems">�t�@�C���A�C�e�����X�g</param>
        /// <param name="includeHeaders">�w�b�_�[���܂߂邩�ǂ���</param>
        /// <param name="includeDetails">�ڍ׏����܂߂邩�ǂ���</param>
        /// <param name="selectedOnly">�I�����ꂽ�t�@�C���݂̂��o�͂��邩�ǂ���</param>
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
        /// �t�@�C�����݂̂��N���b�v�{�[�h�ɃR�s�[�i���s��؂�j
        /// </summary>
        /// <param name="fileItems">�t�@�C���A�C�e�����X�g</param>
        /// <param name="selectedOnly">�I�����ꂽ�t�@�C���݂̂��o�͂��邩�ǂ���</param>
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
        /// CSV�t�B�[���h���G�X�P�[�v
        /// </summary>
        /// <param name="field">�t�B�[���h�l</param>
        /// <returns>�G�X�P�[�v���ꂽ�t�B�[���h�l</returns>
        private static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";
            
            // �_�u���N�H�[�g���G�X�P�[�v
            return field.Replace("\"", "\"\"");
        }

        /// <summary>
        /// 指定された階層数の制限でファイルを取得
        /// </summary>
        /// <param name="rootPath">ルートパス</param>
        /// <param name="searchPattern">検索パターン</param>
        /// <param name="maxDepth">最大階層数</param>
        /// <returns>ファイルパスのリスト</returns>
        private static string[] GetFilesWithDepthLimit(string rootPath, string searchPattern, int maxDepth)
        {
            var files = new List<string>();
            GetFilesRecursive(rootPath, searchPattern, maxDepth, 0, files);
            return files.ToArray();
        }

        /// <summary>
        /// 再帰的にファイルを取得（階層制限付き）
        /// </summary>
        /// <param name="currentPath">現在のパス</param>
        /// <param name="searchPattern">検索パターン</param>
        /// <param name="maxDepth">最大階層数</param>
        /// <param name="currentDepth">現在の階層数</param>
        /// <param name="files">ファイルリスト</param>
        private static void GetFilesRecursive(string currentPath, string searchPattern, int maxDepth, int currentDepth, List<string> files)
        {
            try
            {
                // 現在の階層のファイルを追加
                files.AddRange(Directory.GetFiles(currentPath, searchPattern));

                // 最大階層に達していない場合、サブディレクトリを探索
                if (currentDepth < maxDepth)
                {
                    var directories = Directory.GetDirectories(currentPath);
                    foreach (var directory in directories)
                    {
                        GetFilesRecursive(directory, searchPattern, maxDepth, currentDepth + 1, files);
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                // アクセス権限がないディレクトリはスキップ
            }
            catch (DirectoryNotFoundException)
            {
                // 存在しないディレクトリはスキップ
            }
        }
    }
}