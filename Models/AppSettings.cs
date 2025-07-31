using System;
using System.IO;
using System.Text.Json;

namespace Nico2PDF.Models
{
    /// <summary>
    /// アプリケーション設定を管理するクラス
    /// </summary>
    public class AppSettings
    {
        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Nico2PDF",
            "settings.json"
        );

        /// <summary>
        /// メイン画面のヘッダテキスト
        /// </summary>
        public string HeaderText { get; set; } = "";

        /// <summary>
        /// メイン画面のフッタテキスト
        /// </summary>
        public string FooterText { get; set; } = "";

        /// <summary>
        /// ヘッダ追加フラグ
        /// </summary>
        public bool AddHeader { get; set; } = false;

        /// <summary>
        /// フッタ追加フラグ
        /// </summary>
        public bool AddFooter { get; set; } = false;

        /// <summary>
        /// 結合ファイル名
        /// </summary>
        public string MergeFileName { get; set; } = "結合PDF";

        /// <summary>
        /// ページ番号追加フラグ
        /// </summary>
        public bool AddPageNumber { get; set; } = false;

        /// <summary>
        /// しおり追加フラグ
        /// </summary>
        public bool AddBookmarks { get; set; } = true;

        /// <summary>
        /// フォルダ別グループ化フラグ
        /// </summary>
        public bool GroupByFolder { get; set; } = false;

        /// <summary>
        /// 設定をファイルから読み込む
        /// </summary>
        /// <returns>読み込まれた設定</returns>
        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    var jsonString = File.ReadAllText(SettingsFilePath);
                    var settings = JsonSerializer.Deserialize<AppSettings>(jsonString);
                    return settings ?? new AppSettings();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"設定の読み込みに失敗しました: {ex.Message}");
            }

            return new AppSettings();
        }

        /// <summary>
        /// 設定をファイルに保存する
        /// </summary>
        public void Save()
        {
            try
            {
                var directory = Path.GetDirectoryName(SettingsFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                var jsonString = JsonSerializer.Serialize(this, options);
                File.WriteAllText(SettingsFilePath, jsonString);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"設定の保存に失敗しました: {ex.Message}");
            }
        }

        /// <summary>
        /// MainWindowから設定を更新する
        /// </summary>
        /// <param name="headerText">ヘッダテキスト</param>
        /// <param name="footerText">フッタテキスト</param>
        /// <param name="addHeader">ヘッダ追加フラグ</param>
        /// <param name="addFooter">フッタ追加フラグ</param>
        /// <param name="mergeFileName">結合ファイル名</param>
        /// <param name="addPageNumber">ページ番号追加フラグ</param>
        /// <param name="addBookmarks">しおり追加フラグ</param>
        /// <param name="groupByFolder">フォルダ別グループ化フラグ</param>
        public void UpdateFromMainWindow(string headerText, string footerText, bool addHeader, bool addFooter,
            string mergeFileName, bool addPageNumber, bool addBookmarks, bool groupByFolder)
        {
            HeaderText = headerText;
            FooterText = footerText;
            AddHeader = addHeader;
            AddFooter = addFooter;
            MergeFileName = mergeFileName;
            AddPageNumber = addPageNumber;
            AddBookmarks = addBookmarks;
            GroupByFolder = groupByFolder;
        }
    }
}