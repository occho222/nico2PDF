using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using Nico2PDF.Models;

namespace Nico2PDF.Services
{
    /// <summary>
    /// プロジェクト管理サービス
    /// </summary>
    public class ProjectManager
    {
        private static readonly string ProjectsFilePath = Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? AppDomain.CurrentDomain.BaseDirectory,
            "projects.json"
        );

        /// <summary>
        /// プロジェクトを読み込み
        /// </summary>
        /// <returns>プロジェクトリスト</returns>
        public static List<ProjectData> LoadProjects()
        {
            try
            {
                if (File.Exists(ProjectsFilePath))
                {
                    var json = File.ReadAllText(ProjectsFilePath);
                    var projects = JsonSerializer.Deserialize<List<ProjectData>>(json) ?? new List<ProjectData>();
                    return projects;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"プロジェクトの読み込みに失敗しました: {ex.Message}", "エラー",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
            return new List<ProjectData>();
        }

        /// <summary>
        /// プロジェクトを保存
        /// </summary>
        /// <param name="projects">プロジェクトリスト</param>
        public static void SaveProjects(List<ProjectData> projects)
        {
            try
            {
                var directory = Path.GetDirectoryName(ProjectsFilePath);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory!);

                var json = JsonSerializer.Serialize(projects, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(ProjectsFilePath, json);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"プロジェクトの保存に失敗しました: {ex.Message}", "エラー",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// カテゴリ別のプロジェクト表示順序を取得
        /// </summary>
        /// <param name="projects">プロジェクトリスト</param>
        /// <returns>カテゴリ別に整理されたプロジェクトリスト</returns>
        public static List<ProjectData> GetProjectsByCategoryOrder(List<ProjectData> projects)
        {
            var result = new List<ProjectData>();
            
            // カテゴリでグループ化
            var groupedProjects = projects.GroupBy(p => string.IsNullOrEmpty(p.Category) ? "未分類" : p.Category)
                                          .OrderBy(g => g.Key == "未分類" ? "z" : g.Key) // 未分類を最後に
                                          .ToList();

            foreach (var group in groupedProjects)
            {
                // カテゴリ内でプロジェクト名順に並び替え
                var categoryProjects = group.OrderBy(p => p.Name).ToList();
                result.AddRange(categoryProjects);
            }

            return result;
        }

        /// <summary>
        /// 利用可能なカテゴリ一覧を取得
        /// </summary>
        /// <param name="projects">プロジェクトリスト</param>
        /// <returns>カテゴリ一覧</returns>
        public static List<string> GetAvailableCategories(List<ProjectData> projects)
        {
            var categories = projects.Where(p => !string.IsNullOrEmpty(p.Category))
                                   .Select(p => p.Category)
                                   .Distinct()
                                   .OrderBy(c => c)
                                   .ToList();
            
            return categories;
        }

        /// <summary>
        /// プロジェクトを指定カテゴリに移動
        /// </summary>
        /// <param name="project">移動対象プロジェクト</param>
        /// <param name="newCategory">新しいカテゴリ</param>
        public static void MoveProjectToCategory(ProjectData project, string newCategory)
        {
            project.Category = newCategory ?? "";
        }
    }
}