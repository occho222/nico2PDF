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
    /// �v���W�F�N�g�Ǘ��T�[�r�X
    /// </summary>
    public class ProjectManager
    {
        private static readonly string ProjectsFilePath = Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? AppDomain.CurrentDomain.BaseDirectory,
            "projects.json"
        );

        /// <summary>
        /// �v���W�F�N�g��ǂݍ���
        /// </summary>
        /// <returns>�v���W�F�N�g���X�g</returns>
        public static List<ProjectData> LoadProjects()
        {
            try
            {
                if (File.Exists(ProjectsFilePath))
                {
                    var json = File.ReadAllText(ProjectsFilePath);
                    var projects = JsonSerializer.Deserialize<List<ProjectData>>(json) ?? new List<ProjectData>();
                    
                    // 既存プロジェクトの SubfolderDepth マイグレーション
                    foreach (var project in projects)
                    {
                        // SubfolderDepthが0の場合はデフォルト値(1)を設定
                        if (project.SubfolderDepth == 0)
                        {
                            project.SubfolderDepth = 1;
                        }
                    }
                    
                    return projects;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"�v���W�F�N�g�̓ǂݍ��݂Ɏ��s���܂���: {ex.Message}", "�G���[",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
            return new List<ProjectData>();
        }

        /// <summary>
        /// �v���W�F�N�g��ۑ�
        /// </summary>
        /// <param name="projects">�v���W�F�N�g���X�g</param>
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
                System.Windows.MessageBox.Show($"�v���W�F�N�g�̕ۑ��Ɏ��s���܂���: {ex.Message}", "�G���[",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// �J�e�S���ʂ̃v���W�F�N�g�\���������擾
        /// </summary>
        /// <param name="projects">�v���W�F�N�g���X�g</param>
        /// <returns>�J�e�S���ʂɐ������ꂽ�v���W�F�N�g���X�g</returns>
        public static List<ProjectData> GetProjectsByCategoryOrder(List<ProjectData> projects)
        {
            var result = new List<ProjectData>();
            
            // �J�e�S���ŃO���[�v��
            var groupedProjects = projects.GroupBy(p => string.IsNullOrEmpty(p.Category) ? "������" : p.Category)
                                          .OrderBy(g => g.Key == "������" ? "z" : g.Key) // �����ނ��Ō��
                                          .ToList();

            foreach (var group in groupedProjects)
            {
                // �J�e�S�����Ńv���W�F�N�g�����ɕ��ёւ�
                var categoryProjects = group.OrderBy(p => p.Name).ToList();
                result.AddRange(categoryProjects);
            }

            return result;
        }

        /// <summary>
        /// ���p�\�ȃJ�e�S���ꗗ���擾
        /// </summary>
        /// <param name="projects">�v���W�F�N�g���X�g</param>
        /// <returns>�J�e�S���ꗗ</returns>
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
        /// �v���W�F�N�g���w��J�e�S���Ɉړ�
        /// </summary>
        /// <param name="project">�ړ��Ώۃv���W�F�N�g</param>
        /// <param name="newCategory">�V�����J�e�S��</param>
        public static void MoveProjectToCategory(ProjectData project, string newCategory)
        {
            project.Category = newCategory ?? "";
        }
    }
}