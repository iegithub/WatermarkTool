using System;
using System.IO;
using System.Text.Json;
using WatermarkTool.Models;

namespace WatermarkTool.Services
{
    /// <summary>
    /// 设置持久化服务 - 保存/加载用户设置到JSON文件
    /// </summary>
    public static class SettingsService
    {
        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "WatermarkTool",
            "settings.json"
        );

        /// <summary>
        /// 保存设置到文件
        /// </summary>
        public static bool SaveSettings(WatermarkSettings settings, List<KeywordRule> keywordRules)
        {
            try
            {
                var serial = settings.ToSerializable();
                serial.KeywordRules = keywordRules.Select(r => new KeywordRuleData
                {
                    Keyword = r.Keyword,
                    WatermarkText = r.WatermarkText
                }).ToList();

                var directory = Path.GetDirectoryName(SettingsFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var json = JsonSerializer.Serialize(serial, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                File.WriteAllText(SettingsFilePath, json);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存设置失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 从文件加载设置
        /// </summary>
        public static (WatermarkSettings? settings, List<KeywordRule> keywordRules) LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsFilePath))
                {
                    return (null, new List<KeywordRule>());
                }

                var json = File.ReadAllText(SettingsFilePath);
                var serial = JsonSerializer.Deserialize<SerializableSettings>(json);

                if (serial == null)
                {
                    return (null, new List<KeywordRule>());
                }

                var settings = WatermarkSettings.FromSerializable(serial);
                var keywordRules = serial.KeywordRules.Select(r => new KeywordRule
                {
                    Keyword = r.Keyword,
                    WatermarkText = r.WatermarkText
                }).ToList();

                return (settings, keywordRules);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载设置失败: {ex.Message}");
                return (null, new List<KeywordRule>());
            }
        }

        /// <summary>
        /// 获取设置文件路径
        /// </summary>
        public static string GetSettingsPath()
        {
            return SettingsFilePath;
        }

        /// <summary>
        /// 检查设置文件是否存在
        /// </summary>
        public static bool SettingsExists()
        {
            return File.Exists(SettingsFilePath);
        }
    }
}
