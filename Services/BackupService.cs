using System;
using System.IO;

namespace WatermarkTool.Services
{
    /// <summary>
    /// 文件备份服务 - 保持源目录结构进行备份
    /// </summary>
    public static class BackupService
    {
        /// <summary>
        /// 备份文件，保持原有目录结构
        /// </summary>
        /// <param name="sourceFilePath">源文件路径</param>
        /// <param name="sourceRootPath">源根目录（用户选择的文件夹）</param>
        /// <returns>备份文件路径，失败返回null</returns>
        public static string? BackupFile(string sourceFilePath, string sourceRootPath)
        {
            try
            {
                // 获取源根目录的父目录，用于创建备份根目录
                var sourceRootDir = new DirectoryInfo(sourceRootPath);
                var parentDir = sourceRootDir.Parent;
                
                if (parentDir == null)
                {
                    // 如果源根目录没有父目录（如磁盘根目录），则在同级创建备份
                    parentDir = sourceRootDir;
                }
                
                // 创建备份根目录名称：原文件夹名_backup
                var backupRootName = $"{sourceRootDir.Name}_backup";
                var backupRootPath = Path.Combine(parentDir.FullName, backupRootName);
                
                // 计算相对路径，保持目录结构
                var relativePath = GetRelativePath(sourceFilePath, sourceRootPath);
                var backupFilePath = Path.Combine(backupRootPath, relativePath);
                
                // 确保备份目录存在
                var backupDir = Path.GetDirectoryName(backupFilePath);
                if (!string.IsNullOrEmpty(backupDir))
                {
                    Directory.CreateDirectory(backupDir);
                }
                
                // 处理文件冲突：如果备份文件已存在，添加时间戳重命名
                backupFilePath = GetUniqueBackupPath(backupFilePath);
                
                // 复制文件
                File.Copy(sourceFilePath, backupFilePath, overwrite: false);
                
                return backupFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"备份失败: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// 获取相对路径
        /// </summary>
        private static string GetRelativePath(string fullPath, string basePath)
        {
            // 标准化路径分隔符
            fullPath = Path.GetFullPath(fullPath);
            basePath = Path.GetFullPath(basePath);
            
            // 如果 basePath 是文件，取其目录
            if (File.Exists(basePath))
            {
                basePath = Path.GetDirectoryName(basePath) ?? basePath;
            }
            
            // 确保 basePath 以分隔符结尾，便于正确处理
            if (!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                basePath += Path.DirectorySeparatorChar;
            }
            
            if (fullPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
            {
                return fullPath.Substring(basePath.Length);
            }
            
            // 如果无法获取相对路径，返回文件名
            return Path.GetFileName(fullPath);
        }
        
        /// <summary>
        /// 获取唯一的备份路径（处理冲突）
        /// </summary>
        private static string GetUniqueBackupPath(string backupFilePath)
        {
            if (!File.Exists(backupFilePath))
            {
                return backupFilePath;
            }
            
            // 文件已存在，添加时间戳重命名
            var directory = Path.GetDirectoryName(backupFilePath) ?? "";
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(backupFilePath);
            var extension = Path.GetExtension(backupFilePath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            
            var newFileName = $"{fileNameWithoutExt}.{timestamp}{extension}";
            return Path.Combine(directory, newFileName);
        }
        
        /// <summary>
        /// 批量备份多个文件
        /// </summary>
        /// <param name="sourceFilePaths">源文件路径列表</param>
        /// <param name="sourceRootPath">源根目录</param>
        /// <returns>成功备份的文件数</returns>
        public static int BackupFiles(string[] sourceFilePaths, string sourceRootPath)
        {
            int successCount = 0;
            
            foreach (var filePath in sourceFilePaths)
            {
                var backupPath = BackupFile(filePath, sourceRootPath);
                if (backupPath != null)
                {
                    successCount++;
                }
            }
            
            return successCount;
        }
    }
}
