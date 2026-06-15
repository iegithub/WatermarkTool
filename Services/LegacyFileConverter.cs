using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WatermarkTool.Models;

using WinFont = System.Drawing.Font;
using WinColor = System.Drawing.Color;

namespace WatermarkTool.Services
{
    /// <summary>
    /// Office 文件格式转换服务
    /// 将 .doc/.xls 老版本格式转换为 .docx/.xlsx
    /// </summary>
    public static class LegacyFileConverter
    {
        /// <summary>
        /// 检查是否为老版本格式
        /// </summary>
        public static bool IsLegacyFormat(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            return ext == ".doc" || ext == ".xls";
        }

        /// <summary>
        /// 将老版本 Office 文件转换为新格式
        /// 优先使用 Office COM 自动化，失败则提示用户手动转换
        /// </summary>
        public static ConvertResult ConvertToNewFormat(string filePath)
        {
            try
            {
                var ext = Path.GetExtension(filePath).ToLowerInvariant();
                string newExt = ext == ".doc" ? ".docx" : ".xlsx";
                string dir = Path.GetDirectoryName(filePath) ?? "";
                string newPath = Path.Combine(dir, Path.GetFileNameWithoutExtension(filePath) + newExt);

                // 如果目标文件已存在，先删除
                if (File.Exists(newPath))
                    File.Delete(newPath);

                bool success = false;
                string? error = null;

                if (ext == ".doc")
                {
                    success = ConvertDocToDocx(filePath, newPath, out error);
                }
                else if (ext == ".xls")
                {
                    success = ConvertXlsToXlsx(filePath, newPath, out error);
                }

                return new ConvertResult
                {
                    Success = success,
                    ConvertedPath = success ? newPath : null,
                    ErrorMessage = error,
                    OriginalPath = filePath
                };
            }
            catch (Exception ex)
            {
                return new ConvertResult
                {
                    Success = false,
                    ConvertedPath = null,
                    ErrorMessage = $"转换失败: {ex.Message}",
                    OriginalPath = filePath
                };
            }
        }

        /// <summary>
        /// 使用 Word COM 自动化将 .doc 转换为 .docx
        /// </summary>
        private static bool ConvertDocToDocx(string sourcePath, string targetPath, out string? error)
        {
            error = null;
            try
            {
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    error = "未检测到 Microsoft Word，无法自动转换。\n请手动用 Word 打开文件并另存为 .docx 格式。";
                    return false;
                }

                dynamic wordApp = Activator.CreateInstance(wordType)!;
                try
                {
                    wordApp.Visible = false;
                    wordApp.DisplayAlerts = 0; // wdAlertsNone

                    dynamic doc = wordApp.Documents.Open(sourcePath, ReadOnly: true, AddToRecentFiles: false);
                    try
                    {
                        // wdFormatDocumentDefault = 12 (docx)
                        const int wdFormatDocumentDefault = 12;
                        doc.SaveAs2(targetPath, wdFormatDocumentDefault);
                        return true;
                    }
                    finally
                    {
                        doc.Close(0); // wdDoNotSaveChanges
                    }
                }
                finally
                {
                    wordApp.Quit(0); // wdDoNotSaveChanges
                    ReleaseComObject(wordApp);
                }
            }
            catch (Exception ex)
            {
                error = $"Word 转换失败: {ex.Message}\n请手动用 Word 打开文件并另存为 .docx 格式。";
                return false;
            }
        }

        /// <summary>
        /// 使用 Excel COM 自动化将 .xls 转换为 .xlsx
        /// </summary>
        private static bool ConvertXlsToXlsx(string sourcePath, string targetPath, out string? error)
        {
            error = null;
            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    error = "未检测到 Microsoft Excel，无法自动转换。\n请手动用 Excel 打开文件并另存为 .xlsx 格式。";
                    return false;
                }

                dynamic excelApp = Activator.CreateInstance(excelType)!;
                try
                {
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;

                    dynamic workbook = excelApp.Workbooks.Open(sourcePath, ReadOnly: true);
                    try
                    {
                        // xlOpenXMLWorkbook = 51 (xlsx)
                        const int xlOpenXMLWorkbook = 51;
                        workbook.SaveAs(targetPath, xlOpenXMLWorkbook);
                        return true;
                    }
                    finally
                    {
                        workbook.Close(false);
                    }
                }
                finally
                {
                    excelApp.Quit();
                    ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                error = $"Excel 转换失败: {ex.Message}\n请手动用 Excel 打开文件并另存为 .xlsx 格式。";
                return false;
            }
        }

        /// <summary>
        /// 释放 COM 对象
        /// </summary>
        private static void ReleaseComObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch { }
        }
    }

    /// <summary>
    /// 转换结果
    /// </summary>
    public class ConvertResult
    {
        public bool Success { get; set; }
        public string? ConvertedPath { get; set; }
        public string? ErrorMessage { get; set; }
        public string OriginalPath { get; set; } = "";
    }
}
