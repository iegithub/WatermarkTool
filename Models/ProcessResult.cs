namespace WatermarkTool.Models
{
    /// <summary>
    /// 批量处理结果
    /// </summary>
    public class ProcessResult
    {
        public string FilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string WatermarkText { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public string MatchedKeyword { get; set; } = string.Empty;
    }
}
