namespace WatermarkTool.Models
{
    /// <summary>
    /// 关键字匹配规则
    /// </summary>
    public class KeywordRule
    {
        /// <summary>关键字（用于匹配文件名）</summary>
        public string Keyword { get; set; } = string.Empty;

        /// <summary>匹配时使用的水印文字（如果为空则使用关键字本身）</summary>
        public string WatermarkText { get; set; } = string.Empty;

        /// <summary>实际获取水印文字</summary>
        public string GetWatermarkText()
        {
            return string.IsNullOrWhiteSpace(WatermarkText) ? Keyword : WatermarkText;
        }
    }
}
