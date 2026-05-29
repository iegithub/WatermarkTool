using System.Collections.Generic;
using System.IO;
using System.Linq;
using WatermarkTool.Models;

namespace WatermarkTool.Services
{
    /// <summary>
    /// 关键字匹配服务 - 根据文件名匹配水印文字
    /// </summary>
    public static class KeywordMatcher
    {
        /// <summary>
        /// 根据文件名匹配关键字，返回匹配的水印文字
        /// </summary>
        /// <param name="fileName">文件名（不含路径）</param>
        /// <param name="keywordRules">关键字规则列表</param>
        /// <param name="defaultText">默认水印文字（无匹配时使用）</param>
        /// <returns>匹配结果：水印文字和匹配到的关键字</returns>
        public static (string watermarkText, string matchedKeyword) Match(string fileName, List<KeywordRule> keywordRules, string defaultText)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return (defaultText, string.Empty);

            var nameOnly = Path.GetFileNameWithoutExtension(fileName);

            foreach (var rule in keywordRules)
            {
                if (string.IsNullOrWhiteSpace(rule.Keyword))
                    continue;

                if (nameOnly.Contains(rule.Keyword, StringComparison.OrdinalIgnoreCase))
                {
                    return (rule.GetWatermarkText(), rule.Keyword);
                }
            }

            return (defaultText, string.Empty);
        }
    }
}
