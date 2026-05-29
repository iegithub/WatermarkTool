using System.Drawing;

namespace WatermarkTool.Models
{
    /// <summary>
    /// 水印配置模型
    /// </summary>
    public class WatermarkSettings
    {
        /// <summary>水印文字</summary>
        public string Text { get; set; } = "水印";

        /// <summary>水印颜色</summary>
        public Color Color { get; set; } = Color.FromArgb(128, 200, 200, 200);

        /// <summary>字体大小（磅）</summary>
        public float FontSize { get; set; } = 48f;

        /// <summary>字体名称</summary>
        public string FontFamily { get; set; } = "微软雅黑";

        /// <summary>水印样式</summary>
        public WatermarkStyle Style { get; set; } = WatermarkStyle.ArtisticText;

        /// <summary>水印位置（页面中的预设位置）</summary>
        public WatermarkPosition Position { get; set; } = WatermarkPosition.Center;

        /// <summary>自定义X坐标（百分比 0-100，相对于页面宽度）</summary>
        public float CustomX { get; set; } = 50f;

        /// <summary>自定义Y坐标（百分比 0-100，相对于页面高度）</summary>
        public float CustomY { get; set; } = 50f;

        /// <summary>旋转角度（度）</summary>
        public float Rotation { get; set; } = -30f;

        /// <summary>透明度（0-255）</summary>
        public int Opacity { get; set; } = 128;

        /// <summary>是否使用自定义位置</summary>
        public bool UseCustomPosition { get; set; } = false;
    }

    /// <summary>
    /// 水印样式枚举
    /// </summary>
    public enum WatermarkStyle
    {
        /// <summary>艺术字效果（带阴影、轮廓）</summary>
        ArtisticText,
        /// <summary>普通半透明文字</summary>
        SemiTransparent
    }

    /// <summary>
    /// 水印位置枚举
    /// </summary>
    public enum WatermarkPosition
    {
        Center,
        TopLeft,
        TopCenter,
        TopRight,
        MiddleLeft,
        MiddleRight,
        BottomLeft,
        BottomCenter,
        BottomRight,
        Custom
    }
}
