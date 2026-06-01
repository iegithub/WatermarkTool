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
        public Color Color { get; set; } = Color.Red; // 默认红色

        /// <summary>字体大小（磅）</summary>
        public float FontSize { get; set; } = 8f; // 默认8磅

        /// <summary>字体名称</summary>
        public string FontFamily { get; set; } = "宋体"; // 默认宋体

        /// <summary>水印样式</summary>
        public WatermarkStyle Style { get; set; } = WatermarkStyle.SemiTransparent; // 默认半透明（无阴影）

        /// <summary>水印位置（页面中的预设位置）</summary>
        public WatermarkPosition Position { get; set; } = WatermarkPosition.Center;

        /// <summary>自定义X坐标（百分比 0-100，相对于页面宽度）</summary>
        public float CustomX { get; set; } = 50f;

        /// <summary>自定义Y坐标（百分比 0-100，相对于页面高度）</summary>
        public float CustomY { get; set; } = 50f;

        /// <summary>旋转角度（度）</summary>
        public float Rotation { get; set; } = 0f; // 默认0度

        /// <summary>透明度（0-255）</summary>
        public int Opacity { get; set; } = 128;

        /// <summary>是否使用自定义位置</summary>
        public bool UseCustomPosition { get; set; } = false;

        /// <summary>
        /// 从可序列化设置创建实例
        /// </summary>
        public static WatermarkSettings FromSerializable(SerializableSettings serial)
        {
            return new WatermarkSettings
            {
                Text = serial.Text,
                Color = Color.FromArgb(serial.ColorA, serial.ColorR, serial.ColorG, serial.ColorB),
                FontSize = serial.FontSize,
                FontFamily = serial.FontFamily,
                Style = serial.Style,
                Position = serial.Position,
                CustomX = serial.CustomX,
                CustomY = serial.CustomY,
                Rotation = serial.Rotation,
                Opacity = serial.Opacity,
                UseCustomPosition = serial.UseCustomPosition
            };
        }

        /// <summary>
        /// 转換為可序列化設置
        /// </summary>
        public SerializableSettings ToSerializable()
        {
            return new SerializableSettings
            {
                Text = this.Text,
                ColorR = this.Color.R,
                ColorG = this.Color.G,
                ColorB = this.Color.B,
                ColorA = this.Color.A,
                FontSize = this.FontSize,
                FontFamily = this.FontFamily,
                Style = this.Style,
                Position = this.Position,
                CustomX = this.CustomX,
                CustomY = this.CustomY,
                Rotation = this.Rotation,
                Opacity = this.Opacity,
                UseCustomPosition = this.UseCustomPosition
            };
        }
    }

    /// <summary>
    /// 可序列化的设置类（用于JSON保存）
    /// </summary>
    public class SerializableSettings
    {
        public string Text { get; set; } = "水印";
        public int ColorR { get; set; } = 255; // 红色
        public int ColorG { get; set; } = 0;
        public int ColorB { get; set; } = 0;
        public int ColorA { get; set; } = 255;
        public float FontSize { get; set; } = 8f; // 默认8磅
        public string FontFamily { get; set; } = "宋体"; // 默认宋体
        public WatermarkStyle Style { get; set; } = WatermarkStyle.SemiTransparent; // 默认半透明
        public WatermarkPosition Position { get; set; } = WatermarkPosition.Center;
        public float CustomX { get; set; } = 50f;
        public float CustomY { get; set; } = 50f;
        public float Rotation { get; set; } = 0f; // 默认0度
        public int Opacity { get; set; } = 128;
        public bool UseCustomPosition { get; set; } = false;
        
        // 关键字规则
        public List<KeywordRuleData> KeywordRules { get; set; } = new();
    }

    /// <summary>
    /// 关键字规则数据（可序列化）
    /// </summary>
    public class KeywordRuleData
    {
        public string Keyword { get; set; } = "";
        public string WatermarkText { get; set; } = "";
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
