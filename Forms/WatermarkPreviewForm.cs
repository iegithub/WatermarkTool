using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using WatermarkTool.Models;

namespace WatermarkTool.Forms
{
    /// <summary>
    /// 水印预览窗体 - 可视化预览水印效果，支持鼠标拖拽调整位置
    /// </summary>
    public class WatermarkPreviewForm : Form
    {
        private PictureBox? _previewBox;
        private WatermarkSettings _settings;
        private bool _isDragging;
        private Point _dragStart;
        private float _watermarkX; // 百分比 0-100
        private float _watermarkY; // 百分比 0-100

        public WatermarkSettings Settings => _settings;

        public event Action<WatermarkSettings>? SettingsChanged;

        public WatermarkPreviewForm(WatermarkSettings settings)
        {
            _settings = settings;
            _watermarkX = settings.CustomX;
            _watermarkY = settings.CustomY;

            InitializeForm();
            InitializeControls();
        }

        private void InitializeForm()
        {
            Text = "水印预览 - 拖拽调整位置";
            Size = new Size(700, 500);
            MinimumSize = new Size(500, 400);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(240, 240, 240);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
        }

        private void InitializeControls()
        {
            // 顶部提示
            var tipLabel = new Label
            {
                Text = "💡 拖拽水印文字可调整位置，关闭窗口确认位置",
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("微软雅黑", 9f),
                ForeColor = Color.FromArgb(80, 80, 80),
                BackColor = Color.FromArgb(245, 245, 245)
            };
            Controls.Add(tipLabel);

            // 预览区域
            _previewBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                SizeMode = PictureBoxSizeMode.Normal
            };
            _previewBox.Paint += PreviewBox_Paint;
            _previewBox.MouseDown += PreviewBox_MouseDown;
            _previewBox.MouseMove += PreviewBox_MouseMove;
            _previewBox.MouseUp += PreviewBox_MouseUp;
            _previewBox.Cursor = Cursors.Hand;
            Controls.Add(_previewBox);

            // 底部坐标显示
            var coordLabel = new Label
            {
                Text = "",
                Dock = DockStyle.Bottom,
                Height = 25,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Consolas", 9f),
                ForeColor = Color.FromArgb(100, 100, 100),
                BackColor = Color.FromArgb(245, 245, 245)
            };
            coordLabel.Name = "coordLabel";
            Controls.Add(coordLabel);
        }

        private void PreviewBox_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            var rect = _previewBox.ClientRectangle;

            // 绘制模拟页面背景
            using var bgBrush = new SolidBrush(Color.White);
            g.FillRectangle(bgBrush, rect);

            // 绘制页面边框
            using var borderPen = new Pen(Color.FromArgb(200, 200, 200));
            g.DrawRectangle(borderPen, 1, 1, rect.Width - 2, rect.Height - 2);

            // 绘制模拟文字行（表示文档内容）
            using var textBrush = new SolidBrush(Color.FromArgb(220, 220, 220));
            using var lineFont = new Font("微软雅黑", 10f);
            for (int i = 0; i < 20; i++)
            {
                int y = 40 + i * 22;
                if (y > rect.Height - 20) break;
                int lineWidth = 200 + (i * 37) % 300;
                g.FillRectangle(textBrush, 40, y, lineWidth, 12);
            }

            // 绘制水印
            DrawWatermark(g, rect);

            // 更新坐标显示
            var coordLabel = Controls["coordLabel"] as Label;
            if (coordLabel != null)
            {
                coordLabel.Text = $"X: {_watermarkX:F1}%  Y: {_watermarkY:F1}%  |  旋转: {_settings.Rotation:F0}°  |  大小: {_settings.FontSize:F0}pt";
            }
        }

        private void DrawWatermark(Graphics g, Rectangle rect)
        {
            float x = rect.Width * _watermarkX / 100f;
            float y = rect.Height * _watermarkY / 100f;

            var fontSize = _settings.FontSize;
            var font = new Font(_settings.FontFamily, fontSize, FontStyle.Bold);

            var watermarkColor = Color.FromArgb(_settings.Opacity, _settings.Color.R, _settings.Color.G, _settings.Color.B);
            var text = _settings.Text;

            g.TranslateTransform(x, y);
            g.RotateTransform(_settings.Rotation);

            if (_settings.Style == WatermarkStyle.ArtisticText)
            {
                // 艺术字效果：阴影 + 描边 + 渐变填充
                var textSize = g.MeasureString(text, font);

                // 阴影
                using var shadowBrush = new SolidBrush(Color.FromArgb(_settings.Opacity / 3, 60, 60, 60));
                g.DrawString(text, font, shadowBrush, 3, 3);

                // 描边
                using var path = new GraphicsPath();
                path.AddString(text, font.FontFamily, (int)font.Style, font.Size, new PointF(0, 0), StringFormat.GenericDefault);
                using var outlinePen = new Pen(Color.FromArgb(_settings.Opacity, _settings.Color.R / 2, _settings.Color.G / 2, _settings.Color.B / 2), 2f);
                g.DrawPath(outlinePen, path);

                // 渐变填充
                using var gradBrush = new LinearGradientBrush(
                    new PointF(0, 0),
                    new PointF(textSize.Width, textSize.Height),
                    Color.FromArgb(_settings.Opacity, _settings.Color),
                    Color.FromArgb(_settings.Opacity, Color.FromArgb(
                        Math.Max(0, _settings.Color.R - 40),
                        Math.Max(0, _settings.Color.G - 40),
                        Math.Max(0, _settings.Color.B - 40))));
                g.FillPath(gradBrush, path);
            }
            else
            {
                // 半透明文字
                using var textBrush = new SolidBrush(watermarkColor);
                var textSize = g.MeasureString(text, font);
                g.DrawString(text, font, textBrush, -textSize.Width / 2f, -textSize.Height / 2f);
            }

            g.ResetTransform();
            font.Dispose();
        }

        private void PreviewBox_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _isDragging = true;
                _dragStart = e.Location;
                _previewBox.Cursor = Cursors.SizeAll;
            }
        }

        private void PreviewBox_MouseMove(object? sender, MouseEventArgs e)
        {
            if (_isDragging)
            {
                var dx = e.X - _dragStart.X;
                var dy = e.Y - _dragStart.Y;

                _watermarkX = Math.Max(0, Math.Min(100, _watermarkX + dx * 100f / _previewBox.Width));
                _watermarkY = Math.Max(0, Math.Min(100, _watermarkY + dy * 100f / _previewBox.Height));

                _dragStart = e.Location;
                _settings.CustomX = _watermarkX;
                _settings.CustomY = _watermarkY;
                _settings.UseCustomPosition = true;
                _settings.Position = WatermarkPosition.Custom;

                _previewBox.Invalidate();
                SettingsChanged?.Invoke(_settings);
            }
        }

        private void PreviewBox_MouseUp(object? sender, MouseEventArgs e)
        {
            if (_isDragging)
            {
                _isDragging = false;
                _previewBox.Cursor = Cursors.Hand;
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            _previewBox?.Invalidate();
        }
    }
}
