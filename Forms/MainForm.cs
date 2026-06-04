using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using WatermarkTool.Models;
using WatermarkTool.Services;

namespace WatermarkTool.Forms
{
    /// <summary>
    /// 主窗体 - 使用 TableLayoutPanel 实现自适应布局
    /// 布局结构：
    /// ┌──────────────────────────────────────────────┐
    /// │  菜单栏 (MenuStrip)                            │
    /// ├──────────────────────┬───────────────────────┤
    /// │  文件列表 (ListBox)  │  水印设置 (FlowLayoutPanel)│
    /// │                      │  + 预览区域              │
    /// │                      ├───────────────────────┤
    /// │  关键字规则          │  操作按钮              │
    /// │  + 默认水印文字      │  + 备份选项            │
    /// ├──────────────────────┴───────────────────────┤
    /// │  状态栏 (StatusStrip)                          │
    /// └──────────────────────────────────────────────┘
    /// </summary>
    public class MainForm : Form
    {
        // ===== 控件声明 =====
        private TextBox? txtDefaultWatermark;
        private TextBox? txtFontSize;
        private ComboBox? cmbPosition;
        private ComboBox? cmbStyle;
        private ComboBox? cmbFontFamily;
        private NumericUpDown? nudOpacity;
        private NumericUpDown? nudRotation;
        private Button? btnColorPicker;
        private Panel? pnlColorPreview;
        private ListBox? lstFiles;
        private ListBox? lstKeywords;
        private TextBox? txtKeywordWatermark;
        private TextBox? txtNewKeyword;
        private DataGridView? dgvResults;
        private ToolStripStatusLabel? toolStripStatusLabel;
        private ToolStripProgressBar? toolStripProgressBar;
        private Panel? miniPreviewPanel;

        // ===== 常量 =====
        private const string VERSION = "1.6.0";

        // ===== 数据 =====
        private readonly List<string> _selectedFiles = new();
        private readonly List<KeywordRule> _keywordRules = new();
        private Color _selectedColor = Color.Red;
        private WatermarkSettings _currentSettings;
        private string? _sourceRootPath;

        public MainForm()
        {
            _currentSettings = new WatermarkSettings();
            InitializeForm();
            InitializeUI();
            LoadSavedSettings();
        }

        private void InitializeForm()
        {
            Text = $"批量水印工具 v{VERSION} - Excel & Word";
            Size = new Size(1100, 720);
            MinimumSize = new Size(900, 600);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(245, 245, 248);
            Font = new Font("微软雅黑", 9f);
            this.FormClosing += (s, e) => SaveCurrentSettings();
        }

        private void InitializeUI()
        {
            // ===== 菜单栏 =====
            var menuStrip = new MenuStrip();
            var fileMenu = new ToolStripMenuItem("文件(&F)");
            fileMenu.DropDownItems.AddRange(new ToolStripItem[] {
                new ToolStripMenuItem("保存设置", null, (s, e) => SaveCurrentSettings()),
                new ToolStripMenuItem("重新加载设置", null, (s, e) => LoadSavedSettings()),
                new ToolStripSeparator(),
                new ToolStripMenuItem("退出", null, (s, e) => this.Close())
            });
            var helpMenu = new ToolStripMenuItem("帮助(&H)");
            helpMenu.DropDownItems.AddRange(new ToolStripItem[] {
                new ToolStripMenuItem("使用帮助", null, ShowHelpDialog),
                new ToolStripSeparator(),
                new ToolStripMenuItem("关于", null, ShowAboutDialog)
            });
            menuStrip.Items.AddRange(new ToolStripItem[] { fileMenu, helpMenu });
            Controls.Add(menuStrip);

            // ===== 主容器：使用 TableLayoutPanel =====
            var mainTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.FromArgb(230, 230, 235)
            };
            mainTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45F)); // 左侧 45%
            mainTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55F)); // 右侧 55%

            // ===== 左侧面板 =====
            var leftPanel = CreateLeftPanel();
            mainTable.Controls.Add(leftPanel, 0, 0);

            // ===== 右侧面板 =====
            var rightPanel = CreateRightPanel();
            mainTable.Controls.Add(rightPanel, 1, 0);

            Controls.Add(mainTable);

            // ===== 底部状态栏 =====
            var statusStrip = new StatusStrip();
            toolStripStatusLabel = new ToolStripStatusLabel("就绪");
            toolStripProgressBar = new ToolStripProgressBar { Visible = false, Size = new Size(200, 20) };
            statusStrip.Items.Add(toolStripStatusLabel);
            statusStrip.Items.Add(toolStripProgressBar);
            Controls.Add(statusStrip);
        }

        #region 左侧面板
        private Panel CreateLeftPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 245, 248),
                Padding = new Padding(5)
            };

            // 使用 FlowLayoutPanel 让控件自动排列
            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(5)
            };

            // ===== 文件选择区域 =====
            flow.Controls.Add(CreateSectionLabel("📄 选择文件"));

            var btnRow1 = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 440,
                Height = 36,
                Margin = new Padding(0, 2, 0, 5)
            };
            btnRow1.Controls.Add(CreateButton("添加文件", 100, 30, Color.FromArgb(70, 130, 220), BtnAddFiles_Click));
            btnRow1.Controls.Add(CreateButton("添加文件夹", 100, 30, Color.FromArgb(70, 130, 220), BtnAddFolder_Click));
            btnRow1.Controls.Add(CreateButton("清空", 60, 30, Color.FromArgb(180, 80, 80), BtnClearFiles_Click));
            flow.Controls.Add(btnRow1);

            lstFiles = new ListBox
            {
                Width = 440,
                Height = 180,
                Font = new Font("Consolas", 8.5f),
                SelectionMode = SelectionMode.MultiExtended,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 0, 0, 2)
            };
            flow.Controls.Add(lstFiles);

            // ===== 关键字规则区域 =====
            flow.Controls.Add(CreateSectionLabel("🔑 关键字规则（可选）"));

            var hintLabel = new Label
            {
                Text = "文件名含关键字时使用对应水印，否则用默认水印",
                AutoSize = true,
                ForeColor = Color.FromArgb(120, 120, 120),
                Font = new Font("微软雅黑", 8f),
                Margin = new Padding(0, 0, 0, 5)
            };
            flow.Controls.Add(hintLabel);

            var keywordRow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 440,
                Height = 30,
                Margin = new Padding(0, 0, 0, 5)
            };
            keywordRow.Controls.Add(new Label { Text = "关键字:", AutoSize = true, Margin = new Padding(0, 4, 3, 0) });
            txtNewKeyword = new TextBox { Width = 100, Margin = new Padding(0, 0, 5, 0) };
            keywordRow.Controls.Add(txtNewKeyword);
            keywordRow.Controls.Add(new Label { Text = "水印:", AutoSize = true, Margin = new Padding(0, 4, 3, 0) });
            txtKeywordWatermark = new TextBox { Width = 100, Margin = new Padding(0, 0, 5, 0) };
            keywordRow.Controls.Add(txtKeywordWatermark);
            keywordRow.Controls.Add(CreateButton("添加", 50, 26, Color.FromArgb(80, 160, 80), BtnAddKeyword_Click));
            flow.Controls.Add(keywordRow);

            lstKeywords = new ListBox
            {
                Width = 360,
                Height = 60,
                Font = new Font("微软雅黑", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 0, 5, 0)
            };
            flow.Controls.Add(lstKeywords);

            var btnRemoveKw = CreateButton("删除选中", 75, 26, Color.FromArgb(180, 80, 80), BtnRemoveKeyword_Click);
            flow.Controls.Add(btnRemoveKw);

            // ===== 默认水印文字 =====
            flow.Controls.Add(CreateSectionLabel("📝 默认水印文字"));

            var defaultRow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 440,
                Height = 30,
                Margin = new Padding(0, 0, 0, 5)
            };
            defaultRow.Controls.Add(new Label { Text = "水印文字:", AutoSize = true, Margin = new Padding(0, 4, 3, 0) });
            txtDefaultWatermark = new TextBox { Width = 200, Text = "水印" };
            defaultRow.Controls.Add(txtDefaultWatermark);
            flow.Controls.Add(defaultRow);

            panel.Controls.Add(flow);
            return panel;
        }
        #endregion

        #region 右侧面板
        private Panel CreateRightPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 245, 248),
                Padding = new Padding(5)
            };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(5)
            };

            // ===== 水印设置区域 =====
            flow.Controls.Add(CreateSectionLabel("🎨 水印设置"));

            // 使用紧凑的行布局
            flow.Controls.Add(CreateSettingRow("水印样式:", CreateStyleComboBox()));
            flow.Controls.Add(CreateSettingRow("字体:", CreateFontComboBox()));
            flow.Controls.Add(CreateFontSizeRow());
            flow.Controls.Add(CreateColorRow());
            flow.Controls.Add(CreateSettingRow("透明度:", CreateOpacityRow()));
            flow.Controls.Add(CreateRotationRow());
            flow.Controls.Add(CreatePositionRow());

            // 预览面板
            flow.Controls.Add(new Label
            {
                Text = "预览:",
                AutoSize = true,
                Margin = new Padding(0, 5, 0, 2),
                Font = new Font("微软雅黑", 9f)
            }));

            miniPreviewPanel = new Panel
            {
                Width = 500,
                Height = 120,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Margin = new Padding(0, 0, 0, 10)
            };
            miniPreviewPanel.Paint += MiniPreview_Paint;
            flow.Controls.Add(miniPreviewPanel);

            // ===== 操作区域 =====
            flow.Controls.Add(CreateSectionLabel("🚀 操作"));

            var actionRow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 500,
                Height = 40,
                Margin = new Padding(0, 0, 0, 5)
            };

            var btnStart = new Button
            {
                Text = "开始批量添加水印",
                Size = new Size(180, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(50, 150, 50),
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 10, 0)
            };
            btnStart.Click += BtnStart_Click;
            actionRow.Controls.Add(btnStart);

            var chkBackup = new CheckBox
            {
                Text = "处理前备份",
                Font = new Font("微软雅黑", 9f),
                Checked = true,
                Margin = new Padding(0, 8, 0, 0)
            };
            chkBackup.Name = "chkBackup";
            actionRow.Controls.Add(chkBackup);

            flow.Controls.Add(actionRow);

            panel.Controls.Add(flow);
            return panel;
        }

        // ===== 设置行创建辅助方法 =====

        private FlowLayoutPanel CreateSettingRow(string labelText, Control inputControl)
        {
            var row = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 500,
                Height = 28,
                Margin = new Padding(0, 1, 0, 1)
            };
            var label = new Label
            {
                Text = labelText,
                Width = 75,
                TextAlign = ContentAlignment.MiddleRight,
                Margin = new Padding(0, 3, 5, 0)
            };
            row.Controls.Add(label);
            row.Controls.Add(inputControl);
            return row;
        }

        private ComboBox CreateStyleComboBox()
        {
            cmbStyle = new ComboBox
            {
                Width = 140,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9f)
            };
            cmbStyle.Items.AddRange(new object[] { "艺术字效果", "半透明文字" });
            cmbStyle.SelectedIndex = 1;
            cmbStyle.SelectedIndexChanged += (s, e) =>
            {
                _currentSettings.Style = cmbStyle.SelectedIndex == 0
                    ? WatermarkStyle.ArtisticText : WatermarkStyle.SemiTransparent;
            };
            return cmbStyle;
        }

        private ComboBox CreateFontComboBox()
        {
            cmbFontFamily = new ComboBox
            {
                Width = 140,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9f)
            };
            cmbFontFamily.Items.AddRange(new object[] { "宋体", "微软雅黑", "黑体", "楷体", "Arial", "Times New Roman" });
            cmbFontFamily.SelectedIndex = 0;
            cmbFontFamily.SelectedIndexChanged += (s, e) =>
            {
                _currentSettings.FontFamily = cmbFontFamily.Text;
            };
            return cmbFontFamily;
        }

        private FlowLayoutPanel CreateFontSizeRow()
        {
            var row = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 500,
                Height = 28,
                Margin = new Padding(0, 1, 0, 1)
            };
            row.Controls.Add(new Label { Text = "字体大小:", Width = 75, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(0, 3, 5, 0) });
            txtFontSize = new TextBox { Width = 60, Text = "8" };
            txtFontSize.TextChanged += (s, e) =>
            {
                if (float.TryParse(txtFontSize.Text, out var size))
                    _currentSettings.FontSize = size;
            };
            row.Controls.Add(txtFontSize);
            row.Controls.Add(new Label { Text = "磅", AutoSize = true, Margin = new Padding(5, 3, 0, 0) });
            return row;
        }

        private FlowLayoutPanel CreateColorRow()
        {
            var row = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 500,
                Height = 28,
                Margin = new Padding(0, 1, 0, 1)
            };
            row.Controls.Add(new Label { Text = "水印颜色:", Width = 75, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(0, 3, 5, 0) });
            btnColorPicker = new Button
            {
                Text = "选择颜色",
                Size = new Size(70, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = _selectedColor,
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 5, 0)
            };
            btnColorPicker.Click += BtnColorPicker_Click;
            row.Controls.Add(btnColorPicker);
            pnlColorPreview = new Panel
            {
                Size = new Size(40, 24),
                BackColor = _selectedColor,
                BorderStyle = BorderStyle.FixedSingle
            };
            row.Controls.Add(pnlColorPreview);
            return row;
        }

        private FlowLayoutPanel CreateOpacityRow()
        {
            var row = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 500,
                Height = 28,
                Margin = new Padding(0, 1, 0, 1)
            };
            row.Controls.Add(new Label { Text = "透明度:", Width = 75, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(0, 3, 5, 0) });
            nudOpacity = new NumericUpDown
            {
                Width = 70,
                Minimum = 10,
                Maximum = 255,
                Value = 128
            };
            nudOpacity.ValueChanged += (s, e) =>
            {
                _currentSettings.Opacity = (int)nudOpacity.Value;
            };
            row.Controls.Add(nudOpacity);
            return row;
        }

        private FlowLayoutPanel CreateRotationRow()
        {
            var row = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 500,
                Height = 28,
                Margin = new Padding(0, 1, 0, 1)
            };
            row.Controls.Add(new Label { Text = "旋转角度:", Width = 75, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(0, 3, 5, 0) });
            nudRotation = new NumericUpDown
            {
                Width = 70,
                Minimum = -360,
                Maximum = 360,
                Value = 0
            };
            nudRotation.ValueChanged += (s, e) =>
            {
                _currentSettings.Rotation = (float)nudRotation.Value;
            };
            row.Controls.Add(nudRotation);
            row.Controls.Add(new Label { Text = "度", AutoSize = true, Margin = new Padding(5, 3, 0, 0) });
            return row;
        }

        private FlowLayoutPanel CreatePositionRow()
        {
            var row = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Width = 500,
                Height = 28,
                Margin = new Padding(0, 1, 0, 1)
            };
            row.Controls.Add(new Label { Text = "水印位置:", Width = 75, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(0, 3, 5, 0) });
            cmbPosition = new ComboBox
            {
                Width = 100,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9f)
            };
            cmbPosition.Items.AddRange(new object[]
            {
                "居中", "左上", "上中", "右上", "左中", "右中", "左下", "下中", "右下", "自定义"
            });
            cmbPosition.SelectedIndex = 0;
            cmbPosition.SelectedIndexChanged += CmbPosition_SelectedIndexChanged;
            row.Controls.Add(cmbPosition);

            var btnPreview = new Button
            {
                Text = "预览位置",
                Size = new Size(75, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 130, 220),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Margin = new Padding(10, 0, 0, 0)
            };
            btnPreview.Click += BtnPreview_Click;
            row.Controls.Add(btnPreview);
            return row;
        }

        // ===== 通用辅助方法 =====

        private Label CreateSectionLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = new Font("微软雅黑", 10f, FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 8, 0, 3)
            };
        }

        private Button CreateButton(string text, int width, int height, Color bgColor, EventHandler onClick)
        {
            return new Button
            {
                Text = text,
                Size = new Size(width, height),
                FlatStyle = FlatStyle.Flat,
                BackColor = bgColor,
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 9f),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 5, 0)
            };
        }
        #endregion

        #region 文件操作
        private void BtnAddFiles_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Title = "选择Excel或Word文件",
                Filter = "Office文件|*.xlsx;*.docx|Excel文件|*.xlsx|Word文件|*.docx",
                Multiselect = true
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if (ofd.FileNames.Length > 0)
                {
                    var dir = Path.GetDirectoryName(ofd.FileNames[0]);
                    if (dir != null) _sourceRootPath = dir;
                }
                AddFiles(ofd.FileNames);
            }
        }

        private void BtnAddFolder_Click(object? sender, EventArgs e)
        {
            using var fbd = new FolderBrowserDialog
            {
                Description = "选择包含Excel/Word文件的文件夹"
            };
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                _sourceRootPath = fbd.SelectedPath;
                var files = Directory.GetFiles(fbd.SelectedPath, "*.*", SearchOption.AllDirectories)
                    .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    .ToArray();
                AddFiles(files);
            }
        }

        private void BtnClearFiles_Click(object? sender, EventArgs e)
        {
            _selectedFiles.Clear();
            lstFiles!.Items.Clear();
        }

        private void AddFiles(string[] files)
        {
            foreach (var file in files)
            {
                if (!_selectedFiles.Contains(file))
                {
                    _selectedFiles.Add(file);
                    lstFiles!.Items.Add(Path.GetFileName(file));
                }
            }
        }
        #endregion

        #region 关键字操作
        private void BtnAddKeyword_Click(object? sender, EventArgs e)
        {
            var keyword = txtNewKeyword!.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                MessageBox.Show("请输入关键字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var rule = new KeywordRule { Keyword = keyword, WatermarkText = txtKeywordWatermark!.Text.Trim() };
            _keywordRules.Add(rule);
            lstKeywords!.Items.Add($"{keyword} → {rule.GetWatermarkText()}");
            txtNewKeyword.Clear();
            txtKeywordWatermark.Clear();
        }

        private void BtnRemoveKeyword_Click(object? sender, EventArgs e)
        {
            if (lstKeywords!.SelectedIndex >= 0)
            {
                _keywordRules.RemoveAt(lstKeywords.SelectedIndex);
                lstKeywords.Items.RemoveAt(lstKeywords.SelectedIndex);
            }
        }
        #endregion

        #region 水印设置事件
        private void BtnColorPicker_Click(object? sender, EventArgs e)
        {
            using var cd = new ColorDialog { Color = _selectedColor, FullOpen = true };
            if (cd.ShowDialog() == DialogResult.OK)
            {
                _selectedColor = cd.Color;
                _currentSettings.Color = cd.Color;
                btnColorPicker!.BackColor = cd.Color;
                pnlColorPreview!.BackColor = cd.Color;
                miniPreviewPanel?.Invalidate();
            }
        }

        private void CmbPosition_SelectedIndexChanged(object? sender, EventArgs e)
        {
            _currentSettings.Position = (WatermarkPosition)cmbPosition!.SelectedIndex;
            if (cmbPosition.SelectedIndex < 9)
            {
                var positions = new (float x, float y)[]
                {
                    (50, 50), (15, 15), (50, 15), (85, 15),
                    (15, 50), (85, 50), (15, 85), (50, 85), (85, 85)
                };
                _currentSettings.CustomX = positions[cmbPosition.SelectedIndex].x;
                _currentSettings.CustomY = positions[cmbPosition.SelectedIndex].y;
            }
            miniPreviewPanel?.Invalidate();
        }

        private void BtnPreview_Click(object? sender, EventArgs e)
        {
            cmbPosition!.SelectedIndex = 9;
            _currentSettings.Position = WatermarkPosition.Custom;
            _currentSettings.UseCustomPosition = true;
            using var previewForm = new WatermarkPreviewForm(_currentSettings);
            previewForm.SettingsChanged += (s) =>
            {
                _currentSettings = s;
                miniPreviewPanel?.Invalidate();
            };
            previewForm.ShowDialog();
        }

        private void MiniPreview_Paint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            var rect = miniPreviewPanel!.ClientRectangle;
            g.Clear(Color.White);

            using var lightBrush = new SolidBrush(Color.FromArgb(230, 230, 230));
            for (int i = 0; i < 8; i++)
            {
                int ly = 10 + i * 14;
                if (ly > rect.Height - 10) break;
                int lw = 60 + (i * 31) % 150;
                g.FillRectangle(lightBrush, 15, ly, lw, 7);
            }

            float x = rect.Width * _currentSettings.CustomX / 100f;
            float y = rect.Height * _currentSettings.CustomY / 100f;
            using var font = new Font(_currentSettings.FontFamily, 12f, FontStyle.Bold);
            var color = Color.FromArgb(_currentSettings.Opacity, _currentSettings.Color);
            g.TranslateTransform(x, y);
            g.RotateTransform(_currentSettings.Rotation);
            using var textBrush = new SolidBrush(color);
            var text = txtDefaultWatermark?.Text ?? "水印";
            var size = g.MeasureString(text, font);
            g.DrawString(text, font, textBrush, -size.Width / 2f, -size.Height / 2f);
            g.ResetTransform();
        }
        #endregion

        #region 批量处理
        private async void BtnStart_Click(object? sender, EventArgs e)
        {
            if (_selectedFiles.Count == 0)
            {
                MessageBox.Show("请先添加文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var defaultText = txtDefaultWatermark!.Text.Trim();
            if (string.IsNullOrEmpty(defaultText))
            {
                MessageBox.Show("请设置默认水印文字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _currentSettings.Text = defaultText;

            var chkBackup = FindControlRecursive(this, "chkBackup") as CheckBox;
            bool backup = chkBackup?.Checked ?? true;

            var btn = (Button)sender!;
            btn.Enabled = false;
            toolStripProgressBar!.Visible = true;
            toolStripProgressBar.Value = 0;
            toolStripProgressBar.Maximum = _selectedFiles.Count;

            var results = new List<ProcessResult>();

            await Task.Run(() =>
            {
                for (int i = 0; i < _selectedFiles.Count; i++)
                {
                    var filePath = _selectedFiles[i];
                    var fileName = Path.GetFileName(filePath);
                    var result = new ProcessResult { FilePath = filePath, FileName = fileName };

                    try
                    {
                        var (watermarkText, matchedKeyword) = KeywordMatcher.Match(fileName, _keywordRules, defaultText);
                        result.WatermarkText = watermarkText;
                        result.MatchedKeyword = matchedKeyword;

                        if (backup)
                        {
                            if (!string.IsNullOrEmpty(_sourceRootPath))
                                BackupService.BackupFile(filePath, _sourceRootPath);
                            else
                                File.Copy(filePath, filePath + ".bak", true);
                        }

                        var ext = Path.GetExtension(filePath).ToLowerInvariant();
                        if (ext == ".docx")
                            result.Success = WordWatermarkService.AddWatermark(filePath, watermarkText, _currentSettings);
                        else if (ext == ".xlsx")
                            result.Success = ExcelWatermarkService.AddWatermark(filePath, watermarkText, _currentSettings);
                        else
                        {
                            result.Success = false;
                            result.ErrorMessage = "不支持的文件格式";
                        }

                        if (!result.Success && string.IsNullOrEmpty(result.ErrorMessage))
                            result.ErrorMessage = "未知错误";
                    }
                    catch (Exception ex)
                    {
                        result.Success = false;
                        result.ErrorMessage = ex.Message;
                    }

                    results.Add(result);

                    this.Invoke(new Action(() =>
                    {
                        toolStripProgressBar!.Value = i + 1;
                        toolStripStatusLabel!.Text = $"正在处理: {fileName} ({i + 1}/{_selectedFiles.Count})";
                    }));
                }
            });

            btn.Enabled = true;
            toolStripProgressBar.Visible = false;
            ShowResults(results);
        }

        private void ShowResults(List<ProcessResult> results)
        {
            int success = results.Count(r => r.Success);
            int failed = results.Count(r => !r.Success);
            toolStripStatusLabel!.Text = $"处理完成: 成功 {success}, 失败 {failed}, 共 {results.Count} 个文件";

            var resultForm = new Form
            {
                Text = "处理结果",
                Size = new Size(800, 500),
                StartPosition = FormStartPosition.CenterParent,
                Font = new Font("微软雅黑", 9f)
            };

            dgvResults = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AutoGenerateColumns = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AllowUserToAddRows = false
            };

            dgvResults.DataSource = results.Select(r => new
            {
                文件名 = r.FileName,
                水印文字 = r.WatermarkText,
                匹配关键字 = r.MatchedKeyword,
                状态 = r.Success ? "✅ 成功" : "❌ 失败",
                错误信息 = r.ErrorMessage
            }).ToList();

            dgvResults.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            resultForm.Controls.Add(dgvResults);
            resultForm.ShowDialog();
        }
        #endregion

        #region 设置保存与加载
        private void SaveCurrentSettings()
        {
            try
            {
                _currentSettings.Text = txtDefaultWatermark?.Text ?? "水印";
                _currentSettings.Color = _selectedColor;
                _currentSettings.FontSize = float.TryParse(txtFontSize?.Text, out var size) ? size : 8f;
                _currentSettings.FontFamily = cmbFontFamily?.Text ?? "宋体";
                _currentSettings.Style = cmbStyle?.SelectedIndex == 0 ? WatermarkStyle.ArtisticText : WatermarkStyle.SemiTransparent;
                _currentSettings.Position = (WatermarkPosition)(cmbPosition?.SelectedIndex ?? 0);
                _currentSettings.Opacity = (int)(nudOpacity?.Value ?? 128);
                _currentSettings.Rotation = (float)(nudRotation?.Value ?? 0);
                SettingsService.SaveSettings(_currentSettings, _keywordRules);
            }
            catch { }
        }

        private void LoadSavedSettings()
        {
            try
            {
                var (settings, keywordRules) = SettingsService.LoadSettings();
                if (settings != null)
                {
                    _currentSettings = settings;
                    _selectedColor = settings.Color;
                    _keywordRules.Clear();
                    _keywordRules.AddRange(keywordRules);
                    lstKeywords?.Items.Clear();
                    foreach (var rule in _keywordRules)
                        lstKeywords?.Items.Add($"{rule.Keyword} → {rule.GetWatermarkText()}");
                    ApplySettingsToUI();
                }
            }
            catch { }
        }

        private void ApplySettingsToUI()
        {
            if (txtDefaultWatermark != null) txtDefaultWatermark.Text = _currentSettings.Text;
            if (txtFontSize != null) txtFontSize.Text = _currentSettings.FontSize.ToString();
            if (cmbFontFamily != null) cmbFontFamily.SelectedIndex = Math.Max(0, cmbFontFamily.Items.IndexOf(_currentSettings.FontFamily));
            if (cmbStyle != null) cmbStyle.SelectedIndex = (int)_currentSettings.Style;
            if (cmbPosition != null) cmbPosition.SelectedIndex = (int)_currentSettings.Position;
            if (nudOpacity != null) nudOpacity.Value = _currentSettings.Opacity;
            if (nudRotation != null) nudRotation.Value = (decimal)_currentSettings.Rotation;
            if (btnColorPicker != null) btnColorPicker.BackColor = _selectedColor;
            if (pnlColorPreview != null) pnlColorPreview.BackColor = _selectedColor;
            miniPreviewPanel?.Invalidate();
        }
        #endregion

        #region 帮助与关于
        private void ShowAboutDialog(object? sender, EventArgs e)
        {
            MessageBox.Show($@"批量水印工具 v{VERSION}

功能：批量给 Excel/Word 添加水印
技术：.NET 6.0 + Open XML SDK
许可：MIT License", "关于", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowHelpDialog(object? sender, EventArgs e)
        {
            MessageBox.Show(@"【使用帮助】

1. 添加文件：支持 .xlsx 和 .docx
2. 关键字规则：文件名含关键字时自动匹配水印
3. 水印设置：样式、字体、大小、颜色、透明度、旋转、位置
4. 预览位置：点击「预览位置」可拖拽定位
5. 开始处理：勾选备份可保留原文件

【注意】
• Word 水印在页面正文区域（非页眉）
• Excel 水印不触发公式重算
• 备份保持原有目录结构", "使用帮助", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion

        #region 工具方法
        private Control? FindControlRecursive(Control parent, string name)
        {
            if (parent.Name == name) return parent;
            foreach (Control c in parent.Controls)
            {
                var found = FindControlRecursive(c, name);
                if (found != null) return found;
            }
            return null;
        }
        #endregion
    }
}
