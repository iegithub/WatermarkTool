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
    /// 主窗体 - 文件选择、关键字管理、水印设置、批量处理
    /// 布局结构：
    /// ┌──────────────────────────────────────────────┐
    /// │  左侧 (SplitContainer.Panel1)  │  右侧 (Panel2) │
    /// │  ┌────────────────────────┐  │  ┌─────────────┐ │
    /// │  │  文件选择 (上)         │  │  │ 水印设置    │ │
    /// │  │                        │  │  │ + 预览区域  │ │
    /// │  ├────────────────────────┤  │  ├─────────────┤ │
    /// │  │  关键字规则 (下)       │  │  │ 操作按钮    │ │
    /// │  └────────────────────────┘  │  └─────────────┘ │
    /// ├──────────────────────────────────────────────┤
    /// │  状态栏                                           │
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

        // ===== 数据 =====
        private readonly List<string> _selectedFiles = new();
        private readonly List<KeywordRule> _keywordRules = new();
        private Color _selectedColor = Color.FromArgb(200, 200, 200);
        private WatermarkSettings _currentSettings;
        private string? _sourceRootPath;

        public MainForm()
        {
            _currentSettings = new WatermarkSettings();
            InitializeForm();
            InitializeUI();
        }

        private void InitializeForm()
        {
            Text = "批量水印工具 - Excel & Word";
            Size = new Size(1100, 750);
            MinimumSize = new Size(950, 650);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(245, 245, 248);
            Font = new Font("微软雅黑", 9f);
        }

        private void InitializeUI()
        {
            // ===== 主布局：左右分栏 =====
            var mainSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 400,
                FixedPanel = FixedPanel.Panel1,
                BackColor = Color.FromArgb(220, 220, 225)
            };

            // ===== 左侧面板：上下分栏（文件选择 / 关键字规则）=====
            var leftSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 340,
                FixedPanel = FixedPanel.Panel1
            };
            leftSplit.Panel1.Controls.Add(CreateFileSelectionPanel());
            leftSplit.Panel2.Controls.Add(CreateKeywordPanel());
            mainSplit.Panel1.Controls.Add(leftSplit);

            // ===== 右侧面板：上下分栏（水印设置 / 操作按钮）=====
            var rightSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 600,
                FixedPanel = FixedPanel.Panel1
            };
            rightSplit.Panel1.Controls.Add(CreateSettingsPanel());
            rightSplit.Panel2.Controls.Add(CreateActionPanel());
            mainSplit.Panel2.Controls.Add(rightSplit);

            Controls.Add(mainSplit);

            // ===== 底部状态栏 =====
            var statusStrip = new StatusStrip();
            toolStripStatusLabel = new ToolStripStatusLabel("就绪");
            toolStripProgressBar = new ToolStripProgressBar { Visible = false, Size = new Size(200, 20) };
            statusStrip.Items.Add(toolStripStatusLabel);
            statusStrip.Items.Add(toolStripProgressBar);
            Controls.Add(statusStrip);
        }

        #region 文件选择面板
        private Panel CreateFileSelectionPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(250, 250, 252),
                Padding = new Padding(10)
            };

            var titleLabel = new Label
            {
                Text = "📄 选择文件",
                Font = new Font("微软雅黑", 11f, FontStyle.Bold),
                Location = new Point(5, 5),
                AutoSize = true
            };

            var btnAddFiles = new Button
            {
                Text = "添加文件",
                Location = new Point(5, 32),
                Size = new Size(100, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 130, 220),
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 9f),
                Cursor = Cursors.Hand
            };
            btnAddFiles.Click += BtnAddFiles_Click;

            var btnAddFolder = new Button
            {
                Text = "添加文件夹",
                Location = new Point(115, 32),
                Size = new Size(100, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 130, 220),
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 9f),
                Cursor = Cursors.Hand
            };
            btnAddFolder.Click += BtnAddFolder_Click;

            var btnClearFiles = new Button
            {
                Text = "清空",
                Location = new Point(225, 32),
                Size = new Size(60, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(180, 80, 80),
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 9f),
                Cursor = Cursors.Hand
            };
            btnClearFiles.Click += BtnClearFiles_Click;

            lstFiles = new ListBox
            {
                Location = new Point(5, 72),
                Size = new Size(375, 200),
                Font = new Font("Consolas", 8.5f),
                SelectionMode = SelectionMode.MultiExtended,
                BorderStyle = BorderStyle.FixedSingle
            };

            var lblCount = new Label
            {
                Text = "已选 0 个文件",
                Location = new Point(5, 278),
                AutoSize = true,
                ForeColor = Color.FromArgb(120, 120, 120)
            };
            lblCount.Name = "lblFileCount";

            panel.Controls.AddRange(new Control[]
            {
                titleLabel, btnAddFiles, btnAddFolder, btnClearFiles,
                lstFiles, lblCount
            });

            return panel;
        }

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
                    var firstFileDir = Path.GetDirectoryName(ofd.FileNames[0]);
                    if (firstFileDir != null)
                        _sourceRootPath = firstFileDir;
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
            UpdateFileCount();
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
            UpdateFileCount();
        }

        private void UpdateFileCount()
        {
            var lbl = Controls.Find("lblFileCount", true).FirstOrDefault() as Label;
            if (lbl != null)
                lbl.Text = $"已选 {_selectedFiles.Count} 个文件";
        }
        #endregion

        #region 关键字管理面板
        private Panel CreateKeywordPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(250, 250, 252),
                Padding = new Padding(10),
                AutoScroll = true
            };

            var titleLabel = new Label
            {
                Text = "🔑 关键字规则",
                Font = new Font("微软雅黑", 11f, FontStyle.Bold),
                Location = new Point(5, 5),
                AutoSize = true
            };

            var lblHint = new Label
            {
                Text = "文件名包含关键字时，使用对应水印文字；否则使用默认水印",
                Location = new Point(5, 27),
                AutoSize = true,
                ForeColor = Color.FromArgb(120, 120, 120),
                Font = new Font("微软雅黑", 8f)
            };

            var lblNewKeyword = new Label { Text = "关键字:", Location = new Point(5, 52), AutoSize = true };
            txtNewKeyword = new TextBox { Location = new Point(60, 49), Width = 100 };
            var lblKwWatermark = new Label { Text = "水印文字:", Location = new Point(170, 52), AutoSize = true };
            txtKeywordWatermark = new TextBox { Location = new Point(235, 49), Width = 100, PlaceholderText = "(留空=用关键字)" };

            var btnAddKeyword = new Button
            {
                Text = "添加",
                Location = new Point(345, 48),
                Size = new Size(50, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(80, 160, 80),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnAddKeyword.Click += BtnAddKeyword_Click;

            lstKeywords = new ListBox
            {
                Location = new Point(5, 85),
                Size = new Size(300, 80),
                Font = new Font("微软雅黑", 9f),
                BorderStyle = BorderStyle.FixedSingle
            };

            var btnRemoveKeyword = new Button
            {
                Text = "删除选中",
                Location = new Point(315, 85),
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(180, 80, 80),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnRemoveKeyword.Click += BtnRemoveKeyword_Click;

            var lblDefault = new Label { Text = "默认水印文字:", Location = new Point(5, 175), AutoSize = true };
            txtDefaultWatermark = new TextBox
            {
                Location = new Point(105, 172),
                Width = 200,
                Text = "水印"
            };

            panel.Controls.AddRange(new Control[]
            {
                titleLabel, lblHint, lblNewKeyword, txtNewKeyword,
                lblKwWatermark, txtKeywordWatermark, btnAddKeyword,
                lstKeywords, btnRemoveKeyword, lblDefault, txtDefaultWatermark
            });

            return panel;
        }

        private void BtnAddKeyword_Click(object? sender, EventArgs e)
        {
            var keyword = txtNewKeyword!.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                MessageBox.Show("请输入关键字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var rule = new KeywordRule
            {
                Keyword = keyword,
                WatermarkText = txtKeywordWatermark!.Text.Trim()
            };

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

        #region 水印设置面板
        private Panel CreateSettingsPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(250, 250, 252),
                Padding = new Padding(10),
                AutoScroll = true
            };

            var titleLabel = new Label
            {
                Text = "🎨 水印设置",
                Font = new Font("微软雅黑", 11f, FontStyle.Bold),
                Location = new Point(5, 5),
                AutoSize = true
            };

            int y = 32;

            // 水印样式
            cmbStyle = new ComboBox
            {
                Location = new Point(100, y - 3),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9f)
            };
            cmbStyle.Items.AddRange(new object[] { "艺术字效果", "半透明文字" });
            cmbStyle.SelectedIndex = 0;
            cmbStyle.SelectedIndexChanged += (s, e) =>
            {
                _currentSettings.Style = cmbStyle.SelectedIndex == 0
                    ? WatermarkStyle.ArtisticText
                    : WatermarkStyle.SemiTransparent;
            };
            panel.Controls.Add(CreateLabel("水印样式:", 5, y));
            panel.Controls.Add(cmbStyle);
            y += 32;

            // 字体
            cmbFontFamily = new ComboBox
            {
                Location = new Point(100, y - 3),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9f)
            };
            cmbFontFamily.Items.AddRange(new object[] { "微软雅黑", "宋体", "黑体", "楷体", "Arial", "Times New Roman" });
            cmbFontFamily.SelectedIndex = 0;
            cmbFontFamily.SelectedIndexChanged += (s, e) =>
            {
                _currentSettings.FontFamily = cmbFontFamily.Text;
            };
            panel.Controls.Add(CreateLabel("字体:", 5, y));
            panel.Controls.Add(cmbFontFamily);
            y += 32;

            // 字体大小
            txtFontSize = new TextBox
            {
                Location = new Point(100, y - 3),
                Width = 80,
                Text = "48"
            };
            txtFontSize.TextChanged += (s, e) =>
            {
                if (float.TryParse(txtFontSize.Text, out var size))
                    _currentSettings.FontSize = size;
            };
            panel.Controls.Add(CreateLabel("字体大小:", 5, y));
            panel.Controls.Add(txtFontSize);
            panel.Controls.Add(CreateLabel("磅", 185, y));
            y += 32;

            // 颜色
            btnColorPicker = new Button
            {
                Location = new Point(100, y - 3),
                Size = new Size(80, 26),
                Text = "选择颜色",
                FlatStyle = FlatStyle.Flat,
                BackColor = _selectedColor,
                Cursor = Cursors.Hand
            };
            btnColorPicker.Click += BtnColorPicker_Click;

            pnlColorPreview = new Panel
            {
                Location = new Point(190, y - 3),
                Size = new Size(50, 26),
                BackColor = _selectedColor,
                BorderStyle = BorderStyle.FixedSingle
            };
            panel.Controls.Add(CreateLabel("水印颜色:", 5, y));
            panel.Controls.Add(btnColorPicker);
            panel.Controls.Add(pnlColorPreview);
            y += 32;

            // 透明度
            nudOpacity = new NumericUpDown
            {
                Location = new Point(100, y - 3),
                Width = 80,
                Minimum = 10,
                Maximum = 255,
                Value = 128
            };
            nudOpacity.ValueChanged += (s, e) =>
            {
                _currentSettings.Opacity = (int)nudOpacity.Value;
            };
            panel.Controls.Add(CreateLabel("透明度:", 5, y));
            panel.Controls.Add(nudOpacity);
            y += 32;

            // 旋转角度
            nudRotation = new NumericUpDown
            {
                Location = new Point(100, y - 3),
                Width = 80,
                Minimum = -360,
                Maximum = 360,
                Value = -30
            };
            nudRotation.ValueChanged += (s, e) =>
            {
                _currentSettings.Rotation = (float)nudRotation.Value;
            };
            panel.Controls.Add(CreateLabel("旋转角度:", 5, y));
            panel.Controls.Add(nudRotation);
            panel.Controls.Add(CreateLabel("度", 185, y));
            y += 32;

            // 位置
            cmbPosition = new ComboBox
            {
                Location = new Point(100, y - 3),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9f)
            };
            cmbPosition.Items.AddRange(new object[]
            {
                "居中", "左上", "上中", "右上",
                "左中", "右中", "左下", "下中", "右下", "自定义"
            });
            cmbPosition.SelectedIndex = 0;
            cmbPosition.SelectedIndexChanged += CmbPosition_SelectedIndexChanged;

            var btnPreview = new Button
            {
                Location = new Point(260, y - 3),
                Size = new Size(80, 26),
                Text = "预览位置",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 130, 220),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnPreview.Click += BtnPreview_Click;
            panel.Controls.Add(CreateLabel("水印位置:", 5, y));
            panel.Controls.Add(cmbPosition);
            panel.Controls.Add(btnPreview);
            y += 38;

            // 预览面板
            var previewPanel = new Panel
            {
                Location = new Point(5, y),
                Size = new Size(350, 180),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };
            previewPanel.Name = "miniPreview";
            previewPanel.Paint += MiniPreview_Paint;

            panel.Controls.Add(previewPanel);

            return panel;
        }

        private Label CreateLabel(string text, int x, int y)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                Font = new Font("微软雅黑", 9f)
            };
        }

        private void BtnColorPicker_Click(object? sender, EventArgs e)
        {
            using var cd = new ColorDialog
            {
                Color = _selectedColor,
                FullOpen = true
            };

            if (cd.ShowDialog() == DialogResult.OK)
            {
                _selectedColor = cd.Color;
                _currentSettings.Color = cd.Color;
                btnColorPicker!.BackColor = cd.Color;
                pnlColorPreview!.BackColor = cd.Color;
                InvalidateMiniPreview();
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

            InvalidateMiniPreview();
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
                InvalidateMiniPreview();
            };
            previewForm.ShowDialog();
        }

        private void MiniPreview_Paint(object? sender, PaintEventArgs e)
        {
            var panel = (Panel)sender!;
            var g = e.Graphics;
            var rect = panel.ClientRectangle;

            g.Clear(Color.White);

            using var lightBrush = new SolidBrush(Color.FromArgb(230, 230, 230));
            for (int i = 0; i < 12; i++)
            {
                int ly = 15 + i * 15;
                if (ly > rect.Height - 10) break;
                int lw = 100 + (i * 31) % 200;
                g.FillRectangle(lightBrush, 20, ly, lw, 8);
            }

            float x = rect.Width * _currentSettings.CustomX / 100f;
            float y = rect.Height * _currentSettings.CustomY / 100f;
            using var font = new Font(_currentSettings.FontFamily, 14f, FontStyle.Bold);
            var color = Color.FromArgb(_currentSettings.Opacity, _currentSettings.Color);

            g.TranslateTransform(x, y);
            g.RotateTransform(_currentSettings.Rotation);

            using var textBrush = new SolidBrush(color);
            var text = txtDefaultWatermark?.Text ?? "水印";
            var size = g.MeasureString(text, font);
            g.DrawString(text, font, textBrush, -size.Width / 2f, -size.Height / 2f);

            g.ResetTransform();

            using var borderPen = new Pen(Color.FromArgb(200, 200, 200));
            g.DrawRectangle(borderPen, 0, 0, rect.Width - 1, rect.Height - 1);
        }

        private void InvalidateMiniPreview()
        {
            foreach (Control c in Controls)
            {
                var p = FindControlRecursive(c, "miniPreview") as Panel;
                p?.Invalidate();
            }
        }

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

        #region 操作面板
        private Panel CreateActionPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 245, 248),
                Padding = new Padding(10)
            };

            var btnStart = new Button
            {
                Text = "🚀 开始批量添加水印",
                Size = new Size(200, 40),
                Location = new Point(10, 10),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(50, 150, 50),
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 11f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnStart.Click += BtnStart_Click;

            var btnBackup = new CheckBox
            {
                Text = "处理前备份原文件（保持目录结构）",
                Location = new Point(230, 18),
                Font = new Font("微软雅黑", 9f),
                Checked = true
            };
            btnBackup.Name = "chkBackup";

            panel.Controls.Add(btnStart);
            panel.Controls.Add(btnBackup);
            return panel;
        }

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
                    var result = new ProcessResult
                    {
                        FilePath = filePath,
                        FileName = fileName
                    };

                    try
                    {
                        var (watermarkText, matchedKeyword) = KeywordMatcher.Match(
                            fileName, _keywordRules, defaultText);

                        result.WatermarkText = watermarkText;
                        result.MatchedKeyword = matchedKeyword;

                        if (backup)
                        {
                            string? backupPath = null;
                            if (!string.IsNullOrEmpty(_sourceRootPath))
                            {
                                backupPath = BackupService.BackupFile(filePath, _sourceRootPath);
                            }
                            else
                            {
                                backupPath = filePath + ".bak";
                                File.Copy(filePath, backupPath, true);
                            }

                            if (backupPath != null)
                                Console.WriteLine($"已备份: {backupPath}");
                        }

                        var ext = Path.GetExtension(filePath).ToLowerInvariant();
                        if (ext == ".docx")
                        {
                            result.Success = WordWatermarkService.AddWatermark(filePath, watermarkText, _currentSettings);
                        }
                        else if (ext == ".xlsx")
                        {
                            result.Success = ExcelWatermarkService.AddWatermark(filePath, watermarkText, _currentSettings);
                        }
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
    }
}
