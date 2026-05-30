using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        private string? _sourceRootPath; // 用户选择的源根目录（用于保持备份目录结构）

        public MainForm()
        {
            _currentSettings = new WatermarkSettings();
            InitializeForm();
            InitializeUI();
        }

        private void InitializeForm()
        {
            Text = "批量水印工具 - Excel & Word";
            Size = new Size(1100, 720);
            MinimumSize = new Size(900, 600);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(245, 245, 248);
            Font = new Font("微软雅黑", 9f);
        }

        private void InitializeUI()
        {
            // 主布局：左右分栏
            var mainPanel = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 380,
                FixedPanel = FixedPanel.Panel1
            };

            // ===== 左侧面板 =====
            var leftPanel = mainPanel.Panel1;
            leftPanel.Controls.Add(CreateFileSelectionPanel());
            leftPanel.Controls.Add(CreateKeywordPanel());

            // ===== 右侧面板 =====
            var rightPanel = mainPanel.Panel2;
            rightPanel.Controls.Add(CreateSettingsPanel());
            rightPanel.Controls.Add(CreateActionPanel());

            Controls.Add(mainPanel);

            // 底部状态栏
            var statusPanel = new StatusStrip();
            toolStripStatusLabel = new ToolStripStatusLabel("就绪");
            toolStripProgressBar = new ToolStripProgressBar { Visible = false, Size = new Size(200, 20) };
            statusPanel.Items.Add(toolStripStatusLabel);
            statusPanel.Items.Add(toolStripProgressBar);
            Controls.Add(statusPanel);
        }

        #region 文件选择面板
        private Panel CreateFileSelectionPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(8),
            };

            var titleLabel = new Label
            {
                Text = "📄 选择文件",
                Font = new Font("微软雅黑", 11f, FontStyle.Bold),
                Location = new Point(0, 0),
                AutoSize = true
            };

            var btnAddFiles = new Button
            {
                Text = "添加文件",
                Location = new Point(0, 30),
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
                Location = new Point(110, 30),
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
                Location = new Point(220, 30),
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
                Location = new Point(0, 70),
                Size = new Size(360, 200),
                Font = new Font("Consolas", 8.5f),
                SelectionMode = SelectionMode.MultiExtended
            };

            var lblCount = new Label
            {
                Text = "已选 0 个文件",
                Location = new Point(0, 275),
                AutoSize = true,
                ForeColor = Color.FromArgb(120, 120, 120)
            };
            lblCount.Name = "lblFileCount";

            panel.Controls.Add(titleLabel);
            panel.Controls.Add(btnAddFiles);
            panel.Controls.Add(btnAddFolder);
            panel.Controls.Add(btnClearFiles);
            panel.Controls.Add(lstFiles);
            panel.Controls.Add(lblCount);

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
                // 记录源根目录（取第一个文件所在目录的父目录，或目录本身）
                if (ofd.FileNames.Length > 0)
                {
                    var firstFileDir = Path.GetDirectoryName(ofd.FileNames[0]);
                    if (firstFileDir != null)
                    {
                        _sourceRootPath = firstFileDir;
                    }
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
                // 记录源根目录
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
                Dock = DockStyle.Bottom,
                Height = 230,
                Padding = new Padding(8)
            };

            var titleLabel = new Label
            {
                Text = "🔑 关键字规则",
                Font = new Font("微软雅黑", 11f, FontStyle.Bold),
                Location = new Point(0, 0),
                AutoSize = true
            };

            var lblHint = new Label
            {
                Text = "文件名包含关键字时，使用对应水印文字；否则使用默认水印",
                Location = new Point(0, 22),
                AutoSize = true,
                ForeColor = Color.FromArgb(120, 120, 120),
                Font = new Font("微软雅黑", 8f)
            };

            // 输入区域
            var lblNewKeyword = new Label { Text = "关键字:", Location = new Point(0, 48), AutoSize = true };
            txtNewKeyword = new TextBox { Location = new Point(55, 45), Width = 100 };

            var lblKwWatermark = new Label { Text = "水印文字:", Location = new Point(165, 48), AutoSize = true };
            txtKeywordWatermark = new TextBox { Location = new Point(230, 45), Width = 100, PlaceholderText = "(留空=用关键字)" };

            var btnAddKeyword = new Button
            {
                Text = "添加",
                Location = new Point(340, 44),
                Size = new Size(50, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(80, 160, 80),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnAddKeyword.Click += BtnAddKeyword_Click;

            // 关键字列表
            lstKeywords = new ListBox
            {
                Location = new Point(0, 80),
                Size = new Size(300, 100),
                Font = new Font("微软雅黑", 9f)
            };

            var btnRemoveKeyword = new Button
            {
                Text = "删除选中",
                Location = new Point(310, 80),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(180, 80, 80),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnRemoveKeyword.Click += BtnRemoveKeyword_Click;

            // 默认水印文字
            var lblDefault = new Label { Text = "默认水印文字:", Location = new Point(0, 190), AutoSize = true };
            txtDefaultWatermark = new TextBox
            {
                Location = new Point(100, 187),
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
                Padding = new Padding(10),
                AutoScroll = true
            };

            var titleLabel = new Label
            {
                Text = "🎨 水印设置",
                Font = new Font("微软雅黑", 11f, FontStyle.Bold),
                Location = new Point(0, 0),
                AutoSize = true
            };

            int y = 30;

            // 水印样式
            var lblStyle = CreateLabel("水印样式:", 0, y);
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
            y += 35;

            // 字体
            var lblFont = CreateLabel("字体:", 0, y);
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
            y += 35;

            // 字体大小
            var lblSize = CreateLabel("字体大小:", 0, y);
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
            var lblSizeUnit = CreateLabel("磅", 185, y);
            y += 35;

            // 颜色
            var lblColor = CreateLabel("水印颜色:", 0, y);
            btnColorPicker = new Button
            {
                Location = new Point(100, y - 3),
                Size = new Size(80, 28),
                Text = "选择颜色",
                FlatStyle = FlatStyle.Flat,
                BackColor = _selectedColor,
                Cursor = Cursors.Hand
            };
            btnColorPicker.Click += BtnColorPicker_Click;

            pnlColorPreview = new Panel
            {
                Location = new Point(190, y - 3),
                Size = new Size(60, 28),
                BackColor = _selectedColor,
                BorderStyle = BorderStyle.FixedSingle
            };
            y += 35;

            // 透明度
            var lblOpacity = CreateLabel("透明度:", 0, y);
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
            y += 35;

            // 旋转角度
            var lblRotation = CreateLabel("旋转角度:", 0, y);
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
            var lblRotUnit = CreateLabel("度", 185, y);
            y += 35;

            // 位置
            var lblPosition = CreateLabel("水印位置:", 0, y);
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
                Size = new Size(80, 28),
                Text = "预览位置",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 130, 220),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnPreview.Click += BtnPreview_Click;
            y += 45;

            // 位置坐标显示
            var lblCoordHint = new Label
            {
                Location = new Point(0, y),
                AutoSize = true,
                ForeColor = Color.FromArgb(120, 120, 120),
                Font = new Font("微软雅黑", 8f),
                Text = "选择「自定义」后可点击「预览位置」拖拽调整"
            };
            y += 25;

            // 预览面板
            var previewPanel = new Panel
            {
                Location = new Point(0, y),
                Size = new Size(350, 200),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };
            previewPanel.Name = "miniPreview";
            previewPanel.Paint += MiniPreview_Paint;

            panel.Controls.AddRange(new Control[]
            {
                titleLabel, lblStyle, cmbStyle, lblFont, cmbFontFamily,
                lblSize, txtFontSize, lblSizeUnit, lblColor, btnColorPicker, pnlColorPreview,
                lblOpacity, nudOpacity, lblRotation, nudRotation, lblRotUnit,
                lblPosition, cmbPosition, btnPreview, lblCoordHint, previewPanel
            });

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
            var posNames = new[] { "Center", "TopLeft", "TopCenter", "TopRight", "MiddleLeft", "MiddleRight", "BottomLeft", "BottomCenter", "BottomRight", "Custom" };
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
            // 切换到自定义模式
            cmbPosition!.SelectedIndex = 9; // Custom
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

            // 模拟文档内容
            using var lightBrush = new SolidBrush(Color.FromArgb(230, 230, 230));
            for (int i = 0; i < 12; i++)
            {
                int ly = 15 + i * 15;
                if (ly > rect.Height - 10) break;
                int lw = 100 + (i * 31) % 200;
                g.FillRectangle(lightBrush, 20, ly, lw, 8);
            }

            // 绘制水印预览
            float x = rect.Width * _currentSettings.CustomX / 100f;
            float y = rect.Height * _currentSettings.CustomY / 100f;
            var font = new Font(_currentSettings.FontFamily, 14f, FontStyle.Bold);
            var color = Color.FromArgb(_currentSettings.Opacity, _currentSettings.Color);

            g.TranslateTransform(x, y);
            g.RotateTransform(_currentSettings.Rotation);

            using var textBrush = new SolidBrush(color);
            var text = txtDefaultWatermark?.Text ?? "水印";
            var size = g.MeasureString(text, font);
            g.DrawString(text, font, textBrush, -size.Width / 2f, -size.Height / 2f);

            g.ResetTransform();
            font.Dispose();

            // 绘制边框
            using var borderPen = new Pen(Color.FromArgb(200, 200, 200));
            g.DrawRectangle(borderPen, 0, 0, rect.Width - 1, rect.Height - 1);
        }

        private void InvalidateMiniPreview()
        {
            // 刷新当前窗体中的miniPreview
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
                Dock = DockStyle.Bottom,
                Height = 60,
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
                Text = "处理前备份原文件",
                Location = new Point(230, 16),
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

            // 收集当前设置
            _currentSettings.Text = defaultText;

            var chkBackup = FindControlRecursive(this, "chkBackup") as CheckBox;
            bool backup = chkBackup?.Checked ?? true;

            // 禁用按钮
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
                        // 匹配关键字
                        var (watermarkText, matchedKeyword) = KeywordMatcher.Match(
                            fileName, _keywordRules, defaultText);

                        result.WatermarkText = watermarkText;
                        result.MatchedKeyword = matchedKeyword;

                        // 备份 - 保持目录结构
                        if (backup)
                        {
                            string? backupPath = null;
                            if (!string.IsNullOrEmpty(_sourceRootPath))
                            {
                                backupPath = BackupService.BackupFile(filePath, _sourceRootPath);
                            }
                            else
                            {
                                // 如果没有源根目录（理论上不会发生），使用简单备份
                                backupPath = filePath + ".bak";
                                File.Copy(filePath, backupPath, true);
                            }
                            
                            if (backupPath != null)
                            {
                                Console.WriteLine($"已备份: {backupPath}");
                            }
                        }

                        // 根据文件类型添加水印
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
                        {
                            result.ErrorMessage = "未知错误";
                        }
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

            // 恢复按钮
            btn.Enabled = true;
            toolStripProgressBar.Visible = false;

            // 显示结果
            ShowResults(results);
        }

        private void ShowResults(List<ProcessResult> results)
        {
            int success = results.Count(r => r.Success);
            int failed = results.Count(r => !r.Success);

            toolStripStatusLabel!.Text = $"处理完成: 成功 {success}, 失败 {failed}, 共 {results.Count} 个文件";

            // 创建结果窗体
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

            // 设置列宽
            dgvResults.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            resultForm.Controls.Add(dgvResults);
            resultForm.ShowDialog();
        }
        #endregion
    }
}
