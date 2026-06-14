using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
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
    /// Word文档水印添加服务
    /// 使用 OpenXML SDK 原生 API 创建锚定到页面的水印图片
    /// </summary>
    public static class WordWatermarkService
    {
        /// <summary>
        /// 给Word文档添加水印（仅第一页）
        /// </summary>
        public static bool AddWatermark(string filePath, string watermarkText, WatermarkSettings settings)
        {
            try
            {
                using var watermarkImage = GenerateWatermarkImage(watermarkText, settings);

                using var doc = WordprocessingDocument.Open(filePath, true);

                var mainPart = doc.MainDocumentPart;
                if (mainPart == null)
                    return false;

                var body = mainPart.Document.Body;
                if (body == null)
                    return false;

                // 移除已有的水印
                RemoveExistingWatermarks(mainPart, body);

                // 添加图片到 MainDocumentPart
                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var ms = new MemoryStream())
                {
                    watermarkImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Position = 0;
                    imagePart.FeedData(ms);
                }
                string imageRelId = mainPart.GetIdOfPart(imagePart);

                // 计算尺寸和位置
                long imageWidthEmu = watermarkImage.Width * 9525;
                long imageHeightEmu = watermarkImage.Height * 9525;

                // A4 页面尺寸
                long pageWidthEmu = 595440;
                long pageHeightEmu = 841920;

                // 计算位置
                long xEmu = (long)(pageWidthEmu * settings.CustomX / 100f) - imageWidthEmu / 2;
                long yEmu = (long)(pageHeightEmu * settings.CustomY / 100f) - imageHeightEmu / 2;

                // 确保位置有效
                xEmu = Math.Max(0, Math.Min(xEmu, pageWidthEmu - imageWidthEmu));
                yEmu = Math.Max(0, Math.Min(yEmu, pageHeightEmu - imageHeightEmu));

                int rotationEmu = (int)(settings.Rotation * 60000);

                // 创建水印段落
                var watermarkPara = CreateWatermarkParagraph(imageRelId, imageWidthEmu, imageHeightEmu, xEmu, yEmu, rotationEmu);
                body.PrependChild(watermarkPara);

                // 确保 SectionProperties 存在
                if (!body.Elements<SectionProperties>().Any())
                {
                    body.AppendChild(new SectionProperties());
                }

                mainPart.Document.Save();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Word水印添加失败: {ex.Message}");
                Console.WriteLine($"堆栈: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 创建水印段落 - 使用原生 OpenXML SDK API
        /// </summary>
        private static Paragraph CreateWatermarkParagraph(
            string imageRelId, long widthEmu, long heightEmu,
            long xEmu, long yEmu, int rotationEmu)
        {
            // 创建段落
            var paragraph = new Paragraph();
            var run = new Run();

            // 创建 Drawing
            var drawing = new Drawing();

            // 创建 Anchor
            var anchor = new DW.Anchor(
                new DW.SimplePosition { X = 0, Y = 0 },
                new DW.HorizontalPosition(
                    new DW.PositionOffset(xEmu.ToString())
                )
                {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page
                },
                new DW.VerticalPosition(
                    new DW.PositionOffset(yEmu.ToString())
                )
                {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page
                },
                new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent
                {
                    LeftEdge = 0,
                    TopEdge = 0,
                    RightEdge = 0,
                    BottomEdge = 0
                },
                new DW.WrapNone(),
                new DW.DocProperties
                {
                    Id = 1,
                    Name = "Watermark"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new A.NonVisualDrawingProperties
                                {
                                    Id = 0,
                                    Name = "Watermark.png"
                                },
                                new A.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip { Embed = imageRelId },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }
                                )
                                {
                                    Rotation = rotationEmu
                                },
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle
                                }
                            )
                        )
                    )
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }
                )
            )
            {
                DistanceFromTop = 0,
                DistanceFromBottom = 0,
                DistanceFromLeft = 0,
                DistanceFromRight = 0,
                SimplePosition = false,
                RelativeHeight = -251658240,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };

            drawing.Append(anchor);
            run.Append(drawing);
            paragraph.Append(run);

            return paragraph;
        }

        /// <summary>
        /// 生成水印图片
        /// </summary>
        private static Bitmap GenerateWatermarkImage(string text, WatermarkSettings settings)
        {
            var fontSize = (int)(settings.FontSize * 2.5);
            using var font = new WinFont(settings.FontFamily, fontSize, FontStyle.Bold);

            var textSize = TextRenderer.MeasureText(text, font);
            int padding = 30;
            int width = textSize.Width + padding * 2;
            int height = textSize.Height + padding * 2;

            var bitmap = new Bitmap(width, height);
            bitmap.SetResolution(150, 150);

            using var g = Graphics.FromImage(bitmap);
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            g.Clear(WinColor.Transparent);

            using var textBrush = new SolidBrush(WinColor.FromArgb(settings.Opacity, settings.Color));
            g.DrawString(text, font, textBrush, padding, padding);

            return bitmap;
        }

        /// <summary>
        /// 移除已有的所有水印
        /// </summary>
        private static void RemoveExistingWatermarks(MainDocumentPart mainPart, Body body)
        {
            // 移除页眉中的水印
            foreach (var headerPart in mainPart.HeaderParts.ToList())
            {
                try
                {
                    var header = headerPart.Header;
                    if (header != null)
                    {
                        var picParas = header.Elements<Paragraph>()
                            .Where(p => p.Descendants<Drawing>().Any() ||
                                       p.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any())
                            .ToList();

                        foreach (var para in picParas)
                            para.Remove();

                        if (!header.HasChildren)
                            mainPart.DeletePart(headerPart);
                    }
                }
                catch { }
            }

            // 移除 Body 中的水印段落
            var watermarkParas = body.Elements<Paragraph>()
                .Where(p => IsWatermarkParagraph(p))
                .ToList();

            foreach (var para in watermarkParas)
                para.Remove();
        }

        /// <summary>
        /// 判断段落是否为水印段落
        /// </summary>
        private static bool IsWatermarkParagraph(Paragraph para)
        {
            var drawings = para.Descendants<Drawing>().ToList();
            if (!drawings.Any())
                return false;

            foreach (var drawing in drawings)
            {
                var xml = drawing.OuterXml ?? "";
                if (xml.Contains("name=\"Watermark\"") || xml.Contains("name='Watermark'"))
                    return true;
            }

            return false;
        }
    }
}
