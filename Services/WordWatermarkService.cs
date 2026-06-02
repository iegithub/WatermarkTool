using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WatermarkTool.Models;

using WinFont = System.Drawing.Font;
using WinColor = System.Drawing.Color;

namespace WatermarkTool.Services
{
    /// <summary>
    /// Word文档水印添加服务
    /// 使用 OpenXML SDK 创建锚定到页面的水印图片
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
        /// 创建水印段落 - 使用 XML 字符串
        /// </summary>
        private static Paragraph CreateWatermarkParagraph(
            string imageRelId, long widthEmu, long heightEmu,
            long xEmu, long yEmu, int rotationEmu)
        {
            var paragraphXml = $@"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" 
                xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" 
                xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" 
                xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" 
                xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                <w:r>
                    <w:drawing>
                        <wp:anchor distT=""0"" distB=""0"" distL=""0"" distR=""0"" 
                            simplePos=""0"" relativeHeight=""-251658240"" 
                            behindDoc=""true"" locked=""false"" 
                            layoutInCell=""true"" allowOverlap=""true"">
                            <wp:simplePos x=""0"" y=""0""/>
                            <wp:positionH relativeFrom=""page"">
                                <wp:posOffset>{xEmu}</wp:posOffset>
                            </wp:positionH>
                            <wp:positionV relativeFrom=""page"">
                                <wp:posOffset>{yEmu}</wp:posOffset>
                            </wp:positionV>
                            <wp:extent cx=""{widthEmu}"" cy=""{heightEmu}""/>
                            <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>
                            <wp:wrapNone/>
                            <wp:docPr id=""1"" name=""Watermark""/>
                            <wp:cNvGraphicFramePr>
                                <a:graphicFrameLocks noChangeAspect=""1""/>
                            </wp:cNvGraphicFramePr>
                            <a:graphic>
                                <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                                    <pic:pic>
                                        <pic:nvPicPr>
                                            <pic:cNvPr id=""0"" name=""Watermark.png""/>
                                            <pic:cNvPicPr/>
                                        </pic:nvPicPr>
                                        <pic:blipFill>
                                            <a:blip r:embed=""{imageRelId}""/>
                                            <a:stretch>
                                                <a:fillRect/>
                                            </a:stretch>
                                        </pic:blipFill>
                                        <pic:spPr>
                                            <a:xfrm rot=""{rotationEmu}"">
                                                <a:off x=""0"" y=""0""/>
                                                <a:ext cx=""{widthEmu}"" cy=""{heightEmu}""/>
                                            </a:xfrm>
                                            <a:prstGeom prst=""rect"">
                                                <a:avLst/>
                                            </a:prstGeom>
                                        </pic:spPr>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:anchor>
                    </w:drawing>
                </w:r>
            </w:p>";

            return new Paragraph(paragraphXml);
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
