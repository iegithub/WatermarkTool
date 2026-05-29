using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WatermarkTool.Models;

// 使用别名解决命名空间冲突
using WinFont = System.Drawing.Font;
using WinColor = System.Drawing.Color;

namespace WatermarkTool.Services
{
    /// <summary>
    /// Word文档水印添加服务
    /// 使用OpenXML直接操作docx文件，通过页眉添加水印图片，不影响正文内容
    /// </summary>
    public static class WordWatermarkService
    {
        /// <summary>
        /// 给Word文档添加水印
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

                RemoveExistingWatermarkHeaders(mainPart);

                var headerPart = mainPart.AddNewPart<HeaderPart>("rIdWatermarkHeader");
                GenerateHeaderPart(headerPart, watermarkImage, settings);

                var body = mainPart.Document.Body;
                if (body == null)
                    return false;

                EnsureSectionProperties(body);
                var sections = body.Elements<SectionProperties>().ToList();

                foreach (var section in sections)
                {
                    var headerReference = section.Elements<HeaderReference>()
                        .FirstOrDefault(h => h.Type != null && h.Type.Value == HeaderFooterValues.Default);

                    if (headerReference != null)
                    {
                        headerReference.Id = "rIdWatermarkHeader";
                    }
                    else
                    {
                        section.PrependChild(new HeaderReference
                        {
                            Type = HeaderFooterValues.Default,
                            Id = "rIdWatermarkHeader"
                        });
                    }

                    var firstHeaderRef = section.Elements<HeaderReference>()
                        .FirstOrDefault(h => h.Type != null && h.Type.Value == HeaderFooterValues.First);

                    if (firstHeaderRef != null)
                    {
                        firstHeaderRef.Id = "rIdWatermarkHeader";
                    }
                }

                mainPart.Document.Save();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Word水印添加失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成水印图片
        /// </summary>
        private static Bitmap GenerateWatermarkImage(string text, WatermarkSettings settings)
        {
            var fontSize = (int)(settings.FontSize * 2.5);
            using var font = new WinFont(settings.FontFamily, fontSize, FontStyle.Bold);

            var textSize = TextRenderer.MeasureText(text, font);
            int padding = 40;
            int width = textSize.Width + padding * 2;
            int height = textSize.Height + padding * 2;

            var bitmap = new Bitmap(width, height);
            bitmap.SetResolution(150, 150);

            using var g = Graphics.FromImage(bitmap);
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            g.Clear(WinColor.Transparent);

            if (settings.Style == WatermarkStyle.ArtisticText)
            {
                var color = settings.Color;
                var alpha = settings.Opacity;

                using var shadowBrush = new SolidBrush(WinColor.FromArgb(alpha / 3, 50, 50, 50));
                g.DrawString(text, font, shadowBrush, padding + 4, padding + 4);

                using var gradBrush = new LinearGradientBrush(
                    new Point(padding, padding),
                    new Point(padding + textSize.Width, padding + textSize.Height),
                    WinColor.FromArgb(alpha, color),
                    WinColor.FromArgb(alpha, WinColor.FromArgb(
                        Math.Max(0, color.R - 50),
                        Math.Max(0, color.G - 50),
                        Math.Max(0, color.B - 50))));
                g.DrawString(text, font, gradBrush, padding, padding);

                using var path = new GraphicsPath();
                path.AddString(text, font.FontFamily, (int)font.Style, font.Size,
                    new PointF(padding, padding), StringFormat.GenericDefault);
                using var outlinePen = new Pen(WinColor.FromArgb(alpha, color.R / 2, color.G / 2, color.B / 2), 1.5f);
                g.DrawPath(outlinePen, path);
            }
            else
            {
                using var textBrush = new SolidBrush(WinColor.FromArgb(settings.Opacity, settings.Color));
                g.DrawString(text, font, textBrush, padding, padding);
            }

            return bitmap;
        }

        /// <summary>
        /// 移除已有的水印页眉
        /// </summary>
        private static void RemoveExistingWatermarkHeaders(MainDocumentPart mainPart)
        {
            var headerParts = mainPart.HeaderParts.ToList();
            foreach (var headerPart in headerParts)
            {
                try
                {
                    var headerElement = headerPart.Header;
                    if (headerElement != null)
                    {
                        var hasImage = headerElement.Descendants<Drawing>().Any() ||
                                       headerElement.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any();

                        if (hasImage)
                        {
                            mainPart.DeletePart(headerPart);
                        }
                    }
                }
                catch
                {
                    // 忽略无法访问的部分
                }
            }
        }

        /// <summary>
        /// 确保文档有SectionProperties
        /// </summary>
        private static void EnsureSectionProperties(Body body)
        {
            if (!body.Elements<SectionProperties>().Any())
            {
                body.AppendChild(new SectionProperties());
            }
        }

        /// <summary>
        /// 生成页眉部分内容（包含水印图片）
        /// </summary>
        private static void GenerateHeaderPart(HeaderPart headerPart, Image watermarkImage, WatermarkSettings settings)
        {
            using var ms = new MemoryStream();
            watermarkImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            var imageBytes = ms.ToArray();

            var imagePart = headerPart.AddImagePart(ImagePartType.Png);
            using (var imageStream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(imageStream);
            }
            string imageRelId = headerPart.GetIdOfPart(imagePart);

            long pageWidthEmu = 595440;
            long pageHeightEmu = 841920;

            long imageWidthEmu = watermarkImage.Width * 9525;
            long imageHeightEmu = watermarkImage.Height * 9525;

            long xEmu = (long)(pageWidthEmu * settings.CustomX / 100f) - imageWidthEmu / 2;
            long yEmu = (long)(pageHeightEmu * settings.CustomY / 100f) - imageHeightEmu / 2;

            int rotationEmu = (int)(settings.Rotation * 60000);

            var headerXml = $@"
<w:hdr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
        xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
        xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
        xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
        xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
  <w:p>
    <w:pPr>
      <w:pStyle w:val=""Header""/>
    </w:pPr>
    <w:r>
      <w:drawing>
        <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"">
          <wp:extent cx=""{imageWidthEmu}"" cy=""{imageHeightEmu}""/>
          <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>
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
                    <a:off x=""{xEmu}"" y=""{yEmu}""/>
                    <a:ext cx=""{imageWidthEmu}"" cy=""{imageHeightEmu}""/>
                  </a:xfrm>
                  <a:prstGeom prst=""rect"">
                    <a:avLst/>
                  </a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>
</w:hdr>";

            headerPart.Header = new Header(headerXml);
        }
    }
}
