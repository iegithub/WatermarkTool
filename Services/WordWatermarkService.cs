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
    /// 使用 wp:anchor 将水印图片锚定到页面正文区域（非页眉），仅首页显示
    /// 不影响正文内容
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

                // 移除已有的水印（页眉中的图片）
                RemoveExistingWatermarkHeaders(mainPart);

                // 移除已有的锚定水印图片
                RemoveExistingAnchoredWatermarks(body);

                // 添加图片到 MainDocumentPart
                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var ms = new MemoryStream())
                {
                    watermarkImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Position = 0;
                    imagePart.FeedData(ms);
                }
                string imageRelId = mainPart.GetIdOfPart(imagePart);

                // 计算尺寸（EMU）
                long imageWidthEmu = watermarkImage.Width * 9525;
                long imageHeightEmu = watermarkImage.Height * 9525;

                // 计算位置（EMU）- 相对于页面
                // A4: 595440 x 841920 EMU (210mm x 297mm)
                long pageWidthEmu = 595440;
                long pageHeightEmu = 841920;

                long xEmu = (long)(pageWidthEmu * settings.CustomX / 100f) - imageWidthEmu / 2;
                long yEmu = (long)(pageHeightEmu * settings.CustomY / 100f) - imageHeightEmu / 2;

                int rotationEmu = (int)(settings.Rotation * 60000);

                // 创建锚定到页面的水印段落
                var watermarkParagraph = CreateAnchorParagraph(
                    imageRelId, imageWidthEmu, imageHeightEmu,
                    xEmu, yEmu, rotationEmu);

                // 将水印段落插入到正文最前面（第一个段落之前）
                var firstPara = body.Elements<Paragraph>().FirstOrDefault();
                if (firstPara != null)
                {
                    firstPara.InsertBeforeSelf(watermarkParagraph);
                }
                else
                {
                    body.PrependChild(watermarkParagraph);
                }

                // 确保有SectionProperties
                EnsureSectionProperties(body);

                // 设置首页页眉为空（防止首页同时显示页眉水印）
                var sections = body.Elements<SectionProperties>().ToList();
                foreach (var section in sections)
                {
                    // 添加 DifferentPageFirstHeader 以便首页使用独立页眉
                    var titlePg = section.Elements<TitlePage>().FirstOrDefault();
                    if (titlePg == null)
                    {
                        section.PrependChild(new TitlePage());
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
        /// 创建锚定到页面的水印段落（wp:anchor）
        /// </summary>
        private static Paragraph CreateAnchorParagraph(
            string imageRelId, long widthEmu, long heightEmu,
            long xEmu, long yEmu, int rotationEmu)
        {
            // 使用 wp:anchor 将图片锚定到页面
            // behindDoc=true 让水印在文字后面
            // relativeHeight=-1 确保在最底层
            // simplePos=true + posOffset 实现精确定位
            var paragraphXml = $@"
<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
     xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
     xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
     xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
     xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii=""宋体"" w:eastAsia=""宋体"" w:hAnsi=""宋体""/>
    </w:rPr>
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
        /// 生成水印图片（无阴影，半透明）
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

            // 统一使用半透明文字，无阴影
            using var textBrush = new SolidBrush(WinColor.FromArgb(settings.Opacity, settings.Color));
            g.DrawString(text, font, textBrush, padding, padding);

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
                catch { }
            }
        }

        /// <summary>
        /// 移除已有的锚定水印图片（wp:anchor）
        /// </summary>
        private static void RemoveExistingAnchoredWatermarks(Body body)
        {
            var paragraphs = body.Elements<Paragraph>().ToList();
            foreach (var para in paragraphs)
            {
                var drawings = para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().ToList();
                if (drawings.Any())
                {
                    // 检查是否是水印（通过name属性）
                    var isWatermark = false;
                    foreach (var drawing in drawings)
                    {
                        var xml = drawing.OuterXml ?? "";
                        if (xml.Contains("name=\"Watermark\"") || xml.Contains("name='Watermark'"))
                        {
                            isWatermark = true;
                            break;
                        }
                    }
                    if (isWatermark)
                    {
                        para.Remove();
                    }
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
    }
}
