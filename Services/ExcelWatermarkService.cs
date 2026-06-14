using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WatermarkTool.Models;

// 使用别名解决命名空间冲突
using WinFont = System.Drawing.Font;
using WinColor = System.Drawing.Color;

namespace WatermarkTool.Services
{
    /// <summary>
    /// Excel水印添加服务
    /// 通过VML绘图覆盖层方式添加水印，不修改任何单元格内容，不会触发公式重算
    /// </summary>
    public static class ExcelWatermarkService
    {
        /// <summary>
        /// 给Excel文件添加水印
        /// </summary>
        public static bool AddWatermark(string filePath, string watermarkText, WatermarkSettings settings)
        {
            try
            {
                using var watermarkImage = GenerateWatermarkImage(watermarkText, settings);

                using var spreadsheetDoc = SpreadsheetDocument.Open(filePath, true);
                var workbookPart = spreadsheetDoc.WorkbookPart;
                if (workbookPart == null)
                    return false;

                var sheets = workbookPart.Workbook.Descendants<Sheet>().ToList();

                foreach (var sheet in sheets)
                {
                    var worksheetPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart;
                    if (worksheetPart == null)
                        continue;

                    AddWatermarkToSheet(worksheetPart, watermarkImage, settings);
                }

                spreadsheetDoc.Save();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel水印添加失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成水印图片（带透明通道的PNG）
        /// </summary>
        private static Bitmap GenerateWatermarkImage(string text, WatermarkSettings settings)
        {
            var fontSize = (int)(settings.FontSize * 2);
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

            if (settings.Style == WatermarkStyle.ArtisticText)
            {
                var color = settings.Color;
                var alpha = settings.Opacity;

                using var shadowBrush = new SolidBrush(WinColor.FromArgb(alpha / 3, 50, 50, 50));
                g.DrawString(text, font, shadowBrush, padding + 3, padding + 3);

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
        /// 给单个工作表添加水印
        /// </summary>
        private static void AddWatermarkToSheet(WorksheetPart worksheetPart, Image watermarkImage, WatermarkSettings settings)
        {
            var worksheet = worksheetPart.Worksheet;
            if (worksheet == null)
                return;

            using var ms = new MemoryStream();
            watermarkImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            var imageBytes = ms.ToArray();

            var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();
            if (sheetData == null)
                return;

            var pictureWidth = watermarkImage.Width * 0.75;
            var pictureHeight = watermarkImage.Height * 0.75;

            double pageWidthInch = 8.5;
            double pageHeightInch = 11.0;

            double xInch = pageWidthInch * settings.CustomX / 100f;
            double yInch = pageHeightInch * settings.CustomY / 100f;

            var drawings = worksheet.Elements<LegacyDrawing>().ToList();
            foreach (var drawing in drawings)
            {
                try { drawing.Remove(); }
                catch { }
            }

            var existingVml = worksheetPart.GetPartsOfType<VmlDrawingPart>().ToList();
            foreach (var vml in existingVml)
            {
                try { worksheetPart.DeletePart(vml); }
                catch { }
            }

            var vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>("rIdWatermarkVml");

            var imagePart = vmlPart.AddImagePart(ImagePartType.Png, "rIdWatermarkImg");
            using (var imageStream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(imageStream);
            }
            string imageRelId = vmlPart.GetIdOfPart(imagePart);

            string vmlXml = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<xml xmlns:v=""urn:schemas-microsoft-com:vml""
     xmlns:o=""urn:schemas-microsoft-com:office:office""
     xmlns:x=""urn:schemas-microsoft-com:office:excel"">
  <o:shapelayout v:ext=""edit"">
    <o:idmap v:ext=""edit"" data=""1""/>
  </o:shapelayout>
  <v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75""
    o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">
    <v:stroke joinstyle=""miter""/>
    <v:formulas>
      <v:f eqn=""if lineDrawn pixelLineWidth 0""/>
      <v:f eqn=""sum @0 1 0""/>
      <v:f eqn=""sum 0 0 @1""/>
      <v:f eqn=""prod @2 1 2""/>
      <v:f eqn=""prod @3 21600 pixelWidth""/>
      <v:f eqn=""prod @3 21600 pixelHeight""/>
      <v:f eqn=""sum @0 0 1""/>
      <v:f eqn=""prod @6 1 2""/>
      <v:f eqn=""prod @7 21600 pixelWidth""/>
      <v:f eqn=""sum @8 21600 0""/>
      <v:f eqn=""prod @7 21600 pixelHeight""/>
      <v:f eqn=""sum @10 21600 0""/>
    </v:formulas>
    <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>
    <o:lock v:ext=""edit"" aspectratio=""t""/>
  </v:shapetype>
  <v:shape id=""_x0000_s1025"" type=""#_x0000_t75""
    style=""position:absolute;margin-left:{xInch:F2}in;margin-top:{yInch:F2}in;width:{pictureWidth:F2}pt;height:{pictureHeight:F2}pt;z-index:-1;mso-position-horizontal:absolute;mso-position-vertical:absolute""
    o:allowincell=""f"">
    <v:imagedata src=""{imageRelId}"" o:title=""watermark""/>
    <o:lock v:ext=""edit"" rotation=""t""/>
    <x:ClientData ObjectType=""Pict"">
      <x:MoveWithCells/>
      <x:SizeWithCells/>
      <x:LockText>1</x:LockText>
    </x:ClientData>
  </v:shape>
</xml>";

            using (var vmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(vmlXml)))
            {
                vmlPart.FeedData(vmlStream);
            }

            var legacyDrawing = new LegacyDrawing()
            {
                Id = new StringValue("rIdWatermarkVml")
            };

            sheetData.InsertAfterSelf(legacyDrawing);
        }
    }
}
