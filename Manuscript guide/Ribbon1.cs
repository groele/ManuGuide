using Microsoft.Office.Tools.Ribbon;
using Manuscript_guide.Services;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace Manuscript_guide
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            btnDash.Image = CreateDashMixIcon();
            btnDetectAll.Image = CreateDetectAllIcon();
            btnSettings.Image = CreateSettingsIcon();
            btnDash.ShowImage = true;
            btnDetectAll.ShowImage = true;
            btnSettings.ShowImage = true;
            ConfigureButtonTips();
        }

        private void ConfigureButtonTips()
        {
            btnPunc.ScreenTip = "标点与格式检测";
            btnPunc.SuperTip = "检查中英文标点、半角空格、微符号编码和公式周边标点。\n示例：修改前 1 , 2 / μm 编码混用；修改后 1, 2 / μm 统一规范。";

            btnCap.ScreenTip = "大小写规范检测";
            btnCap.SuperTip = "检查缩写首次定义、大小写一致性和图表/文献引用大小写。\n示例：修改前 fig. 2 或 moS2；修改后 Fig. 2 或 MoS2。";

            btnDash.ScreenTip = "横线与混排检测";
            btnDash.SuperTip = "检查连字符、en dash、em dash、复合修饰语和中英文混排空格。\n示例：修改前 5 - 10 nm、MoS2 based；修改后 5-10 nm / 5–10 nm、MoS2-based。";

            btnData.ScreenTip = "数据与数值检测";
            btnData.SuperTip = "检查数值单位空格、科学计数法、范围表达和数学符号。\n示例：修改前 10nm、1 x 10-3；修改后 10 nm、1 × 10^-3。";

            btnItal.ScreenTip = "变量与斜体检测";
            btnItal.SuperTip = "检查变量斜体、函数正体和不应斜体的学术短语。\n示例：修改前 E, sin, in situ 格式混乱；修改后变量斜体、函数/短语保持正体。";

            btnSub.ScreenTip = "角标与改写检测";
            btnSub.SuperTip = "检查化学式、单位幂次、描述性角标和 LaTeX 风格行内角标。\n示例：修改前 Al2O3、cm-3；修改后 Al₂O₃、cm⁻³，或保留 ASCII 并应用 Word 原生角标。";

            btnWord.ScreenTip = "语病与措辞检测";
            btnWord.SuperTip = "检查常见学术语病、冗余措辞和英文表达风险。\n示例：修改前 due to the fact that；修改后 because。";

            btnKeyword.ScreenTip = "关键词标记";
            btnKeyword.SuperTip = "打开关键词检索面板，对指定术语进行全文高亮标记。\n示例：输入 PL, strain 后，会在正文中标出全部匹配位置。";

            btnDetectAll.ScreenTip = "一键检测";
            btnDetectAll.SuperTip = "连续运行全部专项检测，只生成问题队列和高亮标记，不直接修改正文。\n适合先审阅修改前/建议修改后的差异，再逐条接受。";

            btnFixAll.ScreenTip = "一键修复";
            btnFixAll.SuperTip = "先运行全部专项检测，再批量应用可修复建议，并保留淡色底纹作为插件留痕。\n修改前后可在侧边栏复核。";

            btnClear.ScreenTip = "清除标记";
            btnClear.SuperTip = "清除插件生成的诊断底纹和标记。\n是否保留用户手动高亮，取决于“高级设置”中的安全选项。";

            btnSettings.ScreenTip = "高级设置";
            btnSettings.SuperTip = "打开规则、颜色、白名单、阈值和导入导出配置。\n用于调整检测敏感度、留痕颜色和团队共享规则。";
        }

        private void DiagnosticButton_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton button = sender as RibbonButton;
            if (button == null)
            {
                return;
            }

            switch (button.Name)
            {
                case "btnPunc":
                    EventBus.TriggerScan("punc");
                    break;
                case "btnCap":
                    EventBus.TriggerScan("cap");
                    break;
                case "btnDash":
                    EventBus.TriggerScan("dash");
                    break;
                case "btnData":
                    EventBus.TriggerScan("data");
                    break;
                case "btnItal":
                    EventBus.TriggerScan("ital");
                    break;
                case "btnSub":
                    EventBus.TriggerScan("sub");
                    break;
                case "btnWord":
                    EventBus.TriggerScan("word");
                    break;
            }
        }

        private void BtnKeyword_Click(object sender, RibbonControlEventArgs e)
        {
            EventBus.TriggerScan("keyword");
        }

        private void BtnDetectAll_Click(object sender, RibbonControlEventArgs e)
        {
            EventBus.TriggerFullScan();
        }

        private void BtnFixAll_Click(object sender, RibbonControlEventArgs e)
        {
            EventBus.TriggerFullFix();
        }

        private void BtnClear_Click(object sender, RibbonControlEventArgs e)
        {
            EventBus.TriggerClearHighlights("all");
        }

        private void BtnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            EventBus.TriggerOpenSettings();
        }

        private Image CreateDashMixIcon()
        {
            const int scale = 4;
            const int size = 32;
            Bitmap canvas = new Bitmap(size * scale, size * scale);
            using (Graphics g = Graphics.FromImage(canvas))
            using (Pen textPen = new Pen(Color.FromArgb(96, 94, 92), 1.8f * scale))
            using (Pen faintPen = new Pen(Color.FromArgb(196, 196, 196), 1.2f * scale))
            using (Pen dashPen = new Pen(Color.FromArgb(0, 120, 212), 2.8f * scale))
            using (SolidBrush paperBrush = new SolidBrush(Color.FromArgb(250, 250, 250)))
            using (SolidBrush shadowBrush = new SolidBrush(Color.FromArgb(34, 0, 0, 0)))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                FillRoundedRectangle(g, shadowBrush, 5.5f, 5.8f, 22.5f, 21.2f, 3.2f, scale);
                FillRoundedRectangle(g, paperBrush, 4.5f, 4.5f, 22.5f, 21.5f, 3.2f, scale);
                using (Pen borderPen = new Pen(Color.FromArgb(186, 186, 186), 1.1f * scale))
                {
                    DrawRoundedRectangle(g, borderPen, 4.5f, 4.5f, 22.5f, 21.5f, 3.2f, scale);
                }

                DrawRoundedLine(g, faintPen, 9, 10, 23, 10, scale);
                DrawRoundedLine(g, textPen, 9, 15, 13.5f, 15, scale);
                DrawRoundedLine(g, dashPen, 15.4f, 15, 20.6f, 15, scale);
                DrawRoundedLine(g, textPen, 22.5f, 15, 27, 15, scale);
                DrawRoundedLine(g, faintPen, 9, 20, 23, 20, scale);
            }

            Bitmap bitmap = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);
                g.DrawImage(canvas, 0, 0, size, size);
            }

            canvas.Dispose();
            return bitmap;
        }

        private static void DrawRoundedLine(Graphics g, Pen pen, float x1, float y1, float x2, float y2, int scale)
        {
            LineCap originalStart = pen.StartCap;
            LineCap originalEnd = pen.EndCap;
            pen.StartCap = LineCap.Round;
            pen.EndCap = LineCap.Round;
            g.DrawLine(pen, x1 * scale, y1 * scale, x2 * scale, y2 * scale);
            pen.StartCap = originalStart;
            pen.EndCap = originalEnd;
        }

        private static GraphicsPath CreateRoundedRectangle(float x, float y, float width, float height, float radius, int scale)
        {
            float sx = x * scale;
            float sy = y * scale;
            float sw = width * scale;
            float sh = height * scale;
            float sr = radius * scale;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(sx, sy, sr, sr, 180, 90);
            path.AddArc(sx + sw - sr, sy, sr, sr, 270, 90);
            path.AddArc(sx + sw - sr, sy + sh - sr, sr, sr, 0, 90);
            path.AddArc(sx, sy + sh - sr, sr, sr, 90, 90);
            path.CloseFigure();
            return path;
        }

        private static void FillRoundedRectangle(Graphics g, Brush brush, float x, float y, float width, float height, float radius, int scale)
        {
            using (GraphicsPath path = CreateRoundedRectangle(x, y, width, height, radius, scale))
            {
                g.FillPath(brush, path);
            }
        }

        private static void DrawRoundedRectangle(Graphics g, Pen pen, float x, float y, float width, float height, float radius, int scale)
        {
            using (GraphicsPath path = CreateRoundedRectangle(x, y, width, height, radius, scale))
            {
                g.DrawPath(pen, path);
            }
        }

        private Image CreateSettingsIcon()
        {
            const int scale = 4;
            const int size = 32;
            Bitmap canvas = new Bitmap(size * scale, size * scale);
            using (Graphics g = Graphics.FromImage(canvas))
            using (Pen linePen = new Pen(Color.FromArgb(96, 94, 92), 1.8f * scale))
            using (Pen accentPen = new Pen(Color.FromArgb(0, 120, 212), 2.2f * scale))
            using (SolidBrush knobBrush = new SolidBrush(Color.FromArgb(0, 120, 212)))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                DrawRoundedLine(g, linePen, 7, 10, 25, 10, scale);
                DrawRoundedLine(g, linePen, 7, 16, 25, 16, scale);
                DrawRoundedLine(g, linePen, 7, 22, 25, 22, scale);

                g.FillEllipse(knobBrush, 13.5f * scale, 7.5f * scale, 5f * scale, 5f * scale);
                g.FillEllipse(knobBrush, 20.5f * scale, 13.5f * scale, 5f * scale, 5f * scale);
                g.FillEllipse(knobBrush, 9.5f * scale, 19.5f * scale, 5f * scale, 5f * scale);

                g.DrawEllipse(accentPen, 13.5f * scale, 7.5f * scale, 5f * scale, 5f * scale);
                g.DrawEllipse(accentPen, 20.5f * scale, 13.5f * scale, 5f * scale, 5f * scale);
                g.DrawEllipse(accentPen, 9.5f * scale, 19.5f * scale, 5f * scale, 5f * scale);
            }

            Bitmap bitmap = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);
                g.DrawImage(canvas, 0, 0, size, size);
            }

            canvas.Dispose();
            return bitmap;
        }

        private Image CreateDetectAllIcon()
        {
            const int scale = 4;
            const int size = 32;
            Bitmap canvas = new Bitmap(size * scale, size * scale);
            using (Graphics g = Graphics.FromImage(canvas))
            using (SolidBrush paperBrush = new SolidBrush(Color.FromArgb(250, 250, 250)))
            using (SolidBrush shadowBrush = new SolidBrush(Color.FromArgb(30, 0, 0, 0)))
            using (Pen borderPen = new Pen(Color.FromArgb(186, 186, 186), 1.1f * scale))
            using (Pen textPen = new Pen(Color.FromArgb(150, 150, 150), 1.2f * scale))
            using (Pen lensPen = new Pen(Color.FromArgb(0, 120, 212), 2.5f * scale))
            using (Pen checkPen = new Pen(Color.FromArgb(16, 124, 65), 2.4f * scale))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                FillRoundedRectangle(g, shadowBrush, 5.8f, 4.8f, 16.8f, 21.8f, 2.6f, scale);
                FillRoundedRectangle(g, paperBrush, 5, 4, 16.8f, 21.8f, 2.6f, scale);
                DrawRoundedRectangle(g, borderPen, 5, 4, 16.8f, 21.8f, 2.6f, scale);

                DrawRoundedLine(g, textPen, 9, 10, 18, 10, scale);
                DrawRoundedLine(g, textPen, 9, 14, 17, 14, scale);
                DrawRoundedLine(g, textPen, 9, 18, 15, 18, scale);

                g.DrawEllipse(lensPen, 15.2f * scale, 13.3f * scale, 9.8f * scale, 9.8f * scale);
                DrawRoundedLine(g, lensPen, 22.7f, 21.2f, 27, 25.5f, scale);
                DrawRoundedLine(g, checkPen, 16.8f, 18.8f, 19.2f, 21.2f, scale);
                DrawRoundedLine(g, checkPen, 19.2f, 21.2f, 23.8f, 16.6f, scale);
            }

            Bitmap bitmap = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);
                g.DrawImage(canvas, 0, 0, size, size);
            }

            canvas.Dispose();
            return bitmap;
        }
    }
}
