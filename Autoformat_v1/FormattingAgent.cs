using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Autoformat_v1
{
    public class FormattingAgent
    {
        public static void ApplyDesign(string designName, bool addIcons = false)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.ActiveWindow.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select a cluster of shapes first.");
                    return;
                }

                var shapes = selection.ShapeRange;
                
                // If single text box selected, split it into multiple rows
                if (shapes.Count == 1 && shapes[1].HasTextFrame == MsoTriState.msoTrue)
                {
                    var original = shapes[1];
                    var tr = original.TextFrame.TextRange;
                    if (tr.Paragraphs().Count > 1)
                    {
                        float currentTop = original.Top;
                        List<PowerPoint.Shape> newShapes = new List<PowerPoint.Shape>();
                        for (int i = 1; i <= tr.Paragraphs().Count; i++)
                        {
                            var p = tr.Paragraphs(i);
                            string text = p.Text.Trim();
                            if (string.IsNullOrEmpty(text) || text == "\r" || text == "\n") continue;
                            
                            var newShape = app.ActiveWindow.View.Slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, original.Left, currentTop, original.Width, 30);
                            newShape.TextFrame.TextRange.Text = text;
                            try {
                                newShape.TextFrame.TextRange.Font.Name = p.Font.Name;
                                newShape.TextFrame.TextRange.Font.Size = p.Font.Size;
                                newShape.TextFrame.TextRange.Font.Color.RGB = p.Font.Color.RGB;
                                newShape.TextFrame.TextRange.Font.Bold = p.Font.Bold;
                            } catch { }
                            newShape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                            newShapes.Add(newShape);
                            currentTop += 40;
                        }
                        if (newShapes.Count > 1)
                        {
                            original.Delete();
                            newShapes[0].Select(MsoTriState.msoTrue);
                            for(int i = 1; i < newShapes.Count; i++) newShapes[i].Select(MsoTriState.msoFalse);
                            shapes = app.ActiveWindow.Selection.ShapeRange;
                        }
                    }
                }

                if (shapes.Count < 2)
                {
                    MessageBox.Show("Please select a cluster of multiple shapes (header and body), or a single text box with multiple lines.");
                    return;
                }

                // 1. Analyze Selection and Original Bounds
                var userCluster = SelectionAnalyzer.Analyze(shapes);

                // 2. Load Template
                string libraryPath = @"C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\Library.pptx";
                if (!File.Exists(libraryPath))
                {
                    string basePath = app.ActivePresentation.Path;
                    if (basePath.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                        libraryPath = basePath.TrimEnd('/') + "/Library.pptx";
                    else
                        libraryPath = Path.Combine(basePath, "Library.pptx");
                }
                
                PowerPoint.Presentation library = null;
                try
                {
                    library = app.Presentations.Open(libraryPath, ReadOnly: MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);
                    var templateCluster = TemplateManager.ExtractFirstTemplate(library, designName);
                    if (templateCluster == null)
                    {
                        MessageBox.Show("Could not find a valid template cluster in Library.pptx.");
                        return;
                    }

                    // 3. Apply Layout
                    LayoutEngine.ApplyLayout(app.ActiveWindow.View.Slide, userCluster, templateCluster, designName, addIcons);
                }
                finally
                {
                    if (library != null)
                    {
                        library.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error applying design: " + ex.Message + "\n" + ex.StackTrace);
            }
        }
    }

    public class UserCluster
    {
        public PowerPoint.Shape Number { get; set; }
        public PowerPoint.Shape Header { get; set; }
        public List<BodyRow> Rows { get; set; }
        public PowerPoint.Shape Background { get; set; }
        public float BoundsLeft { get; set; }
        public float BoundsTop { get; set; }
        public float BoundsWidth { get; set; }
        public float BoundsHeight { get; set; }
    }

    public class BodyRow
    {
        public PowerPoint.Shape Icon { get; set; }
        public PowerPoint.Shape Subheader { get; set; }
        public PowerPoint.Shape Text { get; set; }
    }

    public class TemplateCluster
    {
        public PowerPoint.Shape NumberPrototype { get; set; }
        public PowerPoint.Shape Header { get; set; }
        public PowerPoint.Shape SeparatorPrototype { get; set; }
        public PowerPoint.Shape SubheaderPrototype { get; set; }
        public PowerPoint.Shape BodyTextPrototype { get; set; }
        public PowerPoint.Shape IconPrototype { get; set; }
        public PowerPoint.Shape Background { get; set; }
    }

    public static class SelectionAnalyzer
    {
        public static UserCluster Analyze(PowerPoint.ShapeRange shapeRange)
        {
            List<PowerPoint.Shape> allShapes = new List<PowerPoint.Shape>();
            float minL = float.MaxValue, minT = float.MaxValue, maxR = float.MinValue, maxB = float.MinValue;

            for (int i = 1; i <= shapeRange.Count; i++)
            {
                var s = shapeRange[i];
                allShapes.Add(s);
                if (s.Left < minL) minL = s.Left;
                if (s.Top < minT) minT = s.Top;
                if (s.Left + s.Width > maxR) maxR = s.Left + s.Width;
                if (s.Top + s.Height > maxB) maxB = s.Top + s.Height;
            }

            var cluster = new UserCluster();
            cluster.Rows = new List<BodyRow>();
            cluster.BoundsLeft = minL;
            cluster.BoundsTop = minT;
            cluster.BoundsWidth = maxR - minL;
            cluster.BoundsHeight = maxB - minT;

            // Identify potential Background
            var largestShape = allShapes.OrderByDescending(s => s.Width * s.Height).First();
            if (allShapes.Count > 2 && largestShape.Width > allShapes.Average(s => s.Width) * 1.5)
            {
                cluster.Background = largestShape;
                allShapes.Remove(largestShape);
            }

            // Identify Header
            cluster.Header = allShapes.OrderBy(s => s.Top).First();
            var bodyShapes = allShapes.Where(s => s.Id != cluster.Header.Id).OrderBy(s => s.Top).ToList();

            // Group into Rows
            List<PowerPoint.Shape> currentRow = new List<PowerPoint.Shape>();
            foreach (var s in bodyShapes)
            {
                if (currentRow.Count == 0)
                {
                    currentRow.Add(s);
                }
                else
                {
                    var prev = currentRow.First();
                    float midY = s.Top + (s.Height / 2);
                    if (midY >= prev.Top && midY <= prev.Top + prev.Height)
                    {
                        currentRow.Add(s);
                    }
                    else
                    {
                        AddRowToCluster(cluster, currentRow);
                        currentRow = new List<PowerPoint.Shape> { s };
                    }
                }
            }
            if (currentRow.Count > 0)
            {
                AddRowToCluster(cluster, currentRow);
            }

            return cluster;
        }

        private static void AddRowToCluster(UserCluster cluster, List<PowerPoint.Shape> rowShapes)
        {
            var sorted = rowShapes.OrderBy(x => x.Left).ToList();
            var bodyRow = new BodyRow();
            if (sorted.Count == 1)
            {
                bodyRow.Text = sorted[0];
            }
            else
            {
                bodyRow.Icon = sorted[0];
                bodyRow.Text = sorted[1];
            }
            cluster.Rows.Add(bodyRow);
        }
    }

    public static class TemplateManager
    {
        public static TemplateCluster ExtractFirstTemplate(PowerPoint.Presentation library, string designName = "Design 1")
        {
            var cluster = new TemplateCluster();
            PowerPoint.Slide foundSlide = null;

            // Search across all slides for the requested design
            foreach (PowerPoint.Slide libSlide in library.Slides)
            {
                cluster.Background = FindShapeByName(libSlide.Shapes, designName, "Decor");
                if (cluster.Background != null)
                {
                    foundSlide = libSlide;
                    cluster.NumberPrototype = FindShapeByName(libSlide.Shapes, designName, "Number");
                    cluster.Header = FindShapeByName(libSlide.Shapes, designName, "Header");
                    cluster.BodyTextPrototype = FindShapeByName(libSlide.Shapes, designName, "Body Text");
                    cluster.IconPrototype = FindShapeByName(libSlide.Shapes, designName, "Icon");
                    cluster.SeparatorPrototype = FindShapeByName(libSlide.Shapes, designName, "Divider");
                    cluster.SubheaderPrototype = FindShapeByName(libSlide.Shapes, designName, "Sub header");
                    if (cluster.SubheaderPrototype == null)
                        cluster.SubheaderPrototype = FindShapeByName(libSlide.Shapes, designName, "Subheader");
                    break; // Found the design, stop searching slides
                }
            }

            if (foundSlide == null && library.Slides.Count > 0)
            {
                foundSlide = library.Slides[1];
            }

            if (foundSlide != null && cluster.SeparatorPrototype == null)
            {
                // Unified fallback for all designs
                cluster.SeparatorPrototype = FindShapeByName(foundSlide.Shapes, "Grey Side Fade R");
                if (cluster.SeparatorPrototype == null)
                    cluster.SeparatorPrototype = FindShapeByName(foundSlide.Shapes, "Grey Centre Fade H Divider");
            }

            // Fallback for Background if not found
            if (cluster.Background == null && foundSlide != null)
            {
                var shapes = GetTemplateShapes(foundSlide.Shapes);
                if (shapes.Count > 0)
                {
                    var largestShape = shapes.OrderByDescending(s => s.Width * s.Height).First();
                    if (largestShape.Width > shapes.Average(s => s.Width) * 1.5)
                    {
                        cluster.Background = largestShape;
                        shapes.Remove(largestShape);
                    }
                }
            }

            // Fallback for Header
            if (cluster.Header == null && foundSlide != null)
            {
                var shapes = GetTemplateShapes(foundSlide.Shapes);
                cluster.Header = shapes.OrderBy(s => s.Top).FirstOrDefault();
            }

            // Fallback for Divider
            if (cluster.SeparatorPrototype == null && library != null)
            {
                foreach (PowerPoint.Slide s in library.Slides)
                {
                    cluster.SeparatorPrototype = FindShapeByName(s.Shapes, "Grey Side Fade R");
                    if (cluster.SeparatorPrototype != null) break;
                }
                
                // Final fallback if even Grey Side Fade R isn't in the library
                if (cluster.SeparatorPrototype == null && foundSlide != null)
                {
                    var shapes = GetTemplateShapes(foundSlide.Shapes);
                    cluster.SeparatorPrototype = shapes.FirstOrDefault(s => s.Type == MsoShapeType.msoLine && s.Width > s.Height);
                }
            }

            // Fallback for BodyText
            if (cluster.BodyTextPrototype == null && foundSlide != null)
            {
                var shapes = GetTemplateShapes(foundSlide.Shapes);
                cluster.BodyTextPrototype = shapes.FirstOrDefault(s => 
                    s.HasTextFrame == MsoTriState.msoTrue && 
                    s.TextFrame.HasText == MsoTriState.msoTrue && 
                    s.Id != cluster.Header?.Id);
            }

            return cluster;
        }

        private static PowerPoint.Shape FindShapeByName(PowerPoint.Shapes shapes, string kw1, string kw2 = null)
        {
            foreach (PowerPoint.Shape s in shapes)
            {
                string norm = s.Name.Replace(" ", "").ToLower();
                bool match1 = norm.Contains(kw1.Replace(" ", "").ToLower());
                bool match2 = kw2 == null || norm.Contains(kw2.Replace(" ", "").ToLower());
                if (match1 && match2) return s;

                if (s.Type == MsoShapeType.msoGroup)
                {
                    var child = FindShapeByNameGroup(s.GroupItems, kw1, kw2);
                    if (child != null) return child;
                }
            }
            return null;
        }

        private static PowerPoint.Shape FindShapeByNameGroup(PowerPoint.GroupShapes shapes, string kw1, string kw2 = null)
        {
            foreach (PowerPoint.Shape s in shapes)
            {
                string norm = s.Name.Replace(" ", "").ToLower();
                bool match1 = norm.Contains(kw1.Replace(" ", "").ToLower());
                bool match2 = kw2 == null || norm.Contains(kw2.Replace(" ", "").ToLower());
                if (match1 && match2) return s;

                if (s.Type == MsoShapeType.msoGroup)
                {
                    var child = FindShapeByNameGroup(s.GroupItems, kw1, kw2);
                    if (child != null) return child;
                }
            }
            return null;
        }

        private static List<PowerPoint.Shape> GetTemplateShapes(PowerPoint.Shapes shapes)
        {
            float redLineLeft = float.MaxValue;
            foreach (PowerPoint.Shape s in shapes)
            {
                if (s.Type == MsoShapeType.msoLine)
                {
                    if (s.Height > s.Width)
                    {
                        if (s.Line.ForeColor.RGB == 255 || s.Line.ForeColor.RGB == 0x0000FF || (s.Line.ForeColor.RGB & 0xFF) == 255)
                        {
                            if (s.Left < redLineLeft) redLineLeft = s.Left;
                        }
                    }
                }
            }
            if (redLineLeft == float.MaxValue) redLineLeft = 9999;
            List<PowerPoint.Shape> templateShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape s in shapes)
            {
                if (s.Left + (s.Width / 2) < redLineLeft)
                {
                    if (!(s.Type == MsoShapeType.msoLine && s.Height > s.Width))
                    {
                        templateShapes.Add(s);
                    }
                }
            }
            return templateShapes;
        }
    }

    public static class LayoutEngine
    {
        public static void ApplyLayout(PowerPoint.Slide slide, UserCluster user, TemplateCluster template, string designName, bool addIcons)
        {
            float bLeft = user.BoundsLeft;
            float bTop = user.BoundsTop;
            float bWidth = user.BoundsWidth;
            float bHeight = user.BoundsHeight;
            
            // Re-map UserCluster logic for KPI Cards (where Number is top, then Header, then Body)
            if (template.NumberPrototype != null)
            {
                user.Number = user.Header;
                user.Header = null;
                if (user.Rows != null && user.Rows.Count > 0)
                {
                    user.Header = user.Rows[0].Text;
                    user.Rows.RemoveAt(0);
                }
            }

            // Calculate relative positioning from the template to preserve the exact design
            float headerOffsetLeft = 15f;
            float headerOffsetRight = 15f;
            float bgOffsetTop = 15f;
            
            if (template.Background != null && template.Header != null)
            {
                headerOffsetLeft = template.Header.Left - template.Background.Left;
                headerOffsetRight = (template.Background.Left + template.Background.Width) - (template.Header.Left + template.Header.Width);
                bgOffsetTop = template.Background.Top - template.Header.Top;
            }

            // The header sits at the top of the user's bounding box
            float finalHeaderTop = bTop;
            if (bgOffsetTop < 0) finalHeaderTop = bTop - bgOffsetTop; // If template BG is higher than Header

            float finalBgTop = bTop;
            if (bgOffsetTop > 0) finalBgTop = bTop + bgOffsetTop; // Header protrudes above Decor Box
            
            float finalBgLeft = bLeft;
            float finalBgWidth = bWidth;
            
            // Base height on template's aspect ratio to preserve the original design ratio
            float finalBgHeight = bHeight - (finalBgTop - bTop);
            if (template.Background != null && template.Background.Width > 0)
            {
                float templateRatio = template.Background.Height / template.Background.Width;
                finalBgHeight = finalBgWidth * templateRatio;
            }
            if (finalBgHeight < 50) finalBgHeight = 50;

            // 1. Create/Format Background Container
            PowerPoint.Shape bgContainer = null;
            if (template.Background != null)
            {
                if (user.Background != null) user.Background.Delete();
                template.Background.Copy();
                var pasted = slide.Shapes.Paste();
                bgContainer = pasted[1];
                bgContainer.Left = finalBgLeft;
                bgContainer.Top = finalBgTop;
                bgContainer.Width = finalBgWidth;
                bgContainer.Height = finalBgHeight;
                bgContainer.ZOrder(MsoZOrderCmd.msoSendToBack);
            }

            float targetHeaderSize = 12f;
            if (template.Header != null) try { targetHeaderSize = template.Header.TextFrame.TextRange.Paragraphs(1).Font.Size; } catch {}
            if (targetHeaderSize <= 0) targetHeaderSize = 12f;
            
            float targetBodySize = 11f; // Always 11 as explicitly requested

            // Format Number for KPI Card
            if (user.Number != null && template.NumberPrototype != null)
            {
                user.Number.Left = finalBgLeft + headerOffsetLeft;
                user.Number.Top = finalHeaderTop;
                user.Number.Width = finalBgWidth - headerOffsetLeft - headerOffsetRight;
                
                try {
                    user.Number.TextFrame.WordWrap = MsoTriState.msoTrue;
                    user.Number.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    
                    var tFont = template.NumberPrototype.TextFrame.TextRange.Font;
                    var uFont = user.Number.TextFrame.TextRange.Font;
                    uFont.Name = tFont.Name;
                    uFont.Size = tFont.Size;
                    uFont.Color.RGB = tFont.Color.RGB;
                    uFont.Bold = tFont.Bold;
                    
                    user.Number.TextFrame.TextRange.ParagraphFormat.Alignment = template.NumberPrototype.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Alignment;
                } catch { }
                
                float dynamicGap = 5f;
                if (template.Header != null) {
                    dynamicGap = template.Header.Top - (template.NumberPrototype.Top + template.NumberPrototype.Height);
                    if (dynamicGap < 0) dynamicGap = 5f; // Safety fallback
                }
                
                // Adjust finalHeaderTop so Header appears exactly below Number with template's gap
                finalHeaderTop = user.Number.Top + user.Number.Height + dynamicGap;
            }

            // 2. Format Header
            if (user.Header != null)
            {
                user.Header.Left = finalBgLeft + headerOffsetLeft;
                user.Header.Top = finalHeaderTop;
                user.Header.Width = finalBgWidth - headerOffsetLeft - headerOffsetRight;

                if (template.Header != null)
                {
                    try {
                        if ((int)template.Header.AutoShapeType > 0)
                            user.Header.AutoShapeType = template.Header.AutoShapeType;
                    } catch { }
                    try {
                        user.Header.Fill.Visible = template.Header.Fill.Visible;
                        if (template.Header.Fill.Visible == MsoTriState.msoTrue)
                            user.Header.Fill.ForeColor.RGB = template.Header.Fill.ForeColor.RGB;
                        
                        user.Header.Line.Visible = template.Header.Line.Visible;
                        if (template.Header.Line.Visible == MsoTriState.msoTrue)
                            user.Header.Line.ForeColor.RGB = template.Header.Line.ForeColor.RGB;
                    } catch { }
                    if (template.Header.HasTextFrame == MsoTriState.msoTrue && user.Header.HasTextFrame == MsoTriState.msoTrue)
                    {
                        try {
                            var tFont = template.Header.TextFrame.TextRange.Font;
                            var uFont = user.Header.TextFrame.TextRange.Font;
                            uFont.Name = tFont.Name;
                            uFont.Size = targetHeaderSize; // Use the dynamically preserved/calculated size
                            uFont.Color.RGB = tFont.Color.RGB;
                            uFont.Bold = tFont.Bold;
                            
                            user.Header.TextFrame.MarginLeft = template.Header.TextFrame.MarginLeft;
                            user.Header.TextFrame.MarginRight = template.Header.TextFrame.MarginRight;
                            user.Header.TextFrame.MarginTop = template.Header.TextFrame.MarginTop;
                            user.Header.TextFrame.MarginBottom = template.Header.TextFrame.MarginBottom;
                        } catch { }
                    }
                }
            }

            // 3. Distribute Rows
            float contentPaddingTop = 40f;
            
            float bodyOffsetLeft = headerOffsetLeft;
            float bodyOffsetRight = headerOffsetRight;
            float sepOffsetLeft = headerOffsetLeft;
            float sepOffsetRight = headerOffsetRight;
            float iconOffsetLeft = headerOffsetLeft;

            if (template.Background != null)
            {
                if (template.BodyTextPrototype != null)
                {
                    contentPaddingTop = template.BodyTextPrototype.Top - template.Background.Top;
                    bodyOffsetLeft = template.BodyTextPrototype.Left - template.Background.Left;
                    bodyOffsetRight = (template.Background.Left + template.Background.Width) - (template.BodyTextPrototype.Left + template.BodyTextPrototype.Width);
                }
                
                if (template.SeparatorPrototype != null)
                {
                    sepOffsetLeft = template.SeparatorPrototype.Left - template.Background.Left;
                    sepOffsetRight = (template.Background.Left + template.Background.Width) - (template.SeparatorPrototype.Left + template.SeparatorPrototype.Width);
                    
                    // If the separator is floating outside the template area, snap it to body text margins
                    if (sepOffsetLeft < 0 || sepOffsetLeft > template.Background.Width)
                    {
                        sepOffsetLeft = bodyOffsetLeft;
                        sepOffsetRight = bodyOffsetRight;
                    }
                }

                if (template.IconPrototype != null)
                {
                    iconOffsetLeft = template.IconPrototype.Left - template.Background.Left;
                }
                else if (addIcons)
                {
                    iconOffsetLeft = 20f; // Place icon cleanly inside the decor box (20 points padding)
                    float requiredBodyOffset = iconOffsetLeft + 40f + 15f; // 40 width + 15 gap
                    if (bodyOffsetLeft < requiredBodyOffset)
                    {
                        bodyOffsetLeft = requiredBodyOffset; // Shrink text box to accommodate icon
                    }
                }
            }

            // Strict 0.3 cm padding top and bottom (0.3 cm = ~8.5039 points)
            float padding = 8.5039f; 
            float startY = finalBgTop + padding;
            if (user.Header != null)
            {
                // Start 0.3 cm below the bottom edge of the Header
                startY = user.Header.Top + user.Header.Height + padding;
            }
            float endY = finalBgTop + finalBgHeight - padding;
            
            int numRows = user.Rows.Count;
            if (numRows > 0)
            {
                // Regroup alternating texts into Subheader/Body text pairs if this design uses subheaders
                if (template.SubheaderPrototype != null)
                {
                    List<PowerPoint.Shape> allTexts = new List<PowerPoint.Shape>();
                    foreach (var r in user.Rows)
                    {
                        if (r.Icon != null) allTexts.Add(r.Icon); // Mistakenly grouped as icon
                        if (r.Subheader != null) allTexts.Add(r.Subheader);
                        if (r.Text != null) allTexts.Add(r.Text);
                    }
                    
                    allTexts = allTexts.OrderBy(s => s.Top).ToList();

                    List<BodyRow> pairedRows = new List<BodyRow>();
                    for (int i = 0; i < allTexts.Count; i += 2)
                    {
                        var newRow = new BodyRow();
                        newRow.Subheader = allTexts[i];
                        if (i + 1 < allTexts.Count)
                        {
                            newRow.Text = allTexts[i + 1];
                        }
                        pairedRows.Add(newRow);
                    }
                    user.Rows = pairedRows;
                    numRows = user.Rows.Count;
                }

                // Safely copy the Separator once before the loop to avoid COM Clipboard race conditions
                PowerPoint.Shape localSep = null;
                if (template.SeparatorPrototype != null && numRows > 1)
                {
                    for (int retries = 0; retries < 3; retries++)
                    {
                        try
                        {
                            template.SeparatorPrototype.Copy();
                            System.Threading.Thread.Sleep(50); // Give clipboard time to populate
                            var pasted = slide.Shapes.Paste();
                            localSep = pasted[1];
                            localSep.Visible = MsoTriState.msoFalse; // Hide it until duplicated
                            break;
                        }
                        catch { System.Threading.Thread.Sleep(100); }
                    }
                }

                float subheaderOffsetLeft = bodyOffsetLeft;
                if (template.SubheaderPrototype != null && template.Background != null)
                {
                    subheaderOffsetLeft = template.SubheaderPrototype.Left - template.Background.Left;
                }

                // Pass 1: Apply formatting (fonts, width, autosize) so heights calculate correctly
                for (int i = 0; i < numRows; i++)
                {
                    var row = user.Rows[i];

                    if (row.Subheader != null)
                    {
                        row.Subheader.Left = finalBgLeft + subheaderOffsetLeft;
                        row.Subheader.Width = finalBgWidth - subheaderOffsetLeft - bodyOffsetRight;

                        if (template.SubheaderPrototype != null)
                        {
                            try {
                                row.Subheader.TextFrame.WordWrap = MsoTriState.msoTrue;
                                row.Subheader.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                                
                                var tFont = template.SubheaderPrototype.TextFrame.TextRange.Paragraphs(1).Font;
                                var uFont = row.Subheader.TextFrame.TextRange.Font;
                                uFont.Name = tFont.Name;
                                uFont.Size = tFont.Size; 
                                uFont.Color.RGB = tFont.Color.RGB;
                                uFont.Bold = tFont.Bold;
                                
                                row.Subheader.TextFrame.TextRange.ParagraphFormat.Alignment = template.SubheaderPrototype.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Alignment;
                                row.Subheader.TextFrame.MarginLeft = template.SubheaderPrototype.TextFrame.MarginLeft;
                                row.Subheader.TextFrame.MarginRight = template.SubheaderPrototype.TextFrame.MarginRight;
                                row.Subheader.TextFrame.MarginTop = template.SubheaderPrototype.TextFrame.MarginTop;
                                row.Subheader.TextFrame.MarginBottom = template.SubheaderPrototype.TextFrame.MarginBottom;
                            } catch { }
                        }
                    }

                    if (row.Text != null)
                    {
                        row.Text.Left = finalBgLeft + bodyOffsetLeft;
                        row.Text.Width = finalBgWidth - bodyOffsetLeft - bodyOffsetRight;
                        
                        if (template.BodyTextPrototype != null)
                        {
                            try {
                                // Explicitly enforce "Resize shape to fit text" and word wrap
                                row.Text.TextFrame.WordWrap = MsoTriState.msoTrue;
                                row.Text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;

                                var tFont = template.BodyTextPrototype.TextFrame.TextRange.Paragraphs(1).Font;
                                var uFont = row.Text.TextFrame.TextRange.Font;
                                uFont.Name = tFont.Name;
                                uFont.Size = targetBodySize; // Locked to 11
                                uFont.Color.RGB = tFont.Color.RGB;
                                uFont.Bold = tFont.Bold;
                                
                                row.Text.TextFrame.TextRange.ParagraphFormat.Alignment = template.BodyTextPrototype.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Alignment;
                                row.Text.TextFrame.MarginLeft = template.BodyTextPrototype.TextFrame.MarginLeft;
                                row.Text.TextFrame.MarginRight = template.BodyTextPrototype.TextFrame.MarginRight;
                                row.Text.TextFrame.MarginTop = template.BodyTextPrototype.TextFrame.MarginTop;
                                row.Text.TextFrame.MarginBottom = template.BodyTextPrototype.TextFrame.MarginBottom;
                            } catch { }
                        }
                    } // End of if (row.Text != null)
                } // End of for loop

                // Now calculate the uniform icon size based on the heights of the AutoSized text boxes
                float uniformIconSize = 40f;
                if (addIcons && template.IconPrototype == null)
                {
                    var textHeights = user.Rows.Where(r => r.Text != null).Select(r => r.Text.Height).ToList();
                    if (textHeights.Count > 0)
                    {
                        uniformIconSize = Math.Max(15f, textHeights.Min()); // Minimum text height, capped to at least 15
                    }
                }

                // Pass 1B: Apply Option B automatic icons and set icon sizes
                for (int i = 0; i < numRows; i++)
                {
                    var row = user.Rows[i];

                    // Apply Option B automatic icons if requested
                    if (addIcons && row.Icon == null && row.Text != null)
                    {
                        try {
                            string textToSearch = row.Subheader != null ? row.Subheader.TextFrame.TextRange.Text : row.Text.TextFrame.TextRange.Text;
                            string svgPath = AutoIconService.SaveSvgToTempFile(AutoIconService.GetBestIconSvg(textToSearch));
                            if (svgPath != null)
                            {
                                row.Icon = slide.Shapes.AddPicture(svgPath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, 50, 50);
                                File.Delete(svgPath);
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show("AutoIconService returned null SVG for text: " + textToSearch);
                            }
                        } catch (Exception ex) {
                            System.Windows.Forms.MessageBox.Show("Error generating icon: " + ex.Message);
                        }
                    }

                    if (row.Icon != null)
                    {
                        row.Icon.Left = finalBgLeft + iconOffsetLeft;
                        if (template.IconPrototype != null)
                        {
                            try {
                                row.Icon.Width = template.IconPrototype.Width;
                                row.Icon.Height = template.IconPrototype.Height;

                                row.Icon.Fill.Visible = template.IconPrototype.Fill.Visible;
                                if (template.IconPrototype.Fill.Visible == MsoTriState.msoTrue)
                                    row.Icon.Fill.ForeColor.RGB = template.IconPrototype.Fill.ForeColor.RGB;

                                row.Icon.Line.Visible = template.IconPrototype.Line.Visible;
                                if (template.IconPrototype.Line.Visible == MsoTriState.msoTrue)
                                    row.Icon.Line.ForeColor.RGB = template.IconPrototype.Line.ForeColor.RGB;
                            } catch { }
                        }
                        else
                        {
                            try {
                                row.Icon.Width = uniformIconSize;
                                row.Icon.Height = uniformIconSize;
                                // DO NOT override Fill or Line properties. Leave original SVG look as requested.
                            } catch { }
                        }
                    }
                }

                // Pass 2: Calculate dynamic heights and enforce perfect sequencing (Distribute Vertically)
                float availableSpace = endY - startY;
                float sumTextHeights = 0;
                float internalPairGap = 2f; // Slight gap between Subheader and Body text inside a pair

                for (int i = 0; i < numRows; i++)
                {
                    float rh = 0;
                    if (user.Rows[i].Subheader != null) rh += user.Rows[i].Subheader.Height;
                    if (user.Rows[i].Text != null) {
                        if (user.Rows[i].Subheader != null) rh += internalPairGap;
                        rh += user.Rows[i].Text.Height;
                    }
                    if (rh == 0) rh = 20f;
                    sumTextHeights += rh;
                }
                
                float sepHeight = localSep != null ? localSep.Height : 0;
                float sumLineHeights = (numRows > 1) ? (numRows - 1) * sepHeight : 0;
                
                float totalEmptySpace = availableSpace - sumTextHeights - sumLineHeights;
                float numGaps = (numRows > 1) ? (numRows - 1) * 2 : 1;
                float gapHeight = (numRows > 1) ? totalEmptySpace / numGaps : 0;
                
                // Prevent negative gap heights (overlapping) if text is taller than original bounds
                if (gapHeight < 5f && numRows > 1) gapHeight = 5f;
                
                float currentY = startY;
                
                for (int i = 0; i < numRows; i++)
                {
                    var row = user.Rows[i];
                    
                    float rh = 0;
                    if (row.Subheader != null) rh += row.Subheader.Height;
                    if (row.Text != null)
                    {
                        if (row.Subheader != null) rh += internalPairGap;
                        rh += row.Text.Height;
                    }
                    if (rh == 0) rh = 20f;

                    if (row.Subheader != null)
                    {
                        row.Subheader.Top = currentY;
                        currentY += row.Subheader.Height;
                    }

                    if (row.Text != null)
                    {
                        if (row.Subheader != null) currentY += internalPairGap;
                        row.Text.Top = currentY;
                        currentY += row.Text.Height;
                    }

                    if (row.Subheader == null && row.Text == null)
                    {
                        currentY += 20f;
                    }

                    if (row.Icon != null) row.Icon.Top = (currentY - rh) + (rh - row.Icon.Height) / 2;
                    
                    if (i < numRows - 1 && localSep != null)
                    {
                        currentY += gapHeight; // Distribute exact empty space before the line
                        
                        try {
                            var sep = localSep.Duplicate()[1];
                            sep.Visible = MsoTriState.msoTrue;
                            sep.Left = finalBgLeft + sepOffsetLeft;
                            sep.Width = finalBgWidth - sepOffsetLeft - sepOffsetRight;
                            sep.Top = currentY;
                        } catch { }
                        
                        currentY += sepHeight; // Move past the line
                        currentY += gapHeight; // Distribute exact empty space after the line
                    }
                }
                
                // Track the final bottom position to resize the background if needed
                startY = currentY + padding; 

                // Cleanup local template
                if (localSep != null)
                {
                    try { localSep.Delete(); } catch { }
                }
            }

            // Dynamic Background Expansion
            if (bgContainer != null)
            {
                if (startY > finalBgTop + finalBgHeight)
                {
                    float requiredHeight = startY - finalBgTop;
                    bgContainer.Height = requiredHeight;
                    
                    // Enforce the strict aspect ratio from the library template when expanding
                    if (template.Background != null && template.Background.Height > 0)
                    {
                        float aspect = template.Background.Width / template.Background.Height;
                        bgContainer.Width = requiredHeight * aspect;
                    }
                }
            }
        }
    }
}
