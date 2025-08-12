using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    public static class ElementHelper
    {
        public static void InsertElement(string element, int? colorRgb = null)
        {
            var app = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
            var selection = app.ActiveWindow.Selection;

            // Check if selection is in a table cell or multiple table cells
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape tableShape = null;
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        tableShape = shape;
                        break;
                    }
                }
                if (tableShape != null)
                {
                    InsertElementIntoTable(element, tableShape, selection, colorRgb);
                    return;
                }
            }

            // Check if selection is in a text box or shape with text
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                InsertElementIntoText(element, selection, colorRgb);
                return;
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                // Check if any selected shapes have text frames
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        InsertElementIntoShapeText(element, shape, colorRgb);
                        return;
                    }
                }
            }

            // Otherwise, insert as new shape/textbox in the center of the selected slide
            InsertElementIntoSlide(element, app, colorRgb);
        }


        #region Private Helper Methods

        private static void InsertElementIntoTable(string element, Shape tableShape, Selection selection, int? colorRgb = null)
        {
            // If user has cursor in table cell text, insert there
            if (selection.Type == PpSelectionType.ppSelectionText && selection.TextRange != null)
            {
                // Record original formatting before making changes
                var originalFontName = selection.TextRange.Font.Name;
                var originalFontColor = selection.TextRange.Font.Color.RGB;
                
                // Insert symbol + zero-width space for cursor positioning
                var textRange = selection.TextRange;
                textRange.Text = element + "\u200B"; // Zero-width space
                
                // Calculate the length of the element (handles surrogate pairs)
                int elementLength = element.Length;
                
                // Apply symbol formatting to the element part only
                var symbolRange = textRange.Characters(1, elementLength);
                symbolRange.Font.Name = "Segoe UI Symbol";
                if (colorRgb.HasValue)
                {
                    int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                    symbolRange.Font.Color.RGB = bgrColor;
                }
                
                // Apply original formatting to the zero-width space for future typing
                var spacerRange = textRange.Characters(elementLength + 1, 1);
                spacerRange.Font.Name = originalFontName;
                spacerRange.Font.Color.RGB = originalFontColor;
                
                // Position cursor after the symbol but in the formatted spacer
                spacerRange.Select();
            }
            else
            {
                // Otherwise, insert into first cell as fallback (no cursor positioning needed)
                var cell = tableShape.Table.Cell(1, 1);
                var textRange = cell.Shape.TextFrame.TextRange;
                
                // For fallback, just replace the entire cell content
                textRange.Text = element;
                textRange.Font.Name = "Segoe UI Symbol";
                if (colorRgb.HasValue)
                {
                    int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                    textRange.Font.Color.RGB = bgrColor;
                }
            }
        }

        private static void InsertElementIntoText(string element, Selection selection, int? colorRgb = null)
        {
            // Record original formatting before making changes
            var originalFontName = selection.TextRange.Font.Name;
            var originalFontColor = selection.TextRange.Font.Color.RGB;
            
            // Insert symbol + zero-width space for cursor positioning
            var textRange = selection.TextRange;
            textRange.Text = element + "\u200B"; // Zero-width space
            
            // Calculate the length of the element (handles surrogate pairs)
            int elementLength = element.Length;
            
            // Apply symbol formatting to the element part only
            var symbolRange = textRange.Characters(1, elementLength);
            symbolRange.Font.Name = "Segoe UI Symbol";
            if (colorRgb.HasValue)
            {
                int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                symbolRange.Font.Color.RGB = bgrColor;
            }
            
            // Apply original formatting to the zero-width space for future typing
            var spacerRange = textRange.Characters(elementLength + 1, 1);
            spacerRange.Font.Name = originalFontName;
            spacerRange.Font.Color.RGB = originalFontColor;
            
            // Position cursor after the symbol but in the formatted spacer
            spacerRange.Select();
        }

        private static void InsertElementIntoShapeText(string element, Shape shape, int? colorRgb = null)
        {
            // Record original formatting before making changes
            var textRange = shape.TextFrame.TextRange;
            var originalFontName = textRange.Font.Name;
            var originalFontColor = textRange.Font.Color.RGB;
            
            // Insert symbol + zero-width space for cursor positioning
            textRange.Text = element + "\u200B"; // Zero-width space
            
            // Calculate the length of the element (handles surrogate pairs)
            int elementLength = element.Length;
            
            // Apply symbol formatting to the element part only
            var symbolRange = textRange.Characters(1, elementLength);
            symbolRange.Font.Name = "Segoe UI Symbol";
            if (colorRgb.HasValue)
            {
                int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                symbolRange.Font.Color.RGB = bgrColor;
            }
            
            // Apply original formatting to the zero-width space for future typing
            var spacerRange = textRange.Characters(elementLength + 1, 1);
            spacerRange.Font.Name = originalFontName;
            spacerRange.Font.Color.RGB = originalFontColor;
        }

        private static void InsertElementIntoSlide(string element, Application app, int? colorRgb = null)
        {
            // Insert as new shape in the center of the active slide
            var slide = app.ActiveWindow.View.Slide;
            float left = app.ActivePresentation.PageSetup.SlideWidth / 2 - 20;
            float top = app.ActivePresentation.PageSetup.SlideHeight / 2 - 20;
            var shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, 40, 40);
            
            // Configure text properties
            shape.TextFrame.TextRange.Text = element;
            shape.TextFrame.TextRange.Font.Name = "Segoe UI Symbol";
            shape.TextFrame.TextRange.Font.Size = 16;
            
            // Apply color if specified
            if (colorRgb.HasValue)
            {
                // Convert RGB to BGR format for PowerPoint
                int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                shape.TextFrame.TextRange.Font.Color.RGB = bgrColor;
            }
            
            // Center and middle align the text
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
        }

        // Helper method to convert RGB to BGR format (PowerPoint uses BGR internally)
        private static int ConvertRgbToBgr(int rgbColor)
        {
            // Extract RGB components
            int r = (rgbColor >> 16) & 0xFF;
            int g = (rgbColor >> 8) & 0xFF;
            int b = rgbColor & 0xFF;
            
            // Recombine as BGR
            return (b << 16) | (g << 8) | r;
        }

        #endregion
    }
}