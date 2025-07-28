using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    public static class HarveyBall
    {
        /// <summary>
        /// Inserts a Harvey Ball symbol into the current PowerPoint selection.
        /// </summary>
        /// <param name="value">Value from 0 to 4 indicating the Harvey Ball fill level.</param>
        public static void InsertHarveyBall(int value)
        {
            // Unicode Harvey Balls: 0-4
            // 0: ⭘ (U+2B58), 1: ◔ (U+25D4), 2: ◑ (U+25D1), 3: ◕ (U+25D5), 4: ● (U+25CF)
            string[] harveyBalls = { "\u2B58", "\u25D4", "\u25D1", "\u25D5", "\u25CF" };
            if (value < 0 || value > 4) return;
            string harveyBall = harveyBalls[value];

            var app = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
            var selection = app.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                // Insert at cursor in text box
                selection.TextRange.Text = harveyBall;
                selection.TextRange.Font.Name = "Segoe UI Symbol";
            }
            else if (selection.Type == PpSelectionType.ppSelectionSlides)
            {
                // Insert as new shape in the center of the selected slide
                var slide = selection.SlideRange[1];
                float left = app.ActivePresentation.PageSetup.SlideWidth / 2 - 20;
                float top = app.ActivePresentation.PageSetup.SlideHeight / 2 - 20;
                var textbox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, 40, 40);
                textbox.TextFrame.TextRange.Text = harveyBall;
                textbox.TextFrame.TextRange.Font.Name = "Segoe UI Symbol";
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                // Replace text in selected shape(s)
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        shape.TextFrame.TextRange.Text = harveyBall;
                        shape.TextFrame.TextRange.Font.Name = "Segoe UI Symbol";
                    }
                }
            }
        }
    }

   /* public static class ElementHelper
    {
        public static void InsertElement(string element)
        {
            var app = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
            var selection = app.ActiveWindow.Selection;

            // Check if selection is in a table cell or multiple table cells
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape tableShape = null;
                int row = -1, col = -1;
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        tableShape = shape;
                        // Try to get selected cell (if possible)
                        if (shape.Table.SelectedCell != null)
                        {
                            row = shape.Table.SelectedCell.Row;
                            col = shape.Table.SelectedCell.Column;
                        }
                        break;
                    }
                }
                if (tableShape != null)
                {
                    InsertElementIntoTable(element, tableShape, row, col);
                    return;
                }
            }
            // Check if selection is in a text box or shape with text
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                InsertElementIntoText(element);
                return;
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        InsertElementIntoText(element);
                        return;
                    }
                }
            }
            // Otherwise, insert as new shape/textbox in the center of the selected slide
            InsertElementIntoSlide(element);
        }

        public static void InsertElementIntoTable(string element)
        {

        }

        public static void InsertElementIntoText(string element)
        {

        }

        public static void InsertElementIntoSlide(string element)
        {

        }
    }*/
}