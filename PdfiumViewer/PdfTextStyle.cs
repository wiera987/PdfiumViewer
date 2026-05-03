using System;
using System.Drawing;
using static PdfiumViewer.NativeMethods;

namespace PdfiumViewer
{
    [Flags]
    public enum PdfTextStyleFlags : int
    {
        None = 0,
        FillColor       = 1 << 0,
        StrokeColor     = 1 << 1,
        Underline       = 1 << 2,
        Strikethrough   = 1 << 3,
        Highlight       = 1 << 4,
        Squiggly        = 1 << 5,
        AnnotationColor = 1 << 6,

        All = FillColor | StrokeColor | Underline | Strikethrough |
              Highlight | Squiggly | AnnotationColor
    }


    /// <summary>
    /// Represents the style information of style1 specific text character in style1 PDF document.
    /// </summary>
    public class PdfTextStyle
    {
        public int PageIndex { get; set; }
        public int CharIndex { get; set; }

        //public string FontName { get; set; }
        //public float FontSize { get; set; }
        public Color FillColor { get; set; }
        public Color StrokeColor { get; set; }
        public bool IsUnderlined { get; set; }
        public bool IsStrikethrough { get; set; }
        public bool IsHighlighted { get; set; }
        public bool IsSquiggly { get; set; }
        public Color UnderlineColor { get; set; }
        public Color StrikeoutColor { get; set; }
        public Color HighlightColor { get; set; }
        public Color SquigglyColor { get; set; }

        public PdfTextStyle(
            int pageIndex,
            int charIndex,
            //string fontName,
            //float fontSize,
            Color fillColor,
            Color strokeColor,
            bool isUnderlined = false,
            bool isStrikethrough = false,
            bool isHighlighted = false,
            bool isSquiggly = false,
            Color underlineColor = default(Color),
            Color strikeoutColor = default(Color),
            Color highlightColor = default(Color),
            Color squigglyColor = default(Color))
        {
            PageIndex = pageIndex;
            CharIndex = charIndex;
            //FontName = fontName;
            //FontSize = fontSize;
            FillColor = fillColor;
            StrokeColor = strokeColor;
            IsUnderlined = isUnderlined;
            IsStrikethrough = isStrikethrough;
            IsHighlighted = isHighlighted;
            IsSquiggly = isSquiggly;
            UnderlineColor = underlineColor;
            StrikeoutColor = strikeoutColor;
            HighlightColor = highlightColor;
            SquigglyColor = squigglyColor;
        }

        internal void SetColor(FPDF_ANNOTATION_SUBTYPE subtype,
                                Color color)
        {
            switch (subtype)
            {
                case FPDF_ANNOTATION_SUBTYPE.UNDERLINE:
                    UnderlineColor = color;
                    break;
                case FPDF_ANNOTATION_SUBTYPE.STRIKEOUT:
                    StrikeoutColor = color;
                    break;
                case FPDF_ANNOTATION_SUBTYPE.HIGHLIGHT:
                    HighlightColor = color;
                    break;
                case FPDF_ANNOTATION_SUBTYPE.SQUIGGLY:
                    SquigglyColor = color;
                    break;
            }
        }

        /// <summary>
        /// Determines whether two PdfTextStyle instances match according to the specified PdfTextStyleFlags.
        /// </summary>
        /// <param name="style1">The first PdfTextStyle to compare.</param>
        /// <param name="style2">The second PdfTextStyle to compare.</param>
        /// <param name="flag">Flags indicating which style attributes to compare.</param>
        /// <returns>True if the selected style attributes match; otherwise, false.</returns>
        public static bool IsMatched(PdfTextStyle style1, PdfTextStyle style2, PdfTextStyleFlags flag)
        {
            if (flag == PdfTextStyleFlags.None)
            {
                return true;    // No comparison items selected.
            }

            // Comparing selected items.
            if (((flag & PdfTextStyleFlags.FillColor) != 0) &&(style1.FillColor != style2.FillColor))
            {
                return false;
            }
            if (((flag & PdfTextStyleFlags.StrokeColor) != 0) && (style1.StrokeColor != style2.StrokeColor))
            {
                return false;
            }
            if (((flag & PdfTextStyleFlags.Underline) != 0) && (style1.IsUnderlined != style2.IsUnderlined))
            {
                return false;
            }
            if (((flag & PdfTextStyleFlags.Strikethrough) != 0) && (style1.IsStrikethrough != style2.IsStrikethrough))
            {
                return false;
            }
            if (((flag & PdfTextStyleFlags.Highlight) != 0) && (style1.IsHighlighted != style2.IsHighlighted))
            {
                return false;
            }
            if (((flag & PdfTextStyleFlags.Squiggly) != 0) && (style1.IsSquiggly != style2.IsSquiggly))
            {
                return false;
            }
            if (((flag & PdfTextStyleFlags.AnnotationColor) != 0)
                && ((style1.UnderlineColor != style2.UnderlineColor)
                    || (style1.StrikeoutColor != style2.StrikeoutColor)
                    || (style1.HighlightColor != style2.HighlightColor)
                    || (style1.SquigglyColor != style2.SquigglyColor)))
            {
                return false;
            }

            return true;
        }
    }
}