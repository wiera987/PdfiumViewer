using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using static PdfiumViewer.NativeMethods;
using static System.Net.WebRequestMethods;

namespace PdfiumViewer
{
    internal class PdfFile : IDisposable
    {
        private static readonly Encoding FPDFEncoding = new UnicodeEncoding(false, false, false);

        private IntPtr _document;
        private IntPtr _form;
        private bool _disposed;
        private NativeMethods.FPDF_FORMFILLINFO _formCallbacks;
        private GCHandle _formCallbacksHandle;
        private readonly int _id;
        private Stream _stream;
        private List<PageData> _pageData = null;

        public PdfFile(Stream stream, string password)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            PdfLibrary.EnsureLoaded();

            _stream = stream;
            _id = StreamManager.Register(stream);

            var document = NativeMethods.FPDF_LoadCustomDocument(stream, password, _id);
            if (document == IntPtr.Zero)
                throw new PdfException((PdfError)NativeMethods.FPDF_GetLastError());

            LoadDocument(document);
        }

        public PdfBookmarkCollection Bookmarks { get; private set; }

        public bool RenderPDFPageToDC(int pageNumber, IntPtr dc, int dpiX, int dpiY, int boundsOriginX, int boundsOriginY, int boundsWidth, int boundsHeight, NativeMethods.FPDF flags)
        {
            if (_disposed)
                throw new ObjectDisposedException(GetType().Name);

            var pageData = GetPageData(_document, _form, pageNumber);
            {
                NativeMethods.FPDF_RenderPage(dc, pageData.Page, boundsOriginX, boundsOriginY, boundsWidth, boundsHeight, 0, flags);
            }

            return true;
        }

        public bool RenderPDFPageToBitmap(int pageNumber, IntPtr bitmapHandle, int dpiX, int dpiY, int boundsOriginX, int boundsOriginY, int boundsWidth, int boundsHeight, int rotate, NativeMethods.FPDF flags, bool renderFormFill)
        {
            if (_disposed)
                throw new ObjectDisposedException(GetType().Name);

            var pageData = GetPageData(_document, _form, pageNumber);
            {
                NativeMethods.FPDF_RenderPageBitmap(bitmapHandle, pageData.Page, boundsOriginX, boundsOriginY, boundsWidth, boundsHeight, rotate, flags);

                if (renderFormFill)
                {
                    flags &= ~NativeMethods.FPDF.ANNOT;
                    NativeMethods.FPDF_FFLDraw(_form, bitmapHandle, pageData.Page, boundsOriginX, boundsOriginY, boundsWidth, boundsHeight, rotate, flags);
                }
            }

            return true;
        }

        public PdfPageLinks GetPageLinks(int pageNumber, Size pageSize)
        {
            if (_disposed)
                throw new ObjectDisposedException(GetType().Name);

            var links = new List<PdfPageLink>();

            var pageData = GetPageData(_document, _form, pageNumber);
            {
                int link = 0;
                IntPtr annotation;

                while (NativeMethods.FPDFLink_Enumerate(pageData.Page, ref link, out annotation))
                {
                    var destination = NativeMethods.FPDFLink_GetDest(_document, annotation);
                    int? target = null;
                    string uri = null;

                    if (destination != IntPtr.Zero)
                        target = (int)NativeMethods.FPDFDest_GetDestPageIndex(_document, destination);

                    var action = NativeMethods.FPDFLink_GetAction(annotation);
                    if (action != IntPtr.Zero)
                    {
                        const uint length = 1024;
                        var sb = new StringBuilder(1024);
                        NativeMethods.FPDFAction_GetURIPath(_document, action, sb, length);

                        uri = sb.ToString();
                    }

                    var rect = new NativeMethods.FS_RECTF();

                    if (NativeMethods.FPDFLink_GetAnnotRect(annotation, rect) && (target.HasValue || uri != null))
                    {
                        links.Add(new PdfPageLink(
                            new RectangleF(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top),
                            target,
                            uri
                        ));
                    }
                }
            }

            return new PdfPageLinks(links);
        }

        public List<SizeF> GetPDFDocInfo()
        {
            if (_disposed)
                throw new ObjectDisposedException(GetType().Name);

            int pageCount = NativeMethods.FPDF_GetPageCount(_document);
            var result = new List<SizeF>(pageCount);
            _pageData = new List<PageData>(pageCount);

            for (int i = 0; i < pageCount; i++)
            {
                result.Add(GetPDFDocInfo(i));
                _pageData.Add(null);
            }

            return result;
        }

        public SizeF GetPDFDocInfo(int pageNumber)
        {
            double height;
            double width;
            NativeMethods.FPDF_GetPageSizeByIndex(_document, pageNumber, out width, out height);

            return new SizeF((float)width, (float)height);
        }

        public void Save(Stream stream)
        {
            NativeMethods.FPDF_SaveAsCopy(_document, stream, NativeMethods.FPDF_SAVE_FLAGS.FPDF_NO_INCREMENTAL);
        }

        protected void LoadDocument(IntPtr document)
        {
            _document = document;

            NativeMethods.FPDF_GetDocPermissions(_document);

            _formCallbacks = new NativeMethods.FPDF_FORMFILLINFO();
            _formCallbacksHandle = GCHandle.Alloc(_formCallbacks, GCHandleType.Pinned);

            // Depending on whether XFA support is built into the PDFium library, the version
            // needs to be 1 or 2. We don't really care, so we just try one or the other.

            for (int i = 1; i <= 2; i++)
            {
                _formCallbacks.version = i;

                _form = NativeMethods.FPDFDOC_InitFormFillEnvironment(_document, _formCallbacks);
                if (_form != IntPtr.Zero)
                    break;
            }

            NativeMethods.FPDF_SetFormFieldHighlightColor(_form, 0, 0xFFE4DD);
            NativeMethods.FPDF_SetFormFieldHighlightAlpha(_form, 100);

            NativeMethods.FORM_DoDocumentJSAction(_form);
            NativeMethods.FORM_DoDocumentOpenAction(_form);

            Bookmarks = new PdfBookmarkCollection();

            LoadBookmarks(Bookmarks, NativeMethods.FPDF_BookmarkGetFirstChild(document, IntPtr.Zero));
        }

        private void LoadBookmarks(PdfBookmarkCollection bookmarks, IntPtr bookmark)
        {
            if (bookmark == IntPtr.Zero)
                return;

            bookmarks.Add(LoadBookmark(bookmark));
            while ((bookmark = NativeMethods.FPDF_BookmarkGetNextSibling(_document, bookmark)) != IntPtr.Zero)
                bookmarks.Add(LoadBookmark(bookmark));
        }

        private PdfBookmark LoadBookmark(IntPtr bookmark)
        {
            var result = new PdfBookmark
            {
                Title = GetBookmarkTitle(bookmark),
                PageIndex = (int)GetBookmarkPageIndex(bookmark)
            };

            //Action = NativeMethods.FPDF_BookmarkGetAction(_bookmark);
            //if (Action != IntPtr.Zero)
            //    ActionType = NativeMethods.FPDF_ActionGetType(Action);

            var child = NativeMethods.FPDF_BookmarkGetFirstChild(_document, bookmark);
            if (child != IntPtr.Zero)
                LoadBookmarks(result.Children, child);

            return result;
        }

        private string GetBookmarkTitle(IntPtr bookmark)
        {
            uint length = NativeMethods.FPDF_BookmarkGetTitle(bookmark, null, 0);
            byte[] buffer = new byte[length];
            NativeMethods.FPDF_BookmarkGetTitle(bookmark, buffer, length);

            string result = Encoding.Unicode.GetString(buffer);
            if (result.Length > 0 && result[result.Length - 1] == 0)
                result = result.Substring(0, result.Length - 1);

            return result;
        }

        private uint GetBookmarkPageIndex(IntPtr bookmark)
        {
            IntPtr dest = NativeMethods.FPDF_BookmarkGetDest(_document, bookmark);
            if (dest != IntPtr.Zero)
                return NativeMethods.FPDFDest_GetDestPageIndex(_document, dest);

            return 0;
        }

        public PdfMatches Search(string text, bool matchCase, bool wholeWord, int startPage, int endPage)
        {
            var matches = new List<PdfMatch>();

            if (String.IsNullOrEmpty(text))
                return new PdfMatches(startPage, endPage, matches);

            for (int page = startPage; page <= endPage; page++)
            {
                var pageData = GetPageData(_document, _form, page);
                {
                    NativeMethods.FPDF_SEARCH_FLAGS flags = 0;
                    if (matchCase)
                        flags |= NativeMethods.FPDF_SEARCH_FLAGS.FPDF_MATCHCASE;
                    if (wholeWord)
                        flags |= NativeMethods.FPDF_SEARCH_FLAGS.FPDF_MATCHWHOLEWORD;

                    var handle = NativeMethods.FPDFText_FindStart(pageData.TextPage, FPDFEncoding.GetBytes(text), flags, 0);

                    try
                    {
                        while (NativeMethods.FPDFText_FindNext(handle))
                        {
                            int index = NativeMethods.FPDFText_GetSchResultIndex(handle);

                            int matchLength = NativeMethods.FPDFText_GetSchCount(handle);

                            var result = new byte[(matchLength + 1) * 2];
                            NativeMethods.FPDFText_GetText(pageData.TextPage, index, matchLength, result);
                            string match = FPDFEncoding.GetString(result, 0, matchLength * 2);

                            matches.Add(new PdfMatch(
                                match,
                                new PdfTextSpan(page, index, matchLength),
                                page
                            ));
                        }
                    }
                    finally
                    {
                        NativeMethods.FPDFText_FindClose(handle);
                    }
                }
            }

            return new PdfMatches(startPage, endPage, matches);
        }

        public IList<PdfRectangle> GetTextBounds(PdfTextSpan textSpan)
        {
            var result = new List<PdfRectangle>();

            var pageData = GetPageData(_document, _form, textSpan.Page);

            {
                int rect_count = NativeMethods.FPDFText_CountRects(pageData.TextPage, textSpan.Offset, textSpan.Length);

                for (int i = 0; i < rect_count; i++)
                {
                    NativeMethods.FPDFText_GetRect(pageData.TextPage, i, out var left, out var top, out var right, out var bottom);

                    RectangleF bounds = new RectangleF((float)left, (float)top, (float)(right - left), (float)(bottom - top));

                    result.Add(new PdfRectangle(textSpan.Page, bounds));
                }

                return result;
            }
        }

        public Point PointFromPdf(int page, PointF point)
        {
            var pageData = GetPageData(_document, _form, page);
            {
                NativeMethods.FPDF_PageToDevice(
                    pageData.Page,
                    0,
                    0,
                    (int)pageData.Width,
                    (int)pageData.Height,
                    0,
                    point.X,
                    point.Y,
                    out var deviceX,
                    out var deviceY
                );

                return new Point(deviceX, deviceY);
            }
        }

        public Rectangle RectangleFromPdf(int page, RectangleF rect)
        {
            var pageData = GetPageData(_document, _form, page);
            {
                NativeMethods.FPDF_PageToDevice(
                    pageData.Page,
                    0,
                    0,
                    (int)pageData.Width,
                    (int)pageData.Height,
                    0,
                    rect.Left,
                    rect.Top,
                    out var deviceX1,
                    out var deviceY1
                );

                NativeMethods.FPDF_PageToDevice(
                    pageData.Page,
                    0,
                    0,
                    (int)pageData.Width,
                    (int)pageData.Height,
                    0,
                    rect.Right,
                    rect.Bottom,
                    out var deviceX2,
                    out var deviceY2
                );

                return new Rectangle(
                    deviceX1,
                    deviceY1,
                    deviceX2 - deviceX1,
                    deviceY2 - deviceY1
                );
            }
        }

        public PointF PointToPdf(int page, Point point)
        {
            var pageData = GetPageData(_document, _form, page);
            {
                NativeMethods.FPDF_DeviceToPage(
                    pageData.Page,
                    0,
                    0,
                    (int)pageData.Width,
                    (int)pageData.Height,
                    0,
                    point.X,
                    point.Y,
                    out var deviceX,
                    out var deviceY
                );

                return new PointF((float)deviceX, (float)deviceY);
            }
        }

        public RectangleF RectangleToPdf(int page, Rectangle rect)
        {
        	var pageData = GetPageData(_document, _form, page);
            {
                NativeMethods.FPDF_DeviceToPage(
                    pageData.Page,
                    0,
                    0,
                    (int)pageData.Width,
                    (int)pageData.Height,
                    0,
                    rect.Left,
                    rect.Top,
                    out var deviceX1,
                    out var deviceY1
                );

                NativeMethods.FPDF_DeviceToPage(
                    pageData.Page,
                    0,
                    0,
                    (int)pageData.Width,
                    (int)pageData.Height,
                    0,
                    rect.Right,
                    rect.Bottom,
                    out var deviceX2,
                    out var deviceY2
                );

                return new RectangleF(
                    (float)deviceX1,
                    (float)deviceY1,
                    (float)(deviceX2 - deviceX1),
                    (float)(deviceY2 - deviceY1)
                );
            }
        }

        private IList<PdfRectangle> GetTextBounds(IntPtr textPage, int pageNumber, int index, int matchLength)
        {
            var result = new List<PdfRectangle>();
            RectangleF? lastBounds = null;

            for (int i = 0; i < matchLength; i++)
            {
                var bounds = GetBounds(textPage, index + i);

                if (bounds.Width == 0 || bounds.Height == 0)
                    continue;

                if (
                    lastBounds.HasValue &&
                    AreClose(lastBounds.Value.Right, bounds.Left) &&
                    AreClose(lastBounds.Value.Top, bounds.Top) &&
                    AreClose(lastBounds.Value.Bottom, bounds.Bottom)
                )
                {
                    float top = Math.Max(lastBounds.Value.Top, bounds.Top);
                    float bottom = Math.Min(lastBounds.Value.Bottom, bounds.Bottom);

                    lastBounds = new RectangleF(
                        lastBounds.Value.Left,
                        top,
                        bounds.Right - lastBounds.Value.Left,
                        bottom - top
                    );

                    result[result.Count - 1] = new PdfRectangle(pageNumber, lastBounds.Value);
                }
                else
                {
                    lastBounds = bounds;
                    result.Add(new PdfRectangle(pageNumber, bounds));
                }
            }

            return result;
        }

        private bool AreClose(float p1, float p2)
        {
            return Math.Abs(p1 - p2) < 4f;
        }

        private RectangleF GetBounds(IntPtr textPage, int index)
        {
            NativeMethods.FPDFText_GetCharBox(
                textPage,
                index,
                out var left,
                out var right,
                out var bottom,
                out var top
            );

            return new RectangleF(
                (float)left,
                (float)top,
                (float)(right - left),
                (float)(bottom - top)
            );
        }

        public string GetPdfText(int pageNumber)
        {
            var pageData = GetPageData(_document, _form, pageNumber);
            {
                int length = NativeMethods.FPDFText_CountChars(pageData.TextPage);
                return GetPdfText(pageData, new PdfTextSpan(pageNumber, 0, length));
            }
        }

        public string GetPdfText(PdfTextSpan textSpan)
        {
            var pageData = GetPageData(_document, _form, textSpan.Page);
            {
                return GetPdfText(pageData, textSpan);
            }
        }

        private string GetPdfText(PageData pageData, PdfTextSpan textSpan)
        {
            var result = new byte[(textSpan.Length + 1) * 2];
            NativeMethods.FPDFText_GetText(pageData.TextPage, textSpan.Offset, textSpan.Length, result);
            return FPDFEncoding.GetString(result, 0, textSpan.Length * 2);
        }

        public int GetCharIndexAtPos(PdfPoint location, double xTolerance, double yTolerance)
        {
            var pageData = GetPageData(_document, _form, location.Page);
            {
                return NativeMethods.FPDFText_GetCharIndexAtPos(
                    pageData.TextPage,
                    location.Location.X,
                    location.Location.Y,
                    xTolerance,
                    yTolerance
                );
            }
        }

        public bool GetWordAtPosition(PdfPoint location, double xTolerance, double yTolerance, out PdfTextSpan span)
        {
            var index = GetCharIndexAtPos(location, xTolerance, yTolerance);
            if (index < 0)
            {
                span = default(PdfTextSpan);
                return false;
            }

            var baseCharacter = GetCharacter(location.Page, index);
            if (IsWordSeparator(baseCharacter))
            {
                span = default(PdfTextSpan);
                return false;
            }

            int start = index, end = index;

            for (int i = index - 1; i >= 0; i--)
            {
                var c = GetCharacter(location.Page, i);
                if (IsWordSeparator(c))
                    break;
                start = i;
            }

            var count = CountChars(location.Page);
            for (int i = index + 1; i < count; i++)
            {
                var c = GetCharacter(location.Page, i);
                if (IsWordSeparator(c))
                    break;
                end = i;
            }

            span = new PdfTextSpan(location.Page, start, end - start);
            return true;

            bool IsWordSeparator(char c)
            {
                return char.IsSeparator(c) || char.IsPunctuation(c) || char.IsControl(c) || char.IsWhiteSpace(c) || c == '\r' || c == '\n';
            }
        }

        public char GetCharacter(int pageNumber, int index)
        {
            var pageData = GetPageData(_document, _form, pageNumber);
            {
                return NativeMethods.FPDFText_GetUnicode(pageData.TextPage, index);
            }
        }

        public int CountChars(int pageNumber)
        {
            var pageData = GetPageData(_document, _form, pageNumber);
            {
                return NativeMethods.FPDFText_CountChars(pageData.TextPage);
            }
        }

        public void DeletePage(int pageNumber)
        {
            NativeMethods.FPDFPage_Delete(_document, pageNumber);
        }

        public PdfRotation GetPageRotation(int pageNumber)
        {
            return NativeMethods.FPDFPage_GetRotation(GetPageData(_document, _form, pageNumber).Page);
        }

        public void RotatePage(int pageNumber, PdfRotation rotation)
        {
            var pageData = GetPageData(_document, _form, pageNumber);
            {
                NativeMethods.FPDFPage_SetRotation(pageData.Page, rotation);
            }
        }

        public PdfInformation GetInformation()
        {
            var pdfInfo = new PdfInformation();

            pdfInfo.Creator = GetMetaText("Creator");
            pdfInfo.Title = GetMetaText("Title");
            pdfInfo.Author = GetMetaText("Author");
            pdfInfo.Subject = GetMetaText("Subject");
            pdfInfo.Keywords = GetMetaText("Keywords");
            pdfInfo.Producer = GetMetaText("Producer");
            pdfInfo.CreationDate = GetMetaTextAsDate("CreationDate");
            pdfInfo.ModificationDate = GetMetaTextAsDate("ModDate");

            return pdfInfo;
        }

        private string GetMetaText(string tag)
        {
            // Length includes a trailing \0.

            uint length = NativeMethods.FPDF_GetMetaText(_document, tag, null, 0);
            if (length <= 2)
                return string.Empty;

            byte[] buffer = new byte[length];
            NativeMethods.FPDF_GetMetaText(_document, tag, buffer, length);

            return Encoding.Unicode.GetString(buffer, 0, (int)(length - 2));
        }

        public DateTime? GetMetaTextAsDate(string tag)
        {
            string dt = GetMetaText(tag);

            if (string.IsNullOrEmpty(dt))
                return null;

            Regex dtRegex =
                new Regex(
                    @"(?:D:)(?<year>\d\d\d\d)(?<month>\d\d)(?<day>\d\d)(?<hour>\d\d)(?<minute>\d\d)(?<second>\d\d)(?<tz_offset>[+-zZ])?(?<tz_hour>\d\d)?'?(?<tz_minute>\d\d)?'?");

            Match match = dtRegex.Match(dt);

            if (match.Success)
            {
                var year = match.Groups["year"].Value;
                var month = match.Groups["month"].Value;
                var day = match.Groups["day"].Value;
                var hour = match.Groups["hour"].Value;
                var minute = match.Groups["minute"].Value;
                var second = match.Groups["second"].Value;
                var tzOffset = match.Groups["tz_offset"]?.Value;
                var tzHour = match.Groups["tz_hour"]?.Value;
                var tzMinute = match.Groups["tz_minute"]?.Value;

                string formattedDate = $"{year}-{month}-{day}T{hour}:{minute}:{second}.0000000";

                if (!string.IsNullOrEmpty(tzOffset))
                {
                    switch (tzOffset)
                    {
                        case "Z":
                        case "z":
                            formattedDate += "+0";
                            break;
                        case "+":
                        case "-":
                            formattedDate += $"{tzOffset}{tzHour}:{tzMinute}";
                            break;
                    }
                }

                try
                {
                    return DateTime.Parse(formattedDate);
                }
                catch (FormatException)
                {
                    return null;
                }
            }

            return null;
        }

        /// <summary>
        /// Retrieves the text style information for a specific character on a pageNumber.
        /// </summary>
        /// <param name="pageNumber">The pageNumber number.</param>
        /// <param name="index">The character index on the pageNumber.</param>
        /// <returns>A PdfTextStyle object containing style information.</returns>
        public PdfTextStyle GetTextStyle(int pageNumber, int index)
        {
            if (_disposed)
                throw new ObjectDisposedException(GetType().Name);

            bool isUnderlined = false;
            bool isStrikethrough = false;
            bool isHighlighted = false;
            bool IsSquiggly = false;
            Color fillColor = Color.Empty;
            Color strokeColor = Color.Empty;
            Color annotationColor = Color.Empty;

            var pageData = GetPageData(_document, _form, pageNumber);
            {
                NativeMethods.FPDFText_GetFillColor(pageData.TextPage, index, out fillColor);
                NativeMethods.FPDFText_GetStrokeColor(pageData.TextPage, index, out strokeColor);

                // Get the character bounds
                NativeMethods.FPDFText_GetCharBox(
                    pageData.TextPage,
                    index,
                    out var charLeft,
                    out var charRight,
                    out var charBottom,
                    out var charTop
                );

                RectangleF charBounds = new RectangleF(
                    (float)charLeft,
                    (float)charTop,
                    (float)(charRight - charLeft),
                    (float)(charBottom - charTop)
                );

                // Detect underline, strikethrough, and highlight annotations
                int annotCount = NativeMethods.FPDFPage_GetAnnotCount(pageData.Page);
                for (int i = 0; i < annotCount; i++)
                {
                    IntPtr annotHandle = NativeMethods.FPDFPage_GetAnnot(pageData.Page, i);
                    if (annotHandle == IntPtr.Zero)
                        continue;

                    var subtype = NativeMethods.FPDFAnnot_GetSubtype(annotHandle);

                    // Check attachment points if available
                    int attachmentPointCount = (int)NativeMethods.FPDFAnnot_CountAttachmentPoints(annotHandle);
                    for (int j = 0; j < attachmentPointCount; j++)
                    {
                        NativeMethods.FPDFAnnot_GetAttachmentPoints(annotHandle, (UIntPtr)j, out var quadPoints);

                        RectangleF annotBounds = new RectangleF(
                            Math.Min(Math.Min(quadPoints.x1, quadPoints.x2), Math.Min(quadPoints.x3, quadPoints.x4)),
                            Math.Min(Math.Min(quadPoints.y1, quadPoints.y2), Math.Min(quadPoints.y3, quadPoints.y4)),
                            Math.Max(Math.Max(quadPoints.x1, quadPoints.x2), Math.Max(quadPoints.x3, quadPoints.x4)) - Math.Min(Math.Min(quadPoints.x1, quadPoints.x2), Math.Min(quadPoints.x3, quadPoints.x4)),
                            Math.Max(Math.Max(quadPoints.y1, quadPoints.y2), Math.Max(quadPoints.y3, quadPoints.y4)) - Math.Min(Math.Min(quadPoints.y1, quadPoints.y2), Math.Min(quadPoints.y3, quadPoints.y4))
                        );

                        // Check if the character bounds intersect with the annotation bounds
                        if (charBounds.IntersectsWith(annotBounds))
                        {
                            if (subtype == NativeMethods.FPDF_ANNOTATION_SUBTYPE.UNDERLINE)
                            {
                                isUnderlined = true;
                            }
                            else if (subtype == NativeMethods.FPDF_ANNOTATION_SUBTYPE.SQUIGGLY)
                            {
                                IsSquiggly = true;
                            }
                            else if (subtype == NativeMethods.FPDF_ANNOTATION_SUBTYPE.STRIKEOUT)
                            {
                                isStrikethrough = true;
                            }
                            else if (subtype == NativeMethods.FPDF_ANNOTATION_SUBTYPE.HIGHLIGHT)
                            {
                                isHighlighted = true;
                                NativeMethods.FPDFAnnot_GetColor(annotHandle, NativeMethods.FPDFANNOT_COLORTYPE.Color, out annotationColor);
                            }
                        }
                    }

                    NativeMethods.FPDFPage_CloseAnnot(annotHandle);
                }
            }

            return new PdfTextStyle(pageNumber, index, fillColor, strokeColor, isUnderlined, isStrikethrough, isHighlighted, IsSquiggly, annotationColor);
        }

        private PageData GetPageData(IntPtr document, IntPtr form, int pageNumber)
        {
            if (_pageData[pageNumber] == null)
            {
                _pageData[pageNumber] = new PageData(document, form, pageNumber);
            }

            return _pageData[pageNumber];

        }

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                StreamManager.Unregister(_id);

                foreach( var pageData in _pageData)
                {
                    if (pageData != null)
                    {
                        pageData.Dispose();
                    }
                }

                if (_form != IntPtr.Zero)
                {
                    NativeMethods.FORM_DoDocumentAAction(_form, NativeMethods.FPDFDOC_AACTION.WC);
                    NativeMethods.FPDFDOC_ExitFormFillEnvironment(_form);
                    _form = IntPtr.Zero;
                }

                if (_document != IntPtr.Zero)
                {
                    NativeMethods.FPDF_CloseDocument(_document);
                    _document = IntPtr.Zero;
                }

                if (_formCallbacksHandle.IsAllocated)
                    _formCallbacksHandle.Free();

                if (_stream != null)
                {
                    _stream.Dispose();
                    _stream = null;
                }

                _disposed = true;
            }
        }

        private class PageData : IDisposable
        {
            private readonly IntPtr _form;
            private bool _disposed;

            public IntPtr Page { get; private set; }

            public IntPtr TextPage { get; private set; }

            public double Width { get; private set; }

            public double Height { get; private set; }

            public PageData(IntPtr document, IntPtr form, int pageNumber)
            {
                _form = form;

                Page = NativeMethods.FPDF_LoadPage(document, pageNumber);
                TextPage = NativeMethods.FPDFText_LoadPage(Page);
                if (_form != null)
                {
                    NativeMethods.FORM_OnAfterLoadPage(Page, form);
                    NativeMethods.FORM_DoPageAAction(Page, form, NativeMethods.FPDFPAGE_AACTION.OPEN);
                }
                Width = NativeMethods.FPDF_GetPageWidth(Page);
                Height = NativeMethods.FPDF_GetPageHeight(Page);
            }

            public void Dispose()
            {
                if (!_disposed)
                {
                    if (_form != null)
                    {
                        NativeMethods.FORM_DoPageAAction(Page, _form, NativeMethods.FPDFPAGE_AACTION.CLOSE);
                        NativeMethods.FORM_OnBeforeClosePage(Page, _form);
                    }
                    NativeMethods.FPDFText_ClosePage(TextPage);
                    NativeMethods.FPDF_ClosePage(Page);

                    _disposed = true;
                }
            }
        }
    }
}
