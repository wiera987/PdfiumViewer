using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Expando;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace PdfiumViewer
{
    /// <summary>
    /// Control to host PDF documents with support for printing.
    /// </summary>
    public partial class PdfViewer : UserControl
    {
        public event EventHandler PanelBookmarkClosed;

        // Constants for the SendMessage() method.
        private const int WM_HSCROLL = 276;
        private const int SB_LEFT = 6;
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int wMsg,
                                          int wParam, int lParam);

        private IPdfDocument _document;
        private bool _showBookmarks;
        private bool preventPageRefresh;
        private int bookmarkPage;

        /// <summary>
        /// Gets or sets the PDF document.
        /// </summary>
        [DefaultValue(null)]
        public IPdfDocument Document
        {
            get { return _document; }
            set
            {
                if (_document != value)
                {
                    _document = value;

                    if (_document != null)
                    {
                        _renderer.Load(_document);
                        UpdateBookmarks(true);
                    }

                    UpdateEnabled();
                }
            }
        }

        /// <summary>
        /// Get the <see cref="PdfRenderer"/> that renders the PDF document.
        /// </summary>
        public PdfRenderer Renderer
        {
            get { return _renderer; }
        }

        /// <summary>
        /// Gets or sets the default document name used when saving the document.
        /// </summary>
        [DefaultValue(null)]
        public string DefaultDocumentName { get; set; }

        /// <summary>
        /// Gets or sets the default print mode.
        /// </summary>
        [DefaultValue(PdfPrintMode.CutMargin)]
        public PdfPrintMode DefaultPrintMode { get; set; }

        /// <summary>
        /// Gets or sets the way the document should be zoomed initially.
        /// </summary>
        [DefaultValue(PdfViewerZoomMode.FitHeight)]
        public PdfViewerZoomMode ZoomMode
        {
            get { return _renderer.ZoomMode; }
            set { _renderer.ZoomMode = value; }
        }

        /// <summary>
        /// Gets or sets whether the toolbar should be shown.
        /// </summary>
        [DefaultValue(true)]
        public bool ShowToolbar
        {
            get { return _toolStrip.Visible; }
            set { _toolStrip.Visible = value; }
        }

        /// <summary>
        /// Gets or sets whether the bookmarks panel should be shown.
        /// </summary>
        [DefaultValue(true)]
        public bool ShowBookmarks
        {
            get { return _showBookmarks; }
            set
            {
                _showBookmarks = value;
                UpdateBookmarks();
            }
        }

        /// <summary>
        /// Gets or sets the pre-selected printer to be used when the print
        /// dialog shows up.
        /// </summary>
        [DefaultValue(null)]
        public string DefaultPrinter { get; set; }

        /// <summary>
        /// Occurs when a link in the pdf document is clicked.
        /// </summary>
        [Category("Action")]
        [Description("Occurs when a link in the pdf document is clicked.")]
        public event LinkClickEventHandler LinkClick;

        /// <summary>
        /// Called when a link is clicked.
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnLinkClick(LinkClickEventArgs e)
        {
            var handler = LinkClick;
            if (handler != null)
                handler(this, e);
        }

        private void UpdateBookmarks(bool construct = false)
        {
            bool visible = _showBookmarks && _document != null && _document.Bookmarks.Count > 0;

            _container.Panel1Collapsed = !visible;

            if (visible || construct)
            {
                if ((_bookmarks.Nodes.Count == 0) || construct)
                {
                    _bookmarks.Nodes.Clear();
                    foreach (var bookmark in _document.Bookmarks)
                    {
                        _bookmarks.Nodes.Add(GetBookmarkNode(bookmark));
                    }
                }

                // Select a bookmark for the renderer page.
                bookmarkPage = -1;
                SelectBookmarkForPage(_renderer.Page);
            }
        }

        /// <summary>
        /// Select a bookmark for a specified page.
        /// Do not expand new nodes.
        /// </summary>
        public void SelectBookmarkForPage(int page)
        {
            TreeNode validNode = null;
            int bmpage = bookmarkPage;

            if (page != bookmarkPage)
            {
                GetPageNode(page, _bookmarks.Nodes, false, ref validNode);

                // Select the last valid Node.
                if (validNode != null)
                {
                    preventPageRefresh = true;
                    _bookmarks.SelectedNode = validNode;
                    preventPageRefresh = false;
                    bookmarkPage = page;
                }
            }
            //Console.WriteLine("SelectBookmarkForPage: Page={0}!={1}, Text={2}", page, bmpage, validNode?.Text);
        }

        /// <summary>
        /// Recursively search for the node that corresponds to the page from Nodes.
        /// </summary>
        /// <param name="page">The page being selected.</param>
        /// <param name="nodes">Bookmark Nodes being searched recursively</param>
        /// <param name="expand">If true, search while expanding collapsed nodes</param>
        /// <param name="validNode">Currently active Node</param>
        private void GetPageNode(int page, TreeNodeCollection nodes, bool expand, ref TreeNode validNode)
        {
            int validPage = -1;

            foreach (TreeNode node in nodes)
            {
                PdfBookmark bookmark = (PdfBookmark)node.Tag;

                if (bookmark.PageIndex <= page)
                {
                    // The last updated node is valid.
                    // However, if they are on the same page, select the node found first.
                    // Prefer child node over parent node.
                    if (bookmark.PageIndex != validPage)
                    {
                        validNode = node;
                        validPage = bookmark.PageIndex;
                    }
                }

                // If the tree is expanded, explore further.
                // If 'expand' is true, explore nodes even if they are collapsed.
                if (node.IsExpanded || (expand && (node.Nodes != null)))
                {
                    GetPageNode(page, node.Nodes, expand, ref validNode);
                }
            }
        }

        /// <summary>
        /// Get the start page and end page of the selected bookmark node.
        /// </summary>
        /// <returns>Tuple containing the start page and end page of the current bookmark node.</returns>
        public void GetBookmarkPageRange(int page, out int startPage, out int endPage)
        {
            if (_bookmarks.SelectedNode != null)
            {
                TreeNode currentNode = _bookmarks.SelectedNode;
                PdfBookmark currentBookmark = (PdfBookmark)currentNode.Tag;
                startPage = currentBookmark.PageIndex;

                TreeNode endNode = GetEndNode(currentNode);
                if (endNode != null)
                {
                    // When the node after the current node is found.
                    PdfBookmark endBookmark = (PdfBookmark)endNode.Tag;
                    endPage = endBookmark.PageIndex - 1;
                }
                else
                {
                    // If there is no node, it is the last page.
                    endPage = _document.PageCount - 1;
                }

                if ((page < startPage) || (page > endPage))
                {
                    // In case the target page lacks an assigned bookmark.
                    startPage = -1;
                    endPage = -1;
                }
            }
            else
            {
                startPage = -1;
                endPage = -1;
            }
            //Console.WriteLine("GetBookmarkPageRange(): {0} - {1}", startPage, endPage);
        }


        /// <summary>
        /// Get the end node of the current bookmark node.
        /// When the end node at the same level cannot be found, recursively search for the parent node.
        /// </summary>
        /// <param name="currentNode"></param>
        /// <returns></returns>
        private TreeNode GetEndNode(TreeNode currentNode)
        {
            if (currentNode.NextNode != null)
            {
                return currentNode.NextNode;
            }
            else if (currentNode.Parent != null)
            {
                return GetEndNode(currentNode.Parent);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Initializes a new instance of the PdfViewer class.
        /// </summary>
        public PdfViewer()
        {
            DefaultPrintMode = PdfPrintMode.CutMargin;

            InitializeComponent();

            ShowToolbar = true;
            ShowBookmarks = true;

            _bookmarks.HideSelection = false;
            preventPageRefresh = false;
            bookmarkPage = -1;

            UpdateEnabled();
        }

        private void UpdateEnabled()
        {
            _toolStrip.Enabled = _document != null;
        }

        private void _zoomInButton_Click(object sender, EventArgs e)
        {
            _renderer.ZoomIn();
        }

        private void _zoomOutButton_Click(object sender, EventArgs e)
        {
            _renderer.ZoomOut();
        }

        private void _saveButton_Click(object sender, EventArgs e)
        {
            using (var form = new SaveFileDialog())
            {
                form.DefaultExt = ".pdf";
                form.Filter = Properties.Resources.SaveAsFilter;
                form.RestoreDirectory = true;
                form.Title = Properties.Resources.SaveAsTitle;
                form.FileName = DefaultDocumentName;

                if (form.ShowDialog(FindForm()) == DialogResult.OK)
                {
                    try
                    {
                        _document.Save(form.FileName);
                    }
                    catch
                    {
                        MessageBox.Show(
                            FindForm(),
                            Properties.Resources.SaveAsFailedText,
                            Properties.Resources.SaveAsFailedTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                }
            }
        }

        private void _printButton_Click(object sender, EventArgs e)
        {
            using (var form = new PrintDialog())
            using (var document = _document.CreatePrintDocument(DefaultPrintMode))
            {
                form.AllowSomePages = true;
                form.Document = document;
                form.UseEXDialog = true;
                form.Document.PrinterSettings.FromPage = 1;
                form.Document.PrinterSettings.ToPage = _document.PageCount;
                if (DefaultPrinter != null)
                    form.Document.PrinterSettings.PrinterName = DefaultPrinter;

                if (form.ShowDialog(FindForm()) == DialogResult.OK)
                {
                    try
                    {
                        if (form.Document.PrinterSettings.FromPage <= _document.PageCount)
                            form.Document.Print();
                    }
                    catch
                    {
                        // Ignore exceptions; the printer dialog should take care of this.
                    }
                }
            }
        }

        private TreeNode GetBookmarkNode(PdfBookmark bookmark)
        {
            TreeNode node = new TreeNode(bookmark.Title);
            node.Tag = bookmark;
            if (bookmark.Children != null)
            {
                foreach (var child in bookmark.Children)
                {
                    node.Nodes.Add(GetBookmarkNode(child));
                }
            }
            return node;
        }

        private void _bookmarks_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (preventPageRefresh == false)
            {
                _renderer.Page = ((PdfBookmark)e.Node.Tag).PageIndex;
            }
        }

        private void _bookmarks_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            // AfterSelect event is fired when child node has focus.
            preventPageRefresh = true;
        }

        private void _bookmarks_AfterExpand(object sender, TreeViewEventArgs e)
        {
            // AfterSelect event is fired when child node has focus.
            preventPageRefresh = false;

            // Select a bookmark for the renderer page.
            bookmarkPage = -1;
            SelectBookmarkForPage(_renderer.Page);
        }

        private void _bookmarks_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            // AfterSelect event is fired when child node has focus.
            preventPageRefresh = true;
        }

        private void _bookmarks_AfterCollapse(object sender, TreeViewEventArgs e)
        {
            // AfterSelect event is fired when child node has focus.
            preventPageRefresh = false;
        }

        private void _renderer_LinkClick(object sender, LinkClickEventArgs e)
        {
            OnLinkClick(e);
        }

        private void toolStripButtonExpand_Click(object sender, EventArgs e)
        {
            // Expand Top-Level
            bool visible = _showBookmarks && _document != null && _document.Bookmarks.Count > 0;

            if (visible)
            {
                foreach (TreeNode node in _bookmarks.Nodes)
                {
                    node.Expand();
                }

                // Show selected nodes.
                EnsureVisibleWithoutRightScrolling(_bookmarks.SelectedNode);
            }
        }

        private void toolStripSplitButtonCollapse_ButtonClick(object sender, EventArgs e)
        {
            // Collapse Top-Level
            bool visible = _showBookmarks && _document != null && _document.Bookmarks.Count > 0;

            if (visible)
            {
                foreach (TreeNode node in _bookmarks.Nodes)
                {
                    node.Collapse();
                }
                // Show selected nodes.
                EnsureVisibleWithoutRightScrolling(_bookmarks.SelectedNode);
            }
        }

        private void expandAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Expand All
            _bookmarks.ExpandAll();
            // Show selected nodes.
            EnsureVisibleWithoutRightScrolling(_bookmarks.SelectedNode);
        }

        private void collapseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Collapse All
            CollapseChildrenNodes(_bookmarks.Nodes);
            // Show selected nodes.
            EnsureVisibleWithoutRightScrolling(_bookmarks.SelectedNode);
        }

        private void toolStripButtonExpandCurrent_Click(object sender, EventArgs e)
        {
            // Expand current bookmark
            TreeNode validNode = null;

            // Collapse all nodes initially
            CollapseChildrenNodes(_bookmarks.Nodes);
            GetPageNode(_renderer.Page, _bookmarks.Nodes, true, ref validNode);
            ExpandParentNodes(validNode);

            // Show selected nodes.
            EnsureVisibleWithoutRightScrolling(_bookmarks.SelectedNode);
        }

        private void ExpandParentNodes(TreeNode node)
        {
            if (node != null)
            {
                node.Expand();
                if (node.Parent != null)
                {
                    ExpandParentNodes(node.Parent);
                }
            }
        }

        private void CollapseChildrenNodes(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                if (node != null)
                {
                    if (node.Nodes != null)
                    {
                        CollapseChildrenNodes(node.Nodes);
                    }
                    node.Collapse();
                }
            }
        }

        private void EnsureVisibleWithoutRightScrolling(TreeNode node)
        {
            // we do the standard call.. 
            node?.EnsureVisible();

            // ..and afterwards we scroll to the left again!
            SendMessage(_bookmarks.Handle, WM_HSCROLL, SB_LEFT, 0);
        }

        private void toolStripButtonClose_Click(object sender, EventArgs e)
        {
            ShowBookmarks = false;

            // Fire a custom event when a bookmark is closed
            PanelBookmarkClosed?.Invoke(this, e);
        }

    }
}
