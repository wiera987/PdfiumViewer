using System;
using System.Collections.Generic;
using System.Text;

namespace PdfiumViewer
{
    partial class PdfViewer
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PdfViewer));
            this._toolStrip = new System.Windows.Forms.ToolStrip();
            this._saveButton = new System.Windows.Forms.ToolStripButton();
            this._printButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this._zoomInButton = new System.Windows.Forms.ToolStripButton();
            this._zoomOutButton = new System.Windows.Forms.ToolStripButton();
            this._container = new System.Windows.Forms.SplitContainer();
            this.toolStripBookmark = new System.Windows.Forms.ToolStrip();
            this.toolStripButtonExpand = new System.Windows.Forms.ToolStripButton();
            this.toolStripSplitButtonCollapse = new System.Windows.Forms.ToolStripSplitButton();
            this.expandAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.collapseAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripButtonExpandCurrent = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonClose = new System.Windows.Forms.ToolStripButton();
            this._bookmarks = new PdfiumViewer.NativeTreeView();
            this._renderer = new PdfiumViewer.PdfRenderer();
            this._toolStrip.SuspendLayout();
            this._container.Panel1.SuspendLayout();
            this._container.Panel2.SuspendLayout();
            this._container.SuspendLayout();
            this.toolStripBookmark.SuspendLayout();
            this.SuspendLayout();
            // 
            // _toolStrip
            // 
            this._toolStrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this._toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this._saveButton,
            this._printButton,
            this.toolStripSeparator1,
            this._zoomInButton,
            this._zoomOutButton});
            resources.ApplyResources(this._toolStrip, "_toolStrip");
            this._toolStrip.Name = "_toolStrip";
            // 
            // _saveButton
            // 
            this._saveButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this._saveButton.Image = global::PdfiumViewer.Properties.Resources.disk_blue;
            resources.ApplyResources(this._saveButton, "_saveButton");
            this._saveButton.Name = "_saveButton";
            this._saveButton.Click += new System.EventHandler(this._saveButton_Click);
            // 
            // _printButton
            // 
            this._printButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this._printButton.Image = global::PdfiumViewer.Properties.Resources.printer;
            resources.ApplyResources(this._printButton, "_printButton");
            this._printButton.Name = "_printButton";
            this._printButton.Click += new System.EventHandler(this._printButton_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            resources.ApplyResources(this.toolStripSeparator1, "toolStripSeparator1");
            // 
            // _zoomInButton
            // 
            this._zoomInButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this._zoomInButton.Image = global::PdfiumViewer.Properties.Resources.zoom_in;
            resources.ApplyResources(this._zoomInButton, "_zoomInButton");
            this._zoomInButton.Name = "_zoomInButton";
            this._zoomInButton.Click += new System.EventHandler(this._zoomInButton_Click);
            // 
            // _zoomOutButton
            // 
            this._zoomOutButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this._zoomOutButton.Image = global::PdfiumViewer.Properties.Resources.zoom_out;
            resources.ApplyResources(this._zoomOutButton, "_zoomOutButton");
            this._zoomOutButton.Name = "_zoomOutButton";
            this._zoomOutButton.Click += new System.EventHandler(this._zoomOutButton_Click);
            // 
            // _container
            // 
            resources.ApplyResources(this._container, "_container");
            this._container.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this._container.Name = "_container";
            // 
            // _container.Panel1
            // 
            this._container.Panel1.Controls.Add(this.toolStripBookmark);
            this._container.Panel1.Controls.Add(this._bookmarks);
            // 
            // _container.Panel2
            // 
            this._container.Panel2.Controls.Add(this._renderer);
            this._container.TabStop = false;
            // 
            // toolStripBookmark
            // 
            this.toolStripBookmark.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButtonExpand,
            this.toolStripSplitButtonCollapse,
            this.toolStripButtonExpandCurrent,
            this.toolStripButtonClose});
            resources.ApplyResources(this.toolStripBookmark, "toolStripBookmark");
            this.toolStripBookmark.Name = "toolStripBookmark";
            // 
            // toolStripButtonExpand
            // 
            this.toolStripButtonExpand.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            resources.ApplyResources(this.toolStripButtonExpand, "toolStripButtonExpand");
            this.toolStripButtonExpand.Name = "toolStripButtonExpand";
            this.toolStripButtonExpand.Click += new System.EventHandler(this.toolStripButtonExpand_Click);
            // 
            // toolStripSplitButtonCollapse
            // 
            this.toolStripSplitButtonCollapse.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripSplitButtonCollapse.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.expandAllToolStripMenuItem,
            this.collapseAllToolStripMenuItem});
            resources.ApplyResources(this.toolStripSplitButtonCollapse, "toolStripSplitButtonCollapse");
            this.toolStripSplitButtonCollapse.Name = "toolStripSplitButtonCollapse";
            this.toolStripSplitButtonCollapse.ButtonClick += new System.EventHandler(this.toolStripSplitButtonCollapse_ButtonClick);
            // 
            // expandAllToolStripMenuItem
            // 
            this.expandAllToolStripMenuItem.Name = "expandAllToolStripMenuItem";
            resources.ApplyResources(this.expandAllToolStripMenuItem, "expandAllToolStripMenuItem");
            this.expandAllToolStripMenuItem.Click += new System.EventHandler(this.expandAllToolStripMenuItem_Click);
            // 
            // collapseAllToolStripMenuItem
            // 
            this.collapseAllToolStripMenuItem.Name = "collapseAllToolStripMenuItem";
            resources.ApplyResources(this.collapseAllToolStripMenuItem, "collapseAllToolStripMenuItem");
            this.collapseAllToolStripMenuItem.Click += new System.EventHandler(this.collapseAllToolStripMenuItem_Click);
            // 
            // toolStripButtonExpandCurrent
            // 
            this.toolStripButtonExpandCurrent.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            resources.ApplyResources(this.toolStripButtonExpandCurrent, "toolStripButtonExpandCurrent");
            this.toolStripButtonExpandCurrent.Name = "toolStripButtonExpandCurrent";
            this.toolStripButtonExpandCurrent.Click += new System.EventHandler(this.toolStripButtonExpandCurrent_Click);
            // 
            // toolStripButtonClose
            // 
            this.toolStripButtonClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripButtonClose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            resources.ApplyResources(this.toolStripButtonClose, "toolStripButtonClose");
            this.toolStripButtonClose.Name = "toolStripButtonClose";
            this.toolStripButtonClose.Click += new System.EventHandler(this.toolStripButtonClose_Click);
            // 
            // _bookmarks
            // 
            resources.ApplyResources(this._bookmarks, "_bookmarks");
            this._bookmarks.FullRowSelect = true;
            this._bookmarks.Name = "_bookmarks";
            this._bookmarks.ShowLines = false;
            this._bookmarks.BeforeCollapse += new System.Windows.Forms.TreeViewCancelEventHandler(this._bookmarks_BeforeCollapse);
            this._bookmarks.AfterCollapse += new System.Windows.Forms.TreeViewEventHandler(this._bookmarks_AfterCollapse);
            this._bookmarks.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this._bookmarks_BeforeExpand);
            this._bookmarks.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this._bookmarks_AfterExpand);
            this._bookmarks.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this._bookmarks_AfterSelect);
            // 
            // _renderer
            // 
            this._renderer.CompareBounds = ((System.Drawing.RectangleF)(resources.GetObject("_renderer.CompareBounds")));
            this._renderer.Cursor = System.Windows.Forms.Cursors.Default;
            resources.ApplyResources(this._renderer, "_renderer");
            this._renderer.Name = "_renderer";
            this._renderer.Page = 0;
            this._renderer.Rotation = PdfiumViewer.PdfRotation.Rotate0;
            this._renderer.ZoomMode = PdfiumViewer.PdfViewerZoomMode.FitHeight;
            this._renderer.LinkClick += new PdfiumViewer.LinkClickEventHandler(this._renderer_LinkClick);
            // 
            // PdfViewer
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this._container);
            this.Controls.Add(this._toolStrip);
            this.Name = "PdfViewer";
            this._toolStrip.ResumeLayout(false);
            this._toolStrip.PerformLayout();
            this._container.Panel1.ResumeLayout(false);
            this._container.Panel1.PerformLayout();
            this._container.Panel2.ResumeLayout(false);
            this._container.ResumeLayout(false);
            this.toolStripBookmark.ResumeLayout(false);
            this.toolStripBookmark.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip _toolStrip;
        private System.Windows.Forms.ToolStripButton _saveButton;
        private System.Windows.Forms.ToolStripButton _printButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton _zoomInButton;
        private System.Windows.Forms.ToolStripButton _zoomOutButton;
        private System.Windows.Forms.SplitContainer _container;
        private NativeTreeView _bookmarks;
        private PdfRenderer _renderer;
        private System.Windows.Forms.ToolStrip toolStripBookmark;
        private System.Windows.Forms.ToolStripButton toolStripButtonExpand;
        private System.Windows.Forms.ToolStripSplitButton toolStripSplitButtonCollapse;
        private System.Windows.Forms.ToolStripMenuItem collapseAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolStripButtonExpandCurrent;
        private System.Windows.Forms.ToolStripMenuItem expandAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolStripButtonClose;
    }
}
