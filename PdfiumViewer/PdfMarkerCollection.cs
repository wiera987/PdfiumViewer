using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

#pragma warning disable 1591

namespace PdfiumViewer
{
    public class PdfMarkerCollection : Collection<IPdfMarker>
    {
        public event EventHandler CollectionChanged;

        /// <summary>
        /// Clear specified tag.
        /// </summary>
        /// <param name="tag"></param>
        public void Clear(int tag)
        {
            for (int i=Count-1; i>=0; i--)
            {
                if (Items[i].Tag == tag)
                {
                    RemoveAt(i);
                }
            }
        }

        protected override void ClearItems()
        {
            base.ClearItems();

            OnCollectionChanged(EventArgs.Empty);
        }

        protected override void InsertItem(int index, IPdfMarker item)
        {
            base.InsertItem(index, item);

            OnCollectionChanged(EventArgs.Empty);
        }

        protected override void RemoveItem(int index)
        {
            base.RemoveItem(index);

            OnCollectionChanged(EventArgs.Empty);
        }

        protected override void SetItem(int index, IPdfMarker item)
        {
            base.SetItem(index, item);

            OnCollectionChanged(EventArgs.Empty);
        }

        protected virtual void OnCollectionChanged(EventArgs e)
        {
            CollectionChanged?.Invoke(this, e);
        }
    }
}
