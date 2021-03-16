//------------------------------------------------------------------------------
// <copyright file="Gallery.cs" company="Alton X Lam">
// Copyright © 2013 Alton Lam
//
// This software is provided 'as-is', without any express or implied
// warranty. In no event will the authors be held liable for any damages
// arising from the use of this software.
//
// Permission is granted to anyone to use this software for any purpose,
// including commercial applications, and to alter it and redistribute it
// freely, subject to the following restrictions:
//
// 1. The origin of this software must not be misrepresented; you must not
//    claim that you wrote the original software. If you use this software
//    in a product, an acknowledgment in the product documentation would be
//    appreciated but is not required.
//
// 2. Altered source versions must be plainly marked as such, and must not be
//    misrepresented as being the original software.
//
// 3. This notice may not be removed or altered from any source distribution.
//
// Alton X Lam
// https://www.facebook.com/AltonXL
// https://www.codeplex.com/site/users/view/AltonXL
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.IconLib;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ImageMso.Excel
{
    ///<summary>Image gallery of unqiue ImageMso icons gathered by a composite of multiple sources.</summary>
    public partial class Gallery : Form, IExcelAddIn
    {
        ///<summary>ComVisible class for Ribbon call back and event handling methods.</summary>
        [ComVisible(true)]
        public class Ribbon : ExcelRibbon
        {
            public void OnLoad(IRibbonUI ribbonUI)
            {
            }

            ///<summary>Handles the ribbon control button OnClick Event. Activates the existing gallery or builds a new gallery.</summary>
            public void OnAction(IRibbonControl control)
            {
                if (Gallery.Default == null || Gallery.Default.IsDisposed) (Gallery.Default = new Gallery()).Show();
                else Gallery.Default.Activate();
            }
        }

        ///<summary>Underlines the hotkey letter for tool strip menu items.</summary>
        private class HotkeyMenuStripRenderer : ToolStripProfessionalRenderer
        {
            public HotkeyMenuStripRenderer() : base() { }
            public HotkeyMenuStripRenderer(ProfessionalColorTable table) : base(table) { }

            protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
            {
                e.TextFormat &= ~TextFormatFlags.HidePrefix;
                base.OnRenderItemText(e);
            }
        }

        ///<summary>Default instance for static referencing.</summary>
        public static Gallery Default { get; private set; }
        private Images imageMso = Images.Default;
        private KeyEventArgs key = null; // Holds the key state between KeyDown and KeyUp events.
        private string size = string.Empty; // Holds the image size between GotFocus and LostFocus events.
        private string pattern = string.Empty; // Holds the filter pattern between GotFocus and LostFocus events.

        /// <summary>Initializes a new instance of the ImageMso.Gallery class.</summary>
        public Gallery()
        {
            InitializeComponent();

            GalleryMenu.Renderer = new HotkeyMenuStripRenderer();

            if (Default == null) Default = this;

            Icons.SmallImageList = new ImageList();
            Icons.LargeImageList = new ImageList();
            Icons.SmallImageList.ImageSize = new Size(16, 16);
            Icons.LargeImageList.ImageSize = new Size(32, 32);
            Icons.SmallImageList.ColorDepth = ColorDepth.Depth32Bit;
            Icons.LargeImageList.ColorDepth = ColorDepth.Depth32Bit;

            var names = new string[imageMso.Names.Count];
            imageMso.Names.CopyTo(names, 0);
            Icons.BeginUpdate();
            var smallImages = Array.ConvertAll(names, name => imageMso[name, 16, 16]).Where(img => img != null).ToArray();
            var largeImages = Array.ConvertAll(names, name => imageMso[name, 32, 32]).Where(img => img != null).ToArray();
            Icons.SmallImageList.Images.AddRange(smallImages);
            Icons.LargeImageList.Images.AddRange(largeImages);
            Icons.Items.AddRange(Array.ConvertAll(names, name => new ListViewItem(name, Array.IndexOf(names, name))));
            Icons.EndUpdate();

            Small.Checked = Icons.View == View.List;
            Large.Checked = Icons.View == View.LargeIcon;

            using (var image = imageMso["DesignAccentsGallery", 32, 32] ?? imageMso["GroupSmartArtQuickStyles", 32, 32])
                if (image != null) Icon = Icon.FromHandle(image.GetHicon());
            ViewSize.Image = imageMso["ListView", 16, 16];
            Large.Image = imageMso["LargeIcons", 16, 16] ?? imageMso["SmartArtLargerShape", 16, 16];
            Small.Image = imageMso["SmallIcons", 16, 16] ?? imageMso["SmartArtSmallerShape", 16, 16];
            Filter.Image = imageMso["FiltersMenu", 16, 16];
            Clear.Image = imageMso["FilterClearAllFilters", 16, 16];
            CopyPicture.Image = imageMso["CopyPicture", 16, 16];
            CopyText.Image = imageMso["Copy", 16, 16];
            SelectAll.Image = imageMso["SelectAll", 16, 16];
            Dimensions.Image = imageMso["PicturePropertiesSize", 16, 16] ?? imageMso["SizeAndPositionWindow", 16, 16];
            SaveAs.Image = imageMso["PictureFormatDialog", 16, 16];
            Save.Image = imageMso["FileSave", 16, 16];
            Bmp.Image = imageMso["SaveAsBmp", 16, 16];
            Gif.Image = imageMso["SaveAsGif", 16, 16];
            Jpg.Image = imageMso["SaveAsJpg", 16, 16];
            Png.Image = imageMso["SaveAsPng", 16, 16];
            Tif.Image = imageMso["SaveAsTiff", 16, 16];
            Ico.Image = imageMso["InsertImageHtmlTag", 16, 16];
            Background.Image = imageMso["FontColorCycle", 16, 16];
            ExportAs.Image = imageMso["SlideShowResolutionGallery", 16, 16];
            Export.Image = imageMso["ArrangeBySize", 16, 16];
            Support.Image = imageMso["Help", 16, 16];
            Information.Image = imageMso["Info", 16, 16];

            Background.Text = string.Format(Background.Text, Palette.Color.Name);
        }

        ///<summary>Runs when the Excel Add-In is attached. Renders the gallery and brings the form into focus.</summary>
        public void AutoOpen()
        {
            Application.EnableVisualStyles();

            this.Show();
            this.Activate();
        }

        ///<summary>Runs when the Excel Add-In is detached. Cleans up resources no longer being used.</summary>
        public void AutoClose()
        {
        }

        private void Gallery_Disposed(object sender, System.EventArgs e)
        {
            DestroyIcon(Icon.Handle);
        }

        private void Gallery_KeyDown(object sender, KeyEventArgs e)
        {
            key = e;
        }

        private void Gallery_KeyUp(object sender, KeyEventArgs e)
        {
            if (key != null && key.Alt)
            {
                GalleryMenu.Show(RectangleToScreen(Icons.ClientRectangle).Left,
                    RectangleToScreen(Icons.ClientRectangle).Top);
                e.SuppressKeyPress = true;
            }
            key = e;
        }

        private void Icons_ItemDrag(object sender, ItemDragEventArgs e)
        {   // The MemoryStream containing the image must cycle through the clipboard to Drag & Drop correctly.
            // Mysteriously, the direct approach loses the alpha channel (background transparency) or will not drop.
            CopyPicture_Click(sender, e);
            Icons.DoDragDrop(Clipboard.GetDataObject(), DragDropEffects.All);
        }

        private void Icons_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            CopyPicture.Enabled = Icons.FocusedItem.Selected;
            CopyText.Enabled = Icons.SelectedItems.Count > 0;
            Export.Enabled = Icons.SelectedItems.Count > 0;
            Save.Enabled = Icons.SelectedItems.Count > 0;
        }

        private void ViewSize_Check(object sender, EventArgs e)
        {
            if (sender == Large && Icons.View != View.LargeIcon)
            {
                Icons.View = View.LargeIcon;
                Copy32px.PerformClick();
            }
            if (sender == Small && Icons.View != View.List)
            {
                Icons.View = View.List;
                Copy16px.PerformClick();
            }
            Large.Checked = Icons.View == View.LargeIcon;
            Small.Checked = Icons.View == View.List;
        }

        private void SetFilters_GotFocus(object sender, EventArgs e)
        {
            pattern = SetFilters.Text;
        }

        private void SetFilters_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SetFilters_LostFocus(sender, e);
        }

        private void SetFilters_LostFocus(object sender, EventArgs e)
        {
            SetFilters.Text = SetFilters.Text.Trim();
            if (!SetFilters.AutoCompleteCustomSource.Contains(SetFilters.Text))
                SetFilters.AutoCompleteCustomSource.Add(SetFilters.Text);
            if (SetFilters.Text != pattern)
            {
                pattern = SetFilters.Text;
                Clear.Enabled = SetFilters.Text.Length > 0;
                Filter.Checked = SetFilters.Text.Length > 0;
                // Comma separated array of keyword groups which are in turn space separated arrays of keywords.
                var filter = Array.ConvertAll(pattern.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries),
                    search => search.Trim().Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                var names = new string[imageMso.Names.Count];
                imageMso.Names.CopyTo(names, 0);
                Icons.BeginUpdate();
                Icons.Items.Clear();
                // Find all names where there exists a keyword group in the filter such that the name contains
                // all the keywords in the group, convert all names found to listview items and add as a range.
                if (filter.Length > 0) Icons.Items.AddRange(Array.ConvertAll(Array.FindAll(names,
                    name => Array.Exists(filter, group => Array.TrueForAll(group,
                        keyword => name.ToLowerInvariant().Contains(keyword.ToLowerInvariant().Trim())))),
                        name => new ListViewItem(name, Array.IndexOf(names, name))));
                else Icons.Items.AddRange(Array.ConvertAll(names, name => new ListViewItem(name, Array.IndexOf(names, name))));
                if (Icons.Items.Count > 0) Icons.FocusedItem = Icons.Items[0];
                Icons.EndUpdate();
            }
        }

        private void Clear_Click(object sender, EventArgs e)
        {
            if (Filter.Checked)
            {
                SetFilters.Clear();
                SetFilters_LostFocus(sender, e);
            }
        }

        private void CopyPicture_Click(object sender, EventArgs e)
        {
            short[] dims = Parse(Pixels.Text.ToLowerInvariant().Trim());
            using (var image = imageMso[Icons.FocusedItem.Text, dims[0], dims[1]])
            // Clipboard.SetImage(image); // Loses alpha channel (background transparency). Convert to MemoryStream instead.
            using (var stream = new MemoryStream())
            {
                image.Save(stream, ImageFormat.Png);
                Clipboard.SetData("PNG", stream);
            }
        }

        private void CopyText_Click(object sender, EventArgs e)
        {
            var items = new ListViewItem[Icons.SelectedItems.Count];
            Icons.SelectedItems.CopyTo(items, 0);
            string value = string.Join(Environment.NewLine, Array.ConvertAll(items, item => item.Text)).Trim();
            if (!string.IsNullOrEmpty(value)) Clipboard.SetText(value);
        }

        private void SelectAll_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in Icons.Items) item.Selected = true;
        }

        private void Dimensions_Check(object sender, EventArgs e)
        {
            Pixels.Text = ((ToolStripMenuItem)sender).Text.Replace("&", string.Empty);
            foreach (ToolStripItem item in Dimensions.DropDownItems) if (item is ToolStripMenuItem) 
                    ((ToolStripMenuItem)item).Checked = item == sender;
        }

        private void Pixels_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (ToolStripItem item in Dimensions.DropDownItems) if (item is ToolStripMenuItem)
                    ((ToolStripMenuItem)item).Checked = Pixels.Text == ((ToolStripMenuItem)item).Text.Replace("&", string.Empty);
        }

        private void Pixels_GotFocus(object sender, EventArgs e)
        {
            size = Pixels.Text;
        }

        private void Pixels_LostFocus(object sender, EventArgs e)
        {
            if (Regex.IsMatch(Pixels.Text.ToLowerInvariant().Trim(), @"^\d{2,3}\s*(x)?\s*(\d{2,3})?$"))
            {
                var dimensions = Parse(Pixels.Text.ToLowerInvariant().Trim());
                if (Array.TrueForAll(dimensions, dimension => dimension >= 16 && dimension <= 128) &&
                    Array.TrueForAll(dimensions, dimension => dimension == dimensions[0]))
                    if (dimensions.Length == 1) size = dimensions[0] + " x " + dimensions[0];
                    else if (dimensions.Length == 2) size = dimensions[0] + " x " + dimensions[1];
            }
            Pixels.Text = size;
            Pixels_SelectedIndexChanged(sender, e);
        }

        private void Pixels_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) Pixels_LostFocus(sender, e);
        }

        private void Export_Click(object sender, EventArgs e)
        {
            if (Icons.SelectedItems.Count > 0 && SaveTo.ShowDialog() == DialogResult.OK)
            {   // Collect all checked items from the Export menu and parse into an array of width x length sizes.
                var sizes = Array.ConvertAll(Choices(ExportAs.DropDownItems),
                    item => Parse(item.Text.Replace("&", string.Empty).ToLowerInvariant().Trim()));
                foreach (ListViewItem selection in Icons.SelectedItems)
                {
                    SingleIcon icon = (new MultiIcon()).Add(selection.Text);
                    foreach (var size in sizes) icon.Add(imageMso[selection.Text, size[0], size[1]]);
                    icon.Save(Path.Combine(SaveTo.SelectedPath, string.Format("{0}.{1}", selection.Text, "ico")));
                    DestroyIcon(icon.Icon.Handle);
                }
                MessageBox.Show("Export Complete.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Size_Check(object sender, EventArgs e)
        {
            ((ToolStripMenuItem)sender).Checked ^= true;
        }

        private void Save_Click(object sender, EventArgs e)
        {
            if (Icons.SelectedItems.Count > 0 && SaveTo.ShowDialog() == DialogResult.OK)
            {   // Collect all checked items from the Save menu and parse into an array of image formats.
                var formats = Array.ConvertAll(Choices(SaveAs.DropDownItems), item => Format(item.Name));
                var size = Parse(Pixels.Text.ToLowerInvariant().Trim());
                foreach (ListViewItem selection in Icons.SelectedItems)
                    using (var image = imageMso[selection.Text, size[0], size[1]])
                    using (var canvas = new Bitmap(size[0], size[1]))
                    using (var drawing = Graphics.FromImage(canvas))
                    {   // Image is drawn on a colored opaque canvas for formats that do not support transparency.
                        drawing.Clear(Palette.Color);
                        drawing.DrawImage(image, 0, 0);
                        foreach (ImageFormat format in formats)
                        {
                            string filename = Path.Combine(SaveTo.SelectedPath,
                                string.Format("{0}.{1}", selection.Text, Extension(format)));
                            if (format == ImageFormat.Icon)
                            {   // image.Save(filename, format); // Loses color depth. Use IconLib instead.
                                SingleIcon icon = (new MultiIcon()).Add(selection.Text);
                                icon.Add(image);
                                icon.Save(filename);
                                DestroyIcon(icon.Icon.Handle);
                            }
                            else if (format == ImageFormat.Bmp || format == ImageFormat.Gif || format == ImageFormat.Jpeg)
                                canvas.Save(filename, format); // These formats do not support alpha channel transparency.
                            else image.Save(filename, format);
                        }
                    }
                MessageBox.Show("Save Complete.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Format_Check(object sender, EventArgs e)
        {
            ((ToolStripMenuItem)sender).Checked ^= true;
        }

        private void Background_Click(object sender, EventArgs e)
        {
            string text = Background.Text.Replace(Palette.Color.Name, "{0}");
            Palette.ShowDialog();
            Background.Text = string.Format(text, Palette.Color.Name);
        }

        private void Support_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("This action will launch the web browser and open the support website.", "Help",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK) Process.Start(imageMso.Help);
        }

        private void Information_Click(object sender, EventArgs e)
        {
            (new About()).ShowDialog();
        }

        ///<summary>Collect all checked items from a <see cref="System.Windows.Forms.ToolStripItemCollection"/>
        ///and into a <see cref="System.Windows.Forms.ToolStripMenuItem"/> array.</summary>
        ///<param name="items"><see cref="System.Windows.Forms.ToolStripItemCollection"/></param>
        ///<returns><see cref="System.Array"/> of <see cref="System.Windows.Forms.ToolStripMenuItem"/></returns>
        private static ToolStripItem[] Choices(ToolStripItemCollection items)
        {
            var choices = new ToolStripItem[items.Count];
            items.CopyTo(choices, 0);
            return Array.FindAll(choices, item => item is ToolStripMenuItem && ((ToolStripMenuItem)item).Checked);
        }

        ///<summary>Translates the <see cref="System.Drawing.Imaging.ImageFormat"/> 
        ///constant into a 3 or 4 letter image file extension.</summary>
        ///<param name="format"><see cref="System.Drawing.Imaging.ImageFormat"/> constant.</param>
        ///<returns>3 or 4 letter <see cref="System.String"/> file extension.</returns>
        private string Extension(ImageFormat format)
        {
            switch (format.ToString())
            {
                case "Bmp": return "bmp";
                case "Emf": return "emf";
                case "Gif": return "gif";
                case "Icon": return "ico";
                case "Jpeg": return "jpg";
                case "Exif": return "exif";
                case "Png": return "png";
                case "Tiff": return "tif";
                case "Wmf": return "wmf";
                default: return string.Empty;
            }
        }

        ///<summary>Translates the 3 or 4 letter image file extension into a 
        ///<see cref="System.Drawing.Imaging.ImageFormat"/> constant.</summary>
        ///<param name="extension">3 or 4 letter <see cref="System.String"/> file extension.</param>
        ///<returns><see cref="System.Drawing.Imaging.ImageFormat"/> constant.</returns>
        private ImageFormat Format(string extension)
        {
            switch (extension.ToLowerInvariant())
            {
                case "bmp": return ImageFormat.Bmp;
                case "emf": return ImageFormat.Emf;
                case "gif": return ImageFormat.Gif;
                case "ico": return ImageFormat.Icon;
                case "jpeg": return ImageFormat.Jpeg;
                case "jpg": return ImageFormat.Jpeg;
                case "exif": return ImageFormat.Exif;
                case "png": return ImageFormat.Png;
                case "tif": return ImageFormat.Tiff;
                case "tiff": return ImageFormat.Tiff;
                case "wmf": return ImageFormat.Wmf;
                default: return null;
            }
        }

        ///<summary>Parses n dimension sizes from a <see cref="System.String"/> in the form 
        ///"D1 x D2 x … x Dn" to a <see cref="System.Array"/> of <see cref="System.Int16" />.</summary>
        ///<param name="pixels"><see cref="System.String"/> in the form "D1 x D2 x … x Dn".</param>
        ///<returns><see cref="System.Array"/> of <see cref="System.Int16" /> in the form 
        ///{ D1, D2, … , Dn }</returns>
        private short[] Parse(string pixels)
        {
            return Array.ConvertAll(pixels.Split("x".ToCharArray(),
                StringSplitOptions.RemoveEmptyEntries), s => Int16.Parse(s.Trim()));
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private extern static bool DestroyIcon(IntPtr handle);
    }
}