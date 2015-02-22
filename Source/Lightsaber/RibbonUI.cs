using NetOffice.OfficeApi;
using NetOffice.Tools;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = NetOffice.PowerPointApi;
using Resources = Lightsaber.Resources;
using stdole;

namespace Lightsaber
{
    [CustomUI("Lightsaber.RibbonUI.xml")]
    public partial class Addin
    {
        private const int galleryItemSize = 26;
        private Color borderColor = Color.FromArgb(226, 228, 231);
        private Color[,] galleryColorTable = {
            {Color.Yellow, Color.Lime, Color.Cyan, Color.Magenta, Color.Blue},
            {Color.Red, Color.Navy, Color.Teal, Color.Green, Color.Purple},
            {Color.Maroon, Color.Olive, Color.Gray, Color.Silver, Color.Black}
        };

        internal IRibbonUI RibbonUI { get; private set; }

        private Color previoiusColor = Color.Transparent;
        private Color currentColor = Color.Yellow;

        private void Application_WindowSelectionChangeEvent(PowerPoint.Selection Sel)
        {
            this.RibbonUI.InvalidateControl("HighlightButton");
            this.RibbonUI.InvalidateControl("ClearButton");
        }

        private Color GetGalleryColor(int index)
        {
            int columnCount = galleryColorTable.GetLength(1);
            return galleryColorTable[index / columnCount, index % columnCount];
        }

        private void SetColor(Color color)
        {
            currentColor = color;
            if (currentColor != previoiusColor)
            {
                previoiusColor = currentColor;
                this.RibbonUI.InvalidateControl("HighlightButton");
            }
            Highlight();
        }

        private void Highlight()
        {
            Highlight(currentColor);
        }

        #region attribute callbacks

        public void OnLoadRibonUI(IRibbonUI ribbonUI)
        {
            this.RibbonUI = ribbonUI;
            this.Application.WindowSelectionChangeEvent += Application_WindowSelectionChangeEvent;
        }

        public bool GetEnabled(IRibbonControl control)
        {
            return HighlightEnabled;
        }

        public IPictureDisp GetImage(IRibbonControl control)
        {
            Bitmap image = null;

            switch (control.Id)
            {
                case "ColorGallery":
                    image = Resources.Img.color_swatch;
                    break;
                case "MoreColorButton":
                    image = Resources.Img.color;
                    break;
                case "ClearButton":
                    image = Resources.Img.eraser;
                    break;
                default:
                    return null;
            }

            return PictureConverter.ImageToPictureDisp(image);
        }

        public string GetLabel(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "ColorGallery":
                    return Resources.Str.ColorGalleryLabel;
                case "MoreColorButton":
                    return Resources.Str.MoreColorButtonLabel;
                case "ClearButton":
                    return Resources.Str.ClearButtonLabel;
                default:
                    return null;
            }
        }

        public IPictureDisp HighlightButton_GetImage(IRibbonControl control)
        {
            using (var image = Resources.Img.highlighter_color)
            using (var graphics = Graphics.FromImage(image))
            using (var colorBrush = new SolidBrush(currentColor))
            {
                graphics.FillRectangle(colorBrush, 0, 12, 16, 4);
                return PictureConverter.ImageToPictureDisp(image);
            }
        }

        public void HighlightButton_OnAction(IRibbonControl control)
        {
            Highlight();
        }

        public int ColorGallery_GetItemCount(IRibbonControl control)
        {
            return galleryColorTable.Length;
        }

        public IPictureDisp ColorGallery_GetItemImage(IRibbonControl control, int index)
        {
            using(var image = new Bitmap(galleryItemSize, galleryItemSize))
            using (var graphics = Graphics.FromImage(image))
            using (var borderBrush = new SolidBrush(borderColor))
            using (var colorBrush = new SolidBrush(GetGalleryColor(index)))
            {
                graphics.FillRectangle(borderBrush, 0, 0, galleryItemSize, galleryItemSize);
                graphics.FillRectangle(colorBrush, 1, 1, galleryItemSize - 2, galleryItemSize - 2);
                return PictureConverter.ImageToPictureDisp(image);
            }
        }

        public int ColorGallery_GetItemSize(IRibbonControl control)
        {
            return galleryItemSize;
        }

        public void ColorGallery_OnAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            Color color = GetGalleryColor(selectedIndex);
            SetColor(color);
        }

        public void MoreColorButton_OnAction(IRibbonControl control)
        {
            using (var colorDialog = new ColorDialog())
            {
                colorDialog.Color = currentColor;
                if (colorDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                SetColor(colorDialog.Color);
            }
        }

        public void ClearButton_OnAction(IRibbonControl control)
        {
            ClearHighlight();
        }

        #endregion
    }

    internal class PictureConverter : AxHost
    {
        private PictureConverter() : base("") { }

        static public IPictureDisp ImageToPictureDisp(Image image)
        {
            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}

