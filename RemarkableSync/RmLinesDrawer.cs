﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Threading.Tasks;
using RemarkableSync.RmLine;


namespace RemarkableSync
{
    public class RmLinesDrawer
    {
        static public List<Bitmap> DrawPages(List<RmPage> pages)
        {
            List<Bitmap> images = (from page in pages
                                  select DrawPage(page)).ToList();
            return images;
        }

        static public Bitmap DrawPage(RmPage page)
        {
            Bitmap image = new Bitmap(RmConstants.X_MAX, RmConstants.Y_MAX);

            Graphics graphics = Graphics.FromImage(image);
            graphics.SmoothingMode = SmoothingMode.AntiAlias;

            foreach (RmLayer layer in page.Objects)
            {
                DrawLayer(layer, ref graphics);
            }

            return image;
        }

        static private void DrawLayer(RmLayer layer, ref Graphics graphics)
        {
            foreach (RmStroke stroke in layer.Objects)
            {
                if (stroke.IsVisible())
                {
                    DrawStroke(stroke, ref graphics, layer.LayerColor);
                }
            }
        }

        static private void DrawStroke(RmStroke stroke, ref Graphics graphics, Color? color = null)
        {
            if (color == null)
            {
                switch (stroke.Colour)
                {
                    case ColourEnum.GREY:
                        color = Color.Gray;
                        break;
                    case ColourEnum.WHITE:
                        color = Color.White;
                        break;
                    case ColourEnum.BLACK:
                    default:
                        color = Color.Black;
                        break;
                }
            }

            Pen pen = new Pen(color.GetValueOrDefault(Color.Black), stroke.Width);

            GraphicsPath path = new GraphicsPath();
            Point[] points = new Point[stroke.Objects.Count];
            for (int i = 0; i < stroke.Objects.Count; ++i)
            {
                RmSegment segment = (RmSegment) stroke.Objects[i];
                points[i] = new Point((int)segment.X, (int)segment.Y);
            }
            path.AddLines(points);
            graphics.DrawPath(pen, path);

            pen.Dispose();
        }

    }
}
