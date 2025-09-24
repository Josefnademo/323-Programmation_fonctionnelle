using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Drawing;

namespace Rando
{ // Class to assign colors to track points based on elevation
    internal class TrackColorizer
    {
        static Color[] gradient = new Color[]
    {// Gradient from green (low elevation) to red (high elevation)
        Color.FromArgb(255, 144, 238, 144), // Vert clair
        Color.FromArgb(255, 162, 216, 128),
        Color.FromArgb(255, 180, 194, 112),
        Color.FromArgb(255, 198, 172, 96),
        Color.FromArgb(255, 216, 150, 80),
        Color.FromArgb(255, 234, 128, 64),
        Color.FromArgb(255, 244, 106, 48),
        Color.FromArgb(255, 248,  84, 36),
        Color.FromArgb(255, 252,  62, 24),
        Color.FromArgb(255, 254,  48, 18),
        Color.FromArgb(255, 255,  32, 12),
        Color.FromArgb(255, 255,  16,  6),
        Color.FromArgb(255, 255,   0,  0)  // Rouge vif
    };


        // Convert Trackpoints to screen points with associated color
        public static List<(Point point, Color color)> ToColoredPoints(List<Trackpoint> trackpoints, int width, int height)
        {
            var points = TrackConverter.ToPoints(trackpoints, width, height);
            var result = new List<(Point, Color)>();

            for (int i = 0; i < trackpoints.Count; i++)
            {  // Choose color based on elevation (divide by 100 to map to gradient index)
                int index = Math.Min((int)(trackpoints[i].Elevation / 100), gradient.Length - 1);
                Color c = gradient[index];
                result.Add((points[i], c));
            }

            return result;
        }
    }
}