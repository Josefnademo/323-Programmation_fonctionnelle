using Gpx;
using System.Diagnostics;
using System.Numerics;

namespace Rando
{
    public partial class Rando : Form
    {
        private const double LAT_OFFSET = 46.39;
        private const double LONG_OFFSET = 7.61;

        private List<GpxTrackPoint> trackPoints;

        public Rando()
        {
            InitializeComponent();
            using (GpxReader reader = new GpxReader(File.Open("gemmikandersteg.gpx", FileMode.Open)))
            {
                while (reader.Read())
                {
                    switch (reader.ObjectType)
                    {
                        case GpxObjectType.Metadata:
                            Console.WriteLine(reader.Metadata);
                            break;
                        case GpxObjectType.WayPoint:
                            Console.WriteLine(reader.WayPoint);
                            break;
                        case GpxObjectType.Route:
                            Console.WriteLine(reader.Route);
                            break;
                        case GpxObjectType.Track:
                            Console.WriteLine(reader.Track);
                            trackPoints = reader.Track.Segments[0].TrackPoints.ToList();
                            break;
                    }
                }
            }

            double zip = RandoHelpers.TrackLengthByZip(trackPoints);
            double anonymous = RandoHelpers.TrackLengthByAnonymousAggregate(trackPoints);
            double gpxtp = RandoHelpers.TrackLengthByAggregatingGpxTrackPoints(trackPoints);
            double recu = RandoHelpers.TrackLengthRecursive(trackPoints);
        }

        private void Rando_Form_Paint(object sender, PaintEventArgs e)
        {
            Pen myPen = new Pen(Color.Red);
            myPen.Width = 2;

            Point[] points = trackPoints.Select(tp => new Point((int)Math.Round((tp.Longitude - LONG_OFFSET) * 10000), this.ClientRectangle.Height - (int)Math.Round((tp.Latitude - LAT_OFFSET) * 10000))).ToArray();
            this.CreateGraphics().DrawLines(myPen, points);
        }
    }
}
