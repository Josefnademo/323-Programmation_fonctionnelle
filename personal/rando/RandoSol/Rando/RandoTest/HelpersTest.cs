using Gpx;
using Rando;

namespace RandoTest
{
    [TestClass]
    public class HelpersTest
    {
        [TestMethod]
        public void Track_length_test()
        {
            // Arrange:
            // Build a list of a few trackpoints (random)
            // Deltas between consecutive points are always 12, 5 and 0, which means a length of 13 (12^2+5^2=13^2)
            // The length of the track is thus (n-1)*13
            int nbPoints = new Random().Next(20, 50);

            List<GpxTrackPoint> trackPoints = new List<GpxTrackPoint>();
            int lat = 10;
            int lon = 100;
            int alt = 1000;
            for (int i = 0; i < nbPoints; i++)
            {
                GpxTrackPoint tp = new GpxTrackPoint();
                tp.Longitude = lon;
                tp.Latitude = lat;
                tp.Elevation = alt;
                trackPoints.Add(tp);
                switch (i % 3)
                {
                    case 0:
                        lon += 12;
                        lat += 5;
                        break;
                    case 1:
                        lat += 12;
                        alt += 5;
                        break;
                    case 2:
                        alt += 12;
                        lon += 5;
                        break;
                }
            }

            // Act
            double tracklengthByAggregate = Math.Round(RandoHelpers.TrackLengthByAnonymousAggregate(trackPoints), 2);
            double tracklengthByRecursion = Math.Round(RandoHelpers.TrackLengthRecursive(trackPoints), 2);
            double tracklengthByGpxAggregate = Math.Round(RandoHelpers.TrackLengthByAggregatingGpxTrackPoints(trackPoints), 2);
            double tracklengthByZip = Math.Round(RandoHelpers.TrackLengthByZip(trackPoints), 2);

            // Assert
            Assert.AreEqual(13 * (nbPoints - 1)/100.0, tracklengthByRecursion,"Recursion Ko");
            Assert.AreEqual(13 * (nbPoints - 1)/100.0, tracklengthByAggregate, "Aggregate Ko");
            Assert.AreEqual(13 * (nbPoints - 1)/100.0, tracklengthByZip, "Aggregate Zip Ko");
            Assert.AreEqual(13 * (nbPoints - 1)/100.0, tracklengthByGpxAggregate, "Aggregate Gpx Ko");
        }
    }
}