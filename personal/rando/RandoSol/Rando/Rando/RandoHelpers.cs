using Gpx;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace Rando
{
    public static class RandoHelpers
    {
        public static double TrackLengthByAnonymousAggregate(List<GpxTrackPoint> trackPoints)
        {
            // Calcul de la longueur
            return trackPoints.
                Skip(1).
                Select(trackPts => new { Lon = trackPts.Longitude, Lat = trackPts.Latitude, Alt = trackPts.Elevation ?? 0, Dst = 0.0 }).
                Aggregate(
                    new { Lon = (double)trackPoints[0].Longitude, Lat = (double)trackPoints[0].Latitude, Alt = (double)trackPoints[0].Elevation, Dst = (double)0.0 },
                    (a, b) => new { Lon = b.Lon, Lat = b.Lat, Alt = b.Alt, Dst = a.Dst + new GpxTrackPoint() { Longitude = b.Lon, Latitude = b.Lat, Elevation = b.Alt }.GetDistanceFrom(new GpxTrackPoint() { Longitude = a.Lon, Latitude = a.Lat, Elevation = a.Alt }) }
                ).
                Dst;
        }

        public static double TrackLengthByAggregatingGpxTrackPoints(List<GpxTrackPoint> trackPoints)
        {
            // Calcul de la longueur
            return trackPoints.
                Aggregate(
                    trackPoints.First(),
                    (a, b) => new GpxTrackPoint { Longitude = b.Longitude, Latitude = b.Latitude, Elevation = b.Elevation, Course = a.Course + a.GetDistanceFrom(b) }
                ).
                Course ?? 0;
        }

        public static double TrackLengthByZip(List<GpxTrackPoint> trackPoints)
        {
            return trackPoints.Zip(trackPoints.Skip(1), (a, b) => a.GetDistanceFrom(b)).Sum();
        }

        public static double TrackLengthRecursive(List<GpxTrackPoint> track)
        {
            if (track.Count == 2)
            {
                return track[0].GetDistanceFrom(track[1]);
            }
            else
            {
                int middle = track.Count / 2;

                return TrackLengthRecursive(track.Take(middle + 1).ToList()) + TrackLengthRecursive(track.Skip(middle).ToList());
            }
        }


    }
}
