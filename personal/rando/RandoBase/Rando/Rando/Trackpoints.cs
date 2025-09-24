using Microsoft.VisualBasic.Logging;
using System;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;


namespace Rando
{
    // Class representing a single GPS track point
    public class Trackpoint
    {
        public double Latitude { get; set; }   // Latitude in decimal degrees
        public double Longitude { get; set; }  // Longitude in decimal degrees
        public double Elevation { get; set; }  // Elevation in meters
    }


    // Class to read GPX files and return a list of Trackpoints
    public class GPXReader
    {
        // Reads a GPX file and parses all track points
        public static List<Trackpoint> ReadGPX(string filePath)
        {
            List<Trackpoint> points = new List<Trackpoint>();

            XmlDocument doc = new XmlDocument();
            doc.Load(filePath); // Load GPX XML file

            XmlNodeList trkpts = doc.GetElementsByTagName("trkpt"); // Find all <trkpt> nodes

            foreach (XmlNode node in trkpts)
            {//chatgpt help to make syntazx and correctly Parse data

                // Parse latitude and longitude from attributes
                double lat = double.Parse(node.Attributes["lat"].Value, System.Globalization.CultureInfo.InvariantCulture);
                double lon = double.Parse(node.Attributes["lon"].Value, System.Globalization.CultureInfo.InvariantCulture);

                // Parse elevation from child <ele> element
                double ele = double.Parse(node["ele"].InnerText, System.Globalization.CultureInfo.InvariantCulture);

                // Add a new Trackpoint to the list
                points.Add(new Trackpoint
                {
                    Latitude = lat,
                    Longitude = lon,
                    Elevation = ele
                });
            }

            return points; // Return all parsed track points
        }
    }


    public class GPXWriter
    {
        public static void WriteGPX(string filePath, List<Trackpoint> points)
        {
            // creating XML structure
            XNamespace ns = "http://www.topografix.com/GPX/1/1";

            var gpx = new XElement(ns + "gpx",
                new XAttribute("version", "1.1"),
                 new XElement(ns + "trk",
                    new XElement(ns + "trkseg",
                    // Convert each Trackpoint to a <trkpt>
                    points.ConvertAll(p =>
                    new XElement(ns + "trkpt",
                     new XAttribute("lat", p.Latitude.ToString(CultureInfo.InvariantCulture)),
                                new XAttribute("lon", p.Longitude.ToString(CultureInfo.InvariantCulture)),
                                new XElement(ns + "ele", p.Elevation.ToString(CultureInfo.InvariantCulture))
                       )
                    )
                 )
            )
        );

            gpx.Save(filePath);
        }
    }

 
  
}