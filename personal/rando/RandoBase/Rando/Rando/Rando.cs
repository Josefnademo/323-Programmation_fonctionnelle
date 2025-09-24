namespace Rando
{
    public partial class Rando : Form
    {
        private List<Point> points; //points to draw
        public Rando()
        {
            InitializeComponent();

            // Read GPX and convert to points for the shape
            var trackpoints = GPXReader.ReadGPX(); // путь по умолчанию
            points = TrackConverter.ToPoints(trackpoints, this.ClientSize.Width, this.ClientSize.Height);

            // Redraw the form
            this.Invalidate();
        }

        private void Rando_Form_Paint(object sender, PaintEventArgs e)
        {
            if (points == null || points.Count < 2)
                return;

            using (Pen myPen = new Pen(Color.Red, 2))
            {
                e.Graphics.DrawLines(myPen, points.ToArray());
            }
        }

        private void Rando_Load(object sender, EventArgs e)
        {

        }
    }
}
