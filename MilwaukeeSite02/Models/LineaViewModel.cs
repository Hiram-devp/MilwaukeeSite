namespace MilwaukeeSite02.Models
{
    public class LineaViewModel
    {
        public LineaViewModel()
        {
            LineasList = new List<Itemlist?>()
            {
                new Itemlist { Text = "MXC001", Value = 1},
                new Itemlist { Text = "MXC002", Value = 2},
            };

            //ModelosList = new List<Itemlist>()
            //{
            //    new Itemlist { Text = "2802", Value=1},
            //};
        }
        public int LineaId { get; set; }

        public List<Itemlist?> LineasList { get; set; }

        public int ModeloId { get; set; }

        public List<Itemlist> ModelosList { get; set; }
    }
}

