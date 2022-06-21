namespace MilwaukeeSite02.Models
{
    public class ProcessControlModel
    {
      

        public string Model { get; set; }
        public string Line { get; set; }

        public DateTime Date { get; set; }

        public DateTime Hora { get; set; }
        public string Tecnico { get; set; }
        public string Item { get; set; }

        public string ProcessControlPoint { get; set; }
        public string Equipment { get; set; }
        public string Instrument { get; set; }
        public string Values { get; set; }

        public string Unit { get; set; }

        public int[] TestData { get; set; }
    }
}
