 namespace MilwaukeeSite02.Models
{
    public class ChangeOverModel
    {
        public string Model { get; set; }
        public string Line { get; set; }

        public DateTime Date { get; set; }

        public string Tecnico { get; set; }
        public string Item { get; set; }

        public string ProcessControlPoint { get; set; }
        public string Equipment { get; set; }
        public string Instrument { get; set; }
        public string ControlPoint { get; set; }

        public string ControlSPEC { get; set; }

        public string Unit { get; set; }
        public string Tester { get; set; }
    }
}
