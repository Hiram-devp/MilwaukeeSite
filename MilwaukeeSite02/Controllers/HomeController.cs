using ClosedXML.Excel;
using Fingers10.ExcelExport.ActionResults;
using Microsoft.AspNetCore.Mvc;
using MilwaukeeSite02.Models;
using System.Diagnostics;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;

namespace MilwaukeeSite02.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly  List<ChangeOverModel> cho = new List<ChangeOverModel>();
        private readonly  List<TestDataModel> testData = new List<TestDataModel>();
        private readonly List<ProcessControlModel> pc = new List<ProcessControlModel>();
        private string? itemId;
        public HomeController(ILogger<HomeController> logger,  List<ChangeOverModel> cho, List<TestDataModel> testData, List<ProcessControlModel> pc)
        {
            _logger = logger;
            this.cho = cho;
            this.testData = testData;
            this.pc = pc;
            
            
        }
       
       
       
        
        public IActionResult Index()
        {
            
            LineaViewModel model = new LineaViewModel();

                model.ModelosList = new List<Itemlist>
                {
                    new Itemlist {Text ="", Value=1},
                };

            var lista = new List<string>();
            foreach (var x in model.LineasList)
            {
                lista.Add(x.Text);
            }
            var listaModelo = new List<string>();
            foreach (var x in model.ModelosList)
            {
                listaModelo.Add(x.Text);
            }
            ViewBag.Linea = lista;
            
            return View(model);
        }

        public IActionResult OnLineaChange(string type)
        {
            var model = new LineaViewModel();
            if (type == "MXC001")
            {
                model.ModelosList = new List<Itemlist>
                { new Itemlist{Text = "2851", Value = 2851 },
                  new Itemlist{Text = "2850", Value= 2850},
                };
            }
            else if (type == "MXC002")
            {
                model.ModelosList = new List<Itemlist>
                {
                    new Itemlist{Text = "2801", Value = 2801},
                    new Itemlist{Text ="2802", Value=2802},
                    new Itemlist{Text= "2902", Value=2902}
                };
            }
            
           
            return PartialView("FormPartialView", model);
        }

       
        public IActionResult _ProcessControlPartialView(string type, string model)
        {
           
                
                if ((type == "MXC001") && (model == "2851"))
                {
                    #region data
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "1",
                        ProcessControlPoint = "B4",
                        Equipment = "Grease Dispenser",
                        Instrument = "Electronic Scale",
                        Values = "2.0+-0.2",
                        Unit = "g"

                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "3",
                        ProcessControlPoint = "B2",
                        Equipment = "Grease Dispenser",
                        Instrument = "Electronic Scale",
                        Values = "1.5+-0.2",
                        Unit = "g"
                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "4",
                        ProcessControlPoint = "A1",
                        Equipment = "Grease Dispenser",
                        Instrument = "Electronic Scale",
                        Values = "0.5+-0.2",
                        Unit = "g"
                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "9",
                        ProcessControlPoint = "D2",
                        Equipment = "Electric Screw Driver",
                        Instrument = "Torque tester",
                        Values = "4+-1",
                        Unit = "kgf.cm"
                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "10",
                        ProcessControlPoint = "D1",
                        Equipment = "Soldering iron",
                        Instrument = "Infrared thermometer",
                        Values = "380-410",
                        Unit = "℃"
                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "11",
                        ProcessControlPoint = "E10",
                        Equipment = "Electric Screw Driver",
                        Instrument = "Torque tester",
                        Values = "10+-1",
                        Unit = "kgf.cm"
                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "12",
                        ProcessControlPoint = "D10",
                        Equipment = "Electric Screw Driver",
                        Instrument = "Torque tester",
                        Values = "10+-1",
                        Unit = "kgf.cm"
                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "13",
                        ProcessControlPoint = "E12",
                        Equipment = "Electric Screw Driver",
                        Instrument = "Torque tester",
                        Values = "10-14",
                        Unit = "kgf.cm"
                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "14",
                        ProcessControlPoint = "Work Bench",
                        Equipment = "Leakage detection",
                        Instrument = "Multimeter",
                        Values = "<=2",
                        Unit = "v"
                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "15",
                        ProcessControlPoint = "Running booth",
                        Equipment = "Voltage",
                        Instrument = "Multimeter",
                        Values = "20.5",
                        Unit = "v"
                    });
                    pc.Add(new ProcessControlModel()
                    {
                        Model = model,
                        Line = type,
                        Date = DateTime.Now,
                        Tecnico = String.Empty,
                        Item = "16",
                        ProcessControlPoint = "Function tester",
                        Equipment = "Current test|Speed low|Speed high",
                        Instrument = "Visual Inspection",
                        Values = "5.50|1360-2040|3060-3800",
                        Unit = "A|Rpm|Rpm"
                    });

                    #endregion
                }      
                else if ((type == "MXC001") && (model == "2850"))
                {
                     #region data
                pc.Add(new ProcessControlModel()
                    {
                      Model = model,
                      Line = type,
                      Date = DateTime.Now,
                      Tecnico = String.Empty,
                      Item = "1",
                      ProcessControlPoint = "B4",
                       Equipment = "Grease Dispenser",
                       Instrument = "Electronic Scale",
                      Values = "2.0+-0.2",
                       Unit = "g"

                     });
                pc.Add(new ProcessControlModel()
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "3",
                    ProcessControlPoint = "B2",
                    Equipment = "Grease Dispenser",
                    Instrument = "Electronic Scale",
                    Values = "1.5+-0.2",
                    Unit = "g"
                });
                pc.Add(new ProcessControlModel()
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "4",
                    ProcessControlPoint = "A1",
                    Equipment = "Grease Dispenser",
                    Instrument = "Electronic Scale",
                    Values = "0.5+-0.2",
                    Unit = "g"
                });
                pc.Add(new ProcessControlModel()
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "9",
                    ProcessControlPoint = "D2",
                    Equipment = "Electric Screw Driver",
                    Instrument = "Torque tester",
                    Values = "4+-1",
                    Unit = "kgf.cm"
                });
                pc.Add(new ProcessControlModel()
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "10",
                    ProcessControlPoint = "D1",
                    Equipment = "Soldering iron",
                    Instrument = "Infrared thermometer",
                    Values = "380-410",
                    Unit = "℃"
                });
                pc.Add(new ProcessControlModel()
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "11",
                    ProcessControlPoint = "E10",
                    Equipment = "Electric Screw Driver",
                    Instrument = "Torque tester",
                    Values = "10+-1",
                    Unit = "kgf.cm"
                });
                pc.Add(new ProcessControlModel()
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "12",
                    ProcessControlPoint = "E12",
                    Equipment = "Electric Screw Driver",
                    Instrument = "Torque tester",
                    Values = "10+-1",
                    Unit = "kgf.cm"
                }); 
                pc.Add(new ProcessControlModel()
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "13",
                    ProcessControlPoint = "Work Bench",
                    Equipment = "Leakage detection",
                    Instrument = "Multimeter",
                    Values = "<=2",
                    Unit = "v"
                });
                pc.Add(new ProcessControlModel()
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "14",
                    ProcessControlPoint = "Running booth",
                    Equipment = "Voltage",
                    Instrument = "Multimeter",
                    Values = "20.5",
                    Unit = "v"
                });
                pc.Add(new ProcessControlModel()
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "16",
                    ProcessControlPoint = "Function tester",
                    Equipment = "Current test|Speed test",
                    Instrument = "Visual Inspection",
                    Values = "5.50|3060-3800",
                    Unit = "Rpm"
                });
                #endregion
                }
                else if ((type == "MXC002") && (model == "2801"))
                {
                     #region data
                #endregion
                }
                else if ((type == "MXC002") && (model == "2802"))
                {
                     #region data
                #endregion
                }
                else if ((type == "MXC002") && (model == "2902"))
                {
                     #region data
                     #endregion
                }
                return PartialView("_ProcessControlPartialView02", pc);
        }

        public IActionResult _ChangeOverPartialView(string type, string model)
        {

            if (cho.Count > 0)
                cho.Clear();
            if ((type == "MXC001") && (model == "2851"))
            {
                #region data
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "1",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "2",
                    ProcessControlPoint = "B2",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "1.5+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "3",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "5",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "6",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "8",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "10",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "11",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "12",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "13",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "14",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                #endregion
            }
            else if ((type == "MXC001") && (model == "2850"))
            {
                #region data
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "1",
                    ProcessControlPoint = "B4",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "2.0+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "2",
                    ProcessControlPoint = "B2",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "1.5+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "3",
                    ProcessControlPoint = "B1",
                    Equipment = "Air Press",
                    Instrument = String.Empty,
                    ControlPoint = "Press Force",
                    ControlSPEC = "250-350",
                    Unit = "kgf",
                    Tester = "Press force measurement set"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "5",
                    ProcessControlPoint = "A3",
                    Equipment = "Grease Dispenser",
                    Instrument = String.Empty,
                    ControlPoint = "Weight",
                    ControlSPEC = "0.5+-0.2",
                    Unit = "g",
                    Tester = "Electronic scale"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "6",
                    ProcessControlPoint = "B6",
                    Equipment = "Air Press",
                    Instrument = String.Empty,
                    ControlPoint = "Press Force",
                    ControlSPEC = "150-250",
                    Unit = "kgf",
                    Tester = "Press force measurement set"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "8",
                    ProcessControlPoint = "C3",
                    Equipment = "Air Press",
                    Instrument = String.Empty,
                    ControlPoint = "Press Force",
                    ControlSPEC = "300-400",
                    Unit = "kgf",
                    Tester = "Press force measurement set"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "10",
                    ProcessControlPoint = "D2",
                    Equipment = "Electric screw driver",
                    Instrument = String.Empty,
                    ControlPoint = "Torque",
                    ControlSPEC = "4+-1",
                    Unit = "kgf.cm",
                    Tester = "Torque Tester"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "11",
                    ProcessControlPoint = "D1",
                    Equipment = "Soldering Iron",
                    Instrument = String.Empty,
                    ControlPoint = "Temperature",
                    ControlSPEC = "380-410",
                    Unit = "℃",
                    Tester = "Infrared thermometer"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "12",
                    ProcessControlPoint = "E10",
                    Equipment = "Electric screw driver",
                    Instrument = String.Empty,
                    ControlPoint = "Torque",
                    ControlSPEC = "10+-1",
                    Unit = "kgf.cm",
                    Tester = "Torque Tester"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "13",
                    ProcessControlPoint = "E10",
                    Equipment = "Electric screw driver",
                    Instrument = String.Empty,
                    ControlPoint = "Torque",
                    ControlSPEC = "10+-1",
                    Unit = "kgf.cm",
                    Tester = "Toque Tester"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "14",
                    ProcessControlPoint = "E12",
                    Equipment = "Electric screw driver",
                    Instrument = String.Empty,
                    ControlPoint = "Torque",
                    ControlSPEC = "10-14",
                    Unit = "kgf.cm",
                    Tester = "Torque Tester"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "Tester",
                    ProcessControlPoint = "Tester",
                    Equipment = "Running Booth",
                    Instrument = String.Empty,
                    ControlPoint = "Voltaje",
                    ControlSPEC = "20.5",
                    Unit = "V",
                    Tester = "Visual Inspection"
                });
                cho.Add(new ChangeOverModel
                {
                    Model = model,
                    Line = type,
                    Date = DateTime.Now,
                    Tecnico = String.Empty,
                    Item = "EOL",
                    ProcessControlPoint = "Tester",
                    Equipment = "EOL",
                    Instrument = String.Empty,
                    ControlPoint = "Voltaje",
                    ControlSPEC = "20.5",
                    Unit = "V",
                    Tester = "Visual Inspection"
                });
                #endregion
            }
            else if ((type == "MXC002") && (model == "2801"))
            {
                #region data
                #endregion
            }
            else if ((type == "MXC002") && (model == "2802"))
            {
                #region data
                #endregion
            }
            else if ((type == "MXC002") && (model == "2902"))
            {
                #region data
                #endregion
            }
            return PartialView("_ChangeOverPartialView", cho);
        }

        public IActionResult Details(string itemId)
        {
            ViewBag.itemId = itemId.Trim();
            return PartialView("_TestDataModalPartialView", testData);
        }
        public IActionResult ExportExcel(string txtTecnico)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheets = workbook.Worksheets.Add("Items");
                var currentrow = 3;
             
                worksheets.Cell(2, 1).Value = "Modelo:";              
                worksheets.Cell(2, 3).Value = "Linea:";
                worksheets.Cell(2, 5).Value = "Date:";
                worksheets.Cell(2, 7).Value = "Tecnico:";
                worksheets.Cell(2, 13).Value = "TEST DATA";
                worksheets.Cell(currentrow, 1).Value = "Item";
                worksheets.Cell(currentrow, 2).Value = "Process Control Point";
                worksheets.Cell(currentrow, 3).Value = "Equipment No";
                worksheets.Cell(currentrow, 4).Value = "Instrument No";
                worksheets.Cell(currentrow, 5).Value = "Control Point";
                worksheets.Cell(currentrow, 6).Value = "Control SPEC";
                worksheets.Cell(currentrow, 7).Value = "Unit";
                worksheets.Cell(currentrow, 8).Value = "Tester";

                var rango = worksheets.Range("A1:Q2"); //Seleccionamos un rango
                //rango.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick); //Generamos las lineas exteriores
                //rango.Style.Border.SetInsideBorder(XLBorderStyleValues.Medium); //Generamos las lineas interiores
                rango.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; //Alineamos horizontalmente
                rango.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;  //Alineamos verticalmente
                rango.Style.Font.FontSize = 14; //Indicamos el tamaño de la fuente
                rango.Style.Fill.BackgroundColor = XLColor.AliceBlue; //Indicamo

                foreach (var item in cho)
                {
                    currentrow++;
                    worksheets.Cell(2, 2).Value = item.Model;
                    worksheets.Cell(2, 4).Value = item.Line;
                    worksheets.Cell(2, 6).Value = item.Date.ToString(); 
                    worksheets.Cell(2, 8).Value = txtTecnico;
                    worksheets.Cell(currentrow, 1).Value = item.Item;
                    worksheets.Cell(currentrow, 2).Value = item.ProcessControlPoint;
                    worksheets.Cell(currentrow, 3).Value = item.Equipment;
                    worksheets.Cell(currentrow, 4).Value = item.Instrument;
                    worksheets.Cell(currentrow, 5).Value = item.ControlPoint;
                    worksheets.Cell(currentrow, 6).Value = item.ControlSPEC;
                    worksheets.Cell(currentrow, 7).Value = item.Unit;
                    worksheets.Cell(currentrow, 8).Value = item.Tester;

                }
                currentrow = 3;
                foreach (var item in testData)
                {
                    currentrow++;
                    foreach (var item2 in cho)
                    {
                        
                        if (item2.Item.Equals(item.ItemId))
                        {
                            worksheets.Cell(currentrow, 12).Value = item.Field1;
                            worksheets.Cell(currentrow, 13).Value = item.Field2;
                            worksheets.Cell(currentrow, 14).Value = item.Field3;
                            worksheets.Cell(currentrow, 15).Value = item.Field4;
                            worksheets.Cell(currentrow, 16).Value = item.Field5;
                        }
                    }
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream); 
                    var content = stream.ToArray();
                    //Encoding u8 = Encoding.UTF8;
                    //byte[] result = cho.SelectMany(x => u8.GetBytes(x.ToString())).ToArray();


                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "items.xlsx");
                }
            }
            //return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "users.csv");
            //return new ExcelResult<ChangeOverModel>(results, "Demo Sheet Name", "Fingers10");
        }
        public IActionResult ExportExcelProcessControl(string txtTecnico)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheets = workbook.Worksheets.Add("Items");
                var currentrow = 3;

                worksheets.Cell(2, 1).Value = "Modelo:";
                worksheets.Cell(2, 3).Value = "Linea:";
                worksheets.Cell(2, 5).Value = "Date:";
                worksheets.Cell(2, 7).Value = "Tecnico:";
                worksheets.Cell(2, 13).Value = "TEST DATA";
                worksheets.Cell(currentrow, 1).Value = "Item";
                worksheets.Cell(currentrow, 2).Value = "Process Control Point";
                worksheets.Cell(currentrow, 3).Value = "Equipment No";
                worksheets.Cell(currentrow, 4).Value = "Instrument No";
                worksheets.Cell(currentrow, 5).Value = "Values";
                worksheets.Cell(currentrow, 6).Value = "Unit";
     

                var rango = worksheets.Range("A1:M2"); //Seleccionamos un rango
                //rango.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick); //Generamos las lineas exteriores
                //rango.Style.Border.SetInsideBorder(XLBorderStyleValues.Medium); //Generamos las lineas interiores
                rango.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; //Alineamos horizontalmente
                rango.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;  //Alineamos verticalmente
                rango.Style.Font.FontSize = 14; //Indicamos el tamaño de la fuente
                rango.Style.Fill.BackgroundColor = XLColor.AliceBlue; //Indicamo

                foreach (var item in pc)
                {
                    currentrow++;
                    worksheets.Cell(2, 2).Value = item.Model;
                    worksheets.Cell(2, 4).Value = item.Line;
                    worksheets.Cell(2, 6).Value = item.Date.ToString();
                    worksheets.Cell(2, 8).Value = txtTecnico;
                    worksheets.Cell(currentrow, 1).Value = item.Item;
                    worksheets.Cell(currentrow, 2).Value = item.ProcessControlPoint;
                    worksheets.Cell(currentrow, 3).Value = item.Equipment;
                    worksheets.Cell(currentrow, 4).Value = item.Instrument;
                    worksheets.Cell(currentrow, 5).Value = item.Values;
                    worksheets.Cell(currentrow, 6).Value = item.Unit;
                   

                }
                currentrow = 3;
                foreach (var item in testData)
                {
                    currentrow++;
                    foreach (var item2 in pc)
                    {
                        if (item2.Item.Equals(item.ItemId))
                        {
                            worksheets.Cell(currentrow, 12).Value = item.Field1;
                            worksheets.Cell(currentrow, 13).Value = item.Field2;
                            worksheets.Cell(currentrow, 14).Value = item.Field3;
                            worksheets.Cell(currentrow, 15).Value = item.Field4;
                            worksheets.Cell(currentrow, 16).Value = item.Field5;
                        }
                    }
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    //Encoding u8 = Encoding.UTF8;
                    //byte[] result = cho.SelectMany(x => u8.GetBytes(x.ToString())).ToArray();


                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "items.xlsx");
                }
            }
            //return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "users.csv");
            //return new ExcelResult<ChangeOverModel>(results, "Demo Sheet Name", "Fingers10");
        }
        public void TestDataGuardar(string item, string data1, string data2, string data3, string data4, string data5)
        {
            testData.Add(new TestDataModel
            {
                ItemId = item,
                Field1 = data1,
                Field2 = data2,
                Field3 = data3,
                Field4 = data4,
                Field5 = data5

            }); ;

        }
        public IActionResult Privacy()
        {
            return View();
        }

       

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}