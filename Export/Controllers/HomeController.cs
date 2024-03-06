using Export.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Text;

namespace Export.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            var customers = new List<Customer>
            {
                new Customer
                {
                    CustomerID = "C001",
                    ContactName = "John Doe",
                    City = "New York",
                    Country = "USA"
                },
                new Customer
                {
                    CustomerID = "C002",
                    ContactName = "Jane Smith",
                    City = "London",
                    Country = "UK"
                },
                new Customer
                {
                    CustomerID = "C003",
                    ContactName = "Juan Perez",
                    City = "Mexico City",
                    Country = "Mexico"
                }
            };

            return View(customers);
        }


        [HttpPost]
        public FileResult Export()
        {
            //List<object> customers = (from customer in this.Context.Customers.Take(10)
            //                          select new[] {
            //                                customer.CustomerID,
            //                                customer.ContactName,
            //                                customer.City,
            //                                customer.Country
            //                        }).ToList<object>();

            var listcustomers = new List<Customer>
            {
                new Customer
                {
                    CustomerID = "C001",
                    ContactName = "John Doe",
                    City = "New York",
                    Country = "USA"
                },
                new Customer
                {
                    CustomerID = "C002",
                    ContactName = "Jane Smith",
                    City = "London",
                    Country = "UK"
                },
                new Customer
                {
                    CustomerID = "C003",
                    ContactName = "Juan Perez",
                    City = "Mexico City",
                    Country = "Mexico"
                }
            };
            List<object> customers = (from customer in listcustomers
                                      select new[] {
                                            customer.CustomerID,
                                            customer.ContactName,
                                            customer.City,
                                            customer.Country
                                    }).ToList<object>();

            //List<object> customers = listcustomers.ToList<object>();
            //Building an HTML string.
            StringBuilder sb = new StringBuilder();

            //Table start.
            sb.Append("<table border='1' cellpadding='5' cellspacing='0' style='border: 1px solid #ccc;font-family: Arial; font-size: 10pt;'>");

            //Building the Header row.
            sb.Append("<tr>");
            sb.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>CustomerID</th>");
            sb.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>ContactName</th>");
            sb.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>City</th>");
            sb.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>Country</th>");
            sb.Append("</tr>");

            //Building the Data rows.
            for (int i = 0; i < customers.Count; i++)
            {
                string[] customer = (string[])customers[i];
                sb.Append("<tr>");
                for (int j = 0; j < customer.Length; j++)
                {
                    //Append data.
                    sb.Append("<td style='border: 1px solid #ccc'>");
                    sb.Append(customer[j]);
                    sb.Append("</td>");
                }
                sb.Append("</tr>");
            }

            //Table end.
            sb.Append("</table>");

            return File(Encoding.UTF8.GetBytes(sb.ToString()), "application/vnd.ms-word", "Grid.doc");
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