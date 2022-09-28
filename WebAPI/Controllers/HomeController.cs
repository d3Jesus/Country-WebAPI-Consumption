using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using WebAPI.Models;
using NsExcel = Microsoft.Office.Interop.Excel;

namespace WebAPI.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        #region ALL
        [HttpGet]
        public async Task<ActionResult> GetAllCountries()
        {
            //Hosted web API REST Service base url
            string Baseurl = "https://restcountries.com/v3.1/";
            List<Country> country = null;
            using (var client = new HttpClient())
            {
                //Passing service base url
                client.BaseAddress = new Uri(Baseurl);
                client.DefaultRequestHeaders.Clear();
                //Define request data format
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                var Res = await client.GetAsync("all");
                //Checking the response is successful or not which is sent using HttpClient
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details received from web api
                    var response = Res.Content.ReadAsStringAsync().Result;
                    //Deserializing the response received from web api and storing into the Country list
                    country = JsonConvert.DeserializeObject<List<Country>>(response);
                }
            }
            return View(country);
        }
        #endregion
        
        #region NAME
        [HttpGet]
        public async Task<ActionResult> CountryDetails(string countryName)
        {
            //Hosted web API REST Service base url
            string Baseurl = "https://restcountries.com/v3.1/";
            List<Country> country = null;
            using (var client = new HttpClient())
            {
                //Passing service base url
                client.BaseAddress = new Uri(Baseurl);
                client.DefaultRequestHeaders.Clear();
                //Define request data format
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                var Res = await client.GetAsync("name/" + countryName);
                //Checking the response is successful or not which is sent using HttpClient
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details received from web api
                    var response = Res.Content.ReadAsStringAsync().Result;
                    //Deserializing the response received from web api and storing into the Country list
                    country = JsonConvert.DeserializeObject<List<Country>>(response);
                }
            }
            ViewData["CountriesList"] = country;
            //returning view
            return View();
        }
        #endregion

        #region Exportations
        public ActionResult ExportToXls()
        {
            var countriesList = TempData["CountriesList"] as List<Country>;

            //start excel
            NsExcel.Application excapp = new NsExcel.Application();

            //create a blank workbook
            var workbook = excapp.Workbooks.Add(NsExcel.XlSheetType.xlWorksheet);
            var worksheet = (NsExcel.Worksheet)excapp.ActiveSheet;

            //if you want to make excel visible change to true
            excapp.Visible = false;

            #region Name
            worksheet.Cells[1, 1] = "Names";
            worksheet.Cells[2, 1] = "Common";
            worksheet.Cells[2, 2] = countriesList.First().name.common;

            worksheet.Cells[3, 1] = "Official";
            worksheet.Cells[3, 2] = countriesList.First().name.official;
            #endregion

            #region Capital
            worksheet.Cells[4, 1] = "Capitals";
            int row = 6;
            var capitals = countriesList.First().capital;
            for (int i = 0; i < capitals.Count; i++)
            {
                worksheet.Cells[row, 1] = capitals[i];
                row++;
                i++;
            }
            #endregion

            #region Region, Subregion, Area, Population
            worksheet.Cells[(row + 1), 1] = "Region";
            worksheet.Cells[(row + 1), 2] = countriesList.First().region;

            worksheet.Cells[(row + 2), 1] = "Subregion";
            worksheet.Cells[(row + 2), 2] = countriesList.First().subregion;

            worksheet.Cells[(row + 3), 1] = "Area";
            worksheet.Cells[(row + 3), 2] = countriesList.First().area;

            worksheet.Cells[(row + 4), 1] = "Population";
            worksheet.Cells[(row + 4), 2] = countriesList.First().population;
            #endregion

            #region Timezone
            worksheet.Cells[4, 1] = "Capitals";
            row = row + 6;
            var timezones = countriesList.First().timezones;
            for (int i = 0; i < capitals.Count; i++)
            {
                worksheet.Cells[row, 1] = timezones[i];
                row++;
            }
            #endregion

            #region Flags
            worksheet.Cells[(row + 1), 1] = "PNG";
            worksheet.Cells[(row + 1), 2] = countriesList.First().flags.png;

            worksheet.Cells[(row + 2), 1] = "SVG";
            worksheet.Cells[(row + 2), 2] = countriesList.First().flags.svg;
            #endregion

            workbook.SaveAs("CountryDetails", 
                            NsExcel.XlFileFormat.xlWorkbookDefault, 
                            Type.Missing, Type.Missing, true, 
                            false, NsExcel.XlSaveAsAccessMode.xlNoChange, 
                            NsExcel.XlSaveConflictResolution.xlLocalSessionChanges, 
                            Type.Missing, Type.Missing);

            excapp.Quit();
            ////or open one - this is no pleasant, but yue're probably interested in the first parameter
            //string workbookPath = "C:/Users/"+ Environment.UserName + "/Downloads/test.xlsx";
            //workbook = excapp.Workbooks.Open(workbookPath,
            //    0, false, 5, "", "", false, NsExcel.XlPlatform.xlWindows, "",
            //    true, false, 0, true, false, false);

            ////Not done yet. You have to work on a specific sheet - note the cast
            ////You may not have any sheets at all. Then you have to add one with NsExcel.Worksheet.Add()
            //var sheet = (NsExcel.Worksheet)workbook.Sheets[1]; //indexing starts from 1

            ////do something usefull: you select now an individual cell
            //var range = sheet.get_Range("A1", "A1");
            //range.Value2 = "test"; //Value2 is not a typo

            ////now the list
            //string cellName;
            //int counter = 1;
            //foreach (var item in list)
            //{
            //    cellName = "A" + counter.ToString();
            //    range = sheet.get_Range(cellName, cellName);
            //    range.Value2 = item.ToString();
            //    ++counter;
            //}
            return RedirectToAction(nameof(GetAllCountries));
        }
        #endregion

        #region Retrieve Data
        private async Task<List<Country>> RetrieveData(string value)
        {
            //Hosted web API REST Service base url
            string Baseurl = "https://restcountries.com/v3.1/";
            List<Country> country = null;
            using (var client = new HttpClient())
            {
                //Passing service base url
                client.BaseAddress = new Uri(Baseurl);
                client.DefaultRequestHeaders.Clear();
                //Define request data format
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                var Res = await client.GetAsync(value);
                //Checking the response is successful or not which is sent using HttpClient
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details received from web api
                    var response = Res.Content.ReadAsStringAsync().Result;
                    //Deserializing the response received from web api and storing into the Country list
                    country = JsonConvert.DeserializeObject<List<Country>>(response);
                }
            }
            return country;
        }
        #endregion

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
