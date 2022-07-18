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
            IList<Country> countries = null;
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
                    countries = JsonConvert.DeserializeObject<IList<Country>>(response);
                }
                //returning the coutries list to view
                return View(countries);
            }
        }
        #endregion
        
        #region NAME
        [HttpGet]
        public async Task<ActionResult> CountryDetails(string countryName)
        {
            //Hosted web API REST Service base url
            string Baseurl = "https://restcountries.com/v3.1/";
            Country country = null;
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
                    country = JsonConvert.DeserializeObject<Country>(response);
                }
                //returning the coutries list to view
                return View(country);
            }
        }
        #endregion


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
