using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using YelpSharp;
using YelpSharp.Data.Options;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web;
using System.Reflection;
using Search_Engine_Lead_Finder;
using Search_Engine_Lead_Finder.Properties;
using System.Xml;
using System.Text.RegularExpressions;
using FactualDriver;
using FactualDriver.Exceptions;
using YelpSharp.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Specialized;

namespace Houzz_Lead_Spider
{
    public partial class formMain : System.Windows.Forms.Form
    {
        public formMain()
        {
            InitializeComponent();
        }

        string apiGoogleKey = "AIzaSyALG4QStYnoCmzby8HTJGX0XtM7tBGXdbw";
        string apiYPKey = "wdfmzvlbx8";
        string apiFactualKey = "rSSWRWCIxlAvMkQnNGp53u60l3tM4pvAtkrn275U";
        string apiFactualSecret = "lOwvEN9Bync0bRyHKr9ZXgOdWL74cwyaf7UP0err";
        string searchEndPoint = "http://api.sensis.com.au/v1/test/search";
        string apiSensisKey = "3vveh8nddq5mv68j92zj5fur";
        string apiFacebookKey = "674725399294850|gW8b3ArDjkB6pd_uLWvYIcxtN1I";

        public class LocationFacebook
        {
            [JsonProperty(PropertyName = "city")]
            public string city { get; set; }
            [JsonProperty(PropertyName = "state")]
            public string state { get; set; }
            [JsonProperty(PropertyName = "country")]
            public string country { get; set; }
            [JsonProperty(PropertyName = "street")]
            public string street { get; set; }
            [JsonProperty(PropertyName = "zip")]
            public string zip { get; set; }
        }

        public class DatumFacebook
        {
            public string name { get; set; }
            public string id { get; set; }
            public string category { get; set; }
            public string link { get; set; }
            public string website { get; set; }
            public LocationFacebook location { get; set; }
            public string phone { get; set; }
            public List<string> emails { get; set; }
            public bool is_unclaimed { get; set; }
        }

        public class CursorsFacebook
        {
            public string before { get; set; }
            public string after { get; set; }
        }

        public class Paging
        {
            public CursorsFacebook cursors { get; set; }
        }

        public class ErrorFacebook
        {
            public string message { get; set; }
            public string type { get; set; }
            public int code { get; set; }
        }

        public class RootObjectFacebook
        {
            [JsonProperty(PropertyName = "data")]
            public List<DatumFacebook> data { get; set; }
            [JsonProperty(PropertyName = "paging")]
            public Paging paging { get; set; }
            [JsonProperty(PropertyName = "error")]
            public ErrorFacebook error { get; set; }
        }

        public class PrimaryAddress
        {
            [JsonProperty(PropertyName = "state")]
            public string state { get; set; }
            [JsonProperty(PropertyName = "postcode")]
            public string postcode { get; set; }
            [JsonProperty(PropertyName = "suburb")]
            public string suburb { get; set; }
            [JsonProperty(PropertyName = "addressLine")]
            public string addressLine { get; set; }
        }

        public class SensisCategories
        {
            [JsonProperty(PropertyName = "id")]
            public string id { get; set; }
            [JsonProperty(PropertyName = "name")]
            public string name { get; set; }
        }

        public class Result
        {
            [JsonProperty(PropertyName = "primaryContacts")]
            public List<PrimaryContact> primaryContacts { get; set; }
            [JsonProperty(PropertyName = "categories")]
            public List<SensisCategories> categories { get; set; }
            [JsonProperty(PropertyName = "primaryAddress")]
            public PrimaryAddress primaryAddress { get; set; }
            [JsonProperty(PropertyName = "name")]
            public string name { get; set; }
            [JsonProperty(PropertyName = "externalLinks")]
            public List<ExternalLink> externalLinks { get; set; }
        }

        public class SearchResponse
        {
            [JsonProperty(PropertyName = "code")]
            public int code;
            [JsonProperty(PropertyName = "message")]
            public string message;
            [JsonProperty(PropertyName = "totalResults")]
            public int totalResults;
            [JsonProperty(PropertyName = "results")]
            public List<Result> results { get; set; }
            [JsonProperty(PropertyName = "totalPages")]
            public int totalPages { get; set; }
        }

        public class PrimaryContact
        {
            [JsonProperty(PropertyName = "type")]
            public string type { get; set; }
            [JsonProperty(PropertyName = "value")]
            public string value { get; set; }
            [JsonProperty(PropertyName = "description")]
            public string description { get; set; }
        }

        public class ExternalLink
        {
            [JsonProperty(PropertyName = "url")]
            public string url { get; set; }
            [JsonProperty(PropertyName = "displayValue")]
            public string displayValue { get; set; }
            [JsonProperty(PropertyName = "label")]
            public string label { get; set; }
            [JsonProperty(PropertyName = "type")]
            public string type { get; set; }
        }


        public class SsapiSearcher
        {
            readonly Uri endPoint;
            readonly string apiKey;


            public SsapiSearcher(string endPoint, string apiKey)
            {
                this.endPoint = new Uri(endPoint);
                this.apiKey = apiKey;
            }


            public SearchResponse SearchFor(string query, string location, string state, int page)
            {
                // Build the API request
                var url = new Uri(endPoint, "?query=" + Uri.EscapeDataString(query) 
                                            + "&location=" + Uri.EscapeDataString(location) 
                                            + "&key=" + Uri.EscapeDataString(apiKey)
                                            + "&state=" + Uri.EscapeDataString(state) 
                                            + "&rows=30"
                                            + "&page=" + page.ToString() );
                var req = WebRequest.Create(url);

                // Send the request and read the response
                using (var res = req.GetResponse())
                {
                    Stream dataStream = res.GetResponseStream();
                    StreamReader reader = new StreamReader(dataStream);
                    string responseFromServer = reader.ReadToEnd();

                    reader.Close();
                    dataStream.Close();
                    res.Close();

                    var jarray = JsonConvert.DeserializeObject<SearchResponse>(responseFromServer);

                    return jarray as SearchResponse;
                }
            }
        }

        public class InputParamsListingId
        {
            [JsonProperty(PropertyName = "__invalid_name__@xsi.type")]
            public string __invalid_name__type { get; set; }
            public string appId { get; set; }
            public string dnt { get; set; }
            public string format { get; set; }
            public string userIpAddress { get; set; }
            public string userReferrer { get; set; }
            public string requestId { get; set; }
            public string test { get; set; }
            public string userAgent { get; set; }
            public string visitorId { get; set; }
            public int listingId { get; set; }
        }

        public class MetaPropertiesListingId
        {
            [JsonProperty(PropertyName = "__invalid_name__@xsi.type")]
            public string __invalid_name__type { get; set; }
            public string errorCode { get; set; }
            public int listingCount { get; set; }
            public string message { get; set; }
            public InputParams inputParams { get; set; }
            public string requestId { get; set; }
            public string resultCode { get; set; }
            public string totalAvailable { get; set; }
            public string trackingRequestURL { get; set; }
        }

        public class Categories
        {
            public List<string> category { get; set; }
        }

        public class StandardHours
        {
            public string friday { get; set; }
            public string monday { get; set; }
            public string saturday { get; set; }
            public string sunday { get; set; }
            public string thursday { get; set; }
            public string tuesday { get; set; }
            public string wednesday { get; set; }
        }

        public class DefaultHours
        {
            public StandardHours standardHours { get; set; }
        }

        public class DetailedHours
        {
            public DefaultHours defaultHours { get; set; }
        }

        public class ExtraWebsiteURLs
        {
            public List<string> extraWebsiteURL { get; set; }
        }

        public class Akas
        {
            public List<string> aka { get; set; }
        }

        public class ListingDetail
        {
            [JsonProperty(PropertyName = "extraWebsiteURLs")]
            public ExtraWebsiteURLs extraWebsiteURLs { get; set; }
            [JsonProperty(PropertyName = "email")]
            public string email { get; set; }
        }

        public class ListingsDetails
        {
            public List<ListingDetail> listingDetail { get; set; }
        }

        public class ListingsDetailsResult
        {
            public MetaPropertiesListingId metaProperties { get; set; }
            public ListingsDetails listingsDetails { get; set; }
        }

        public class RootObjectYPListingId
        {
            public ListingsDetailsResult listingsDetailsResult { get; set; }
        }

        public class RelatedCategory
        {
            public int count { get; set; }
            public string name { get; set; }
        }

        public class RelatedCategories
        {
            [JsonProperty(PropertyName = "__invalid_name__@xsi.type")]
            public string __invalid_name__type { get; set; }
            public List<RelatedCategory> relatedCategory { get; set; }
        }

        public class InputParams
        {
            [JsonProperty(PropertyName = "__invalid_name__@xsi.type")]
            public string __invalid_name__type { get; set; }
            public string appId { get; set; }
            public string dnt { get; set; }
            public string format { get; set; }
            public string userIpAddress { get; set; }
            public string userReferrer { get; set; }
            public string requestId { get; set; }
            public bool shortUrl { get; set; }
            public string test { get; set; }
            public string userAgent { get; set; }
            public string visitorId { get; set; }
            public int listingCount { get; set; }
            public bool phoneSearch { get; set; }
            public int radius { get; set; }
            public string searchLocation { get; set; }
            public string term { get; set; }
            public string termType { get; set; }
            [JsonProperty(PropertyName = "pageNum")]
            public int pageNum { get; set; }
        }

        public class MetaProperties
        {
            public string didYouMean { get; set; }
            public string errorCode { get; set; }
            public int listingCount { get; set; }
            public string message { get; set; }
            public RelatedCategories relatedCategories { get; set; }
            public InputParams inputParams { get; set; }
            public string requestId { get; set; }
            public string resultCode { get; set; }
            public string searchCity { get; set; }
            public double searchLat { get; set; }
            public double searchLon { get; set; }
            public string searchState { get; set; }
            public string searchType { get; set; }
            public string searchZip { get; set; }
            public int totalAvailable { get; set; }
            public string trackingRequestURL { get; set; }
            public string ypcAttribution { get; set; }
        }

        public class SearchListing
        {
            public string adImage { get; set; }
            public string adImageClick { get; set; }
            public string additionalText { get; set; }
            public string audioURL { get; set; }
            public double averageRating { get; set; }
            public string baseClickURL { get; set; }
            public string businessName { get; set; }
            public string businessNameURL { get; set; }
            public Categories categories { get; set; }
            public string city { get; set; }
            public bool claimed { get; set; }
            public bool claimedStatus { get; set; }
            public bool couponFlag { get; set; }
            public string couponImageURL { get; set; }
            public string couponTitle { get; set; }
            public string couponURL { get; set; }
            public string customLink { get; set; }
            public string customLinkText { get; set; }
            public string description1 { get; set; }
            public string description2 { get; set; }
            public double distance { get; set; }
            public string distribAdImage { get; set; }
            public string distribAdImageClick { get; set; }
            public string email { get; set; }
            public bool hasDisplayAddress { get; set; }
            public bool hasPriorityShading { get; set; }
            public bool isRedListing { get; set; }
            public string latitude { get; set; }
            public int listingId { get; set; }
            public string longitude { get; set; }
            public string moreInfoURL { get; set; }
            public string noAddressMessage { get; set; }
            public bool omitAddress { get; set; }
            public bool omitPhone { get; set; }
            public string openHours { get; set; }
            public string openStatus { get; set; }
            public string paymentMethods { get; set; }
            public string phone { get; set; }
            public string pricePerCall { get; set; }
            public string primaryCategory { get; set; }
            public string printAdImage { get; set; }
            public string printAdImageClick { get; set; }
            public int ratingCount { get; set; }
            public string ringToNumberDisplay { get; set; }
            public string searchResultType { get; set; }
            public string services { get; set; }
            public string slogan { get; set; }
            public string state { get; set; }
            public string street { get; set; }
            public string videoURL { get; set; }
            public string viewPhone { get; set; }
            public string websiteDisplayURL { get; set; }
            public string websiteURL { get; set; }
            public string zip { get; set; }
        }

        public class SearchListings
        {
            public List<SearchListing> searchListing { get; set; }
        }

        public class SearchResult
        {
            public MetaProperties metaProperties { get; set; }
            public SearchListings searchListings { get; set; }
        }

        public class RootObjectYP
        {
            public SearchResult searchResult { get; set; }
        }

        public class Datum
        {
            [JsonProperty(PropertyName = "address")]
            public string address { get; set; }
            [JsonProperty(PropertyName = "country")]
            public string country { get; set; }
            [JsonProperty(PropertyName = "email")]
            public string email { get; set; }
            [JsonProperty(PropertyName = "factual_id")]
            public string factual_id { get; set; }
            [JsonProperty(PropertyName = "locality")]
            public string locality { get; set; }
            [JsonProperty(PropertyName = "name")]
            public string name { get; set; }
            [JsonProperty(PropertyName = "postcode")]
            public string postcode { get; set; }
            [JsonProperty(PropertyName = "region")]
            public string region { get; set; }
            [JsonProperty(PropertyName = "tel")]
            public string tel { get; set; }
            [JsonProperty(PropertyName = "website")]
            public string website { get; set; }
            [JsonProperty(PropertyName = "address_extended")]
            public string address_extended { get; set; }
            [JsonProperty(PropertyName = "category_labels")]
            public List<List<string>> category_labels { get; set; }
        }

        public class Response
        {
            public List<Datum> data { get; set; }
            public int included_rows { get; set; }
        }

        public class RootObjectFactual
        {
            public int version { get; set; }
            public string status { get; set; }
            public Response response { get; set; }
        }


        class Config
        {
            private static Options _options;

            /// <summary>
            /// return the oauth options for using the Yelp API.  I store my keys in the environment settings, but you
            /// can just write them out here, or put them into an app.config file.  For more info, visit
            /// http://www.yelp.com/developers/getting_started/api_access
            /// </summary>
            /// <returns></returns>
            public static Options Options
            {
                get
                {
                    if (_options == null)
                    {
                        // get all of the options out of EnvironmentSettings.  You can easily just put your own keys in here without
                        // doing the env dance, if you so choose
                        _options = new Options()
                        {
                            AccessToken = "NNLmO1zDY4Iz4F3jBSPrnu2-YTs6FtOF",
                            AccessTokenSecret = "s-8yBVI05mf6z04o-283_nsmE9s",
                            ConsumerKey = "OhcAQFH47egXMNWvjQAoUA",
                            ConsumerSecret = "Ttr5avNwaiahY9xye9nGkRHKvno"
                        };

                        if (String.IsNullOrEmpty(_options.AccessToken) ||
                            String.IsNullOrEmpty(_options.AccessTokenSecret) ||
                            String.IsNullOrEmpty(_options.ConsumerKey) ||
                            String.IsNullOrEmpty(_options.ConsumerSecret))
                        {
                            throw new InvalidOperationException("No OAuth info available.  Please modify Config.cs to add your YELP API OAuth keys");
                        }
                    }
                    return _options;
                }
            }
        }

        public class AddressComponent
        {
            public string long_name { get; set; }
            public string short_name { get; set; }
            public List<string> types { get; set; }
        }

        public class Location
        {
            public double lat { get; set; }
            public double lng { get; set; }
        }

        public class Geometry
        {
            public Location location { get; set; }
        }

        public class Aspect
        {
            public int rating { get; set; }
            public string type { get; set; }
        }

        public class Review
        {
            public List<Aspect> aspects { get; set; }
            public string author_name { get; set; }
            public string author_url { get; set; }
            public string language { get; set; }
            public int rating { get; set; }
            public string text { get; set; }
            public int time { get; set; }
        }

        public class Details
        {
            public List<AddressComponent> address_components { get; set; }
            public string adr_address { get; set; }
            public string formatted_address { get; set; }
            public string formatted_phone_number { get; set; }
            public Geometry geometry { get; set; }
            public string icon { get; set; }
            public string id { get; set; }
            public string international_phone_number { get; set; }
            public string name { get; set; }
            public string place_id { get; set; }
            public string reference { get; set; }
            public List<Review> reviews { get; set; }
            public string scope { get; set; }
            public List<string> types { get; set; }
            public string url { get; set; }
            public int user_ratings_total { get; set; }
            public int utc_offset { get; set; }
            public string vicinity { get; set; }
            public string website { get; set; }
        }

        public class RootObjectCompany
        {
            [JsonProperty(PropertyName = "error_message")]
            public string error_message { get; set; }
            [JsonProperty(PropertyName = "html_attributions")]
            public List<object> html_attributions { get; set; }
            [JsonProperty(PropertyName = "result")]
            public Details result { get; set; }
            [JsonProperty(PropertyName = "status")]
            public string status { get; set; }
        }

        public class OpeningHours
        {
            public bool open_now { get; set; }
            public List<object> weekday_text { get; set; }
        }

        public class Photo
        {
            public int height { get; set; }
            public List<string> html_attributions { get; set; }
            public string photo_reference { get; set; }
            public int width { get; set; }
        }

        public class Places
        {
            public string formatted_address { get; set; }
            public Geometry geometry { get; set; }
            public string icon { get; set; }
            public string id { get; set; }
            public string name { get; set; }
            public string place_id { get; set; }
            public string reference { get; set; }
            public List<string> types { get; set; }
            public OpeningHours opening_hours { get; set; }
            public List<Photo> photos { get; set; }
            public bool? permanently_closed { get; set; }
        }

        public class RootObjectPlaces
        {
            [JsonProperty(PropertyName = "error_message")]
            public string error_message { get; set; }
            [JsonProperty(PropertyName = "html_attributions")]
            public List<object> html_attributions { get; set; }
            [JsonProperty(PropertyName = "next_page_token")]
            public string next_page_token { get; set; }
            [JsonProperty(PropertyName = "results")]
            public List<Places> results { get; set; }
            [JsonProperty(PropertyName = "status")]
            public string status { get; set; }
        }

        void saveDataSet()
        {
            string[] path = new string[17];
            for (int i = 0; i <= 15; i++)
                path[i] = Directory.GetCurrentDirectory() + "\\dataColumns\\" + i.ToString() + ".txt";

            for (int i = 0; i < dataCompanyList.ColumnCount; i++)
            {
                string[] data = new string[999999];
                foreach (DataGridViewRow row in dataCompanyList.Rows)
                {
                    if (row.Cells[0].Value != null)
                    {
                        if (row.Cells[i].Value == null)
                            data[row.Index] += "Empty Data";
                        else if (row.Cells[i].Value.ToString() == "")
                            data[row.Index] += "Empty Data";
                        else
                            data[row.Index] += row.Cells[i].Value.ToString();
                    }
                }
                data = data.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                File.WriteAllLines(path[i], data); 
            }                  
        }

        void fixDataSetNullsToEmptyStrings()
        {
            foreach (DataGridViewRow iRow in dataCompanyList.Rows)
            {
                if (iRow.Index != dataCompanyList.Rows.Count - 1)
                {
                    foreach (DataGridViewCell cell in iRow.Cells)
                    {
                        if (cell.Value == null)
                            cell.Value = "";
                    }
                }
            }
        }

        void loadDataset()
        {
            string[] path = new string[16];
            Dictionary<string, string[]> columndata = new Dictionary<string, string[]>();
            for (int i = 0; i <= 15; i++)
            {
                path[i] = Directory.GetCurrentDirectory() + "\\dataColumns\\" + i.ToString() + ".txt";
                columndata.Add(i.ToString(), File.ReadAllLines(path[i]));
            }
            for (int row = 0; row < columndata["0"].Length; row++)
            {
                string[] data = new string[16];
                if (columndata["0"].GetValue(row) != null & columndata["0"].GetValue(row).ToString() != "")
                {
                    for (int column = 0; column <= 15; column++)
                    {
                        try
                        {
                            if (columndata[column.ToString()].GetValue(row) != null)
                                data[column] = columndata[column.ToString()].GetValue(row).ToString() == "Empty Data" ? "" : columndata[column.ToString()].GetValue(row).ToString();
                        }
                        catch {
                            data[column] = "";
                        }
                    }
                    dataCompanyList.Rows.Add(data);
                }
            }
        }

        HtmlElement getElementbyName(WebBrowser webbrowser, string tagname, string name)
        {
            HtmlElementCollection list = webBrowser.Document.GetElementsByTagName(tagname);
            foreach (HtmlElement item in list)
            {
                try
                {
                    if (item.GetAttribute("name") == name)
                    {
                        return item;
                    }
                }
                catch { }
            }
            return null;
        }

        bool duplicateCheckV2(string companyName, string companyPhone, string companyWebsite, string companyEmail)
        {
            navigateAndWait("http://waypoint.houzz.net/");
            while (webBrowser.Document.GetElementById("compiled_leads_company_name") == null)
                Application.DoEvents();
            webBrowser.Document.GetElementById("compiled_leads_company_name").SetAttribute("value", companyName);
            webBrowser.Document.GetElementById("compiled_leads_website").SetAttribute("value", companyWebsite);
            webBrowser.Document.GetElementById("compiled_leads_email_address").SetAttribute("value", companyEmail);
            webBrowser.Document.GetElementById("compiled_leads_phone_number").SetAttribute("value", companyPhone);
            getElementbyName(webBrowser, "input", "commit").InvokeMember("Click");
            while (webBrowser.Document.Body == null || webBrowser.Document.Body.InnerHtml.Contains("results will only indicate they are similar"))
                Application.DoEvents();
            if (webBrowser.Document.Body.InnerHtml.Contains("DUPLICATED"))
                return true;
            else if (webBrowser.Document.Body.InnerHtml.Contains("CHECK SIMILAR RESULTS"))
            {
                string results = webBrowser.Document.Body.InnerHtml.Substring(webBrowser.Document.Body.InnerHtml.IndexOf("CHECK SIMILAR RESULTS"));
                if (companyPhone.StartsWith("64") & searchCountry.Text == "NZ")
                    companyPhone = companyPhone.Substring(2);
                if (results.Contains(companyWebsite.Replace("http://www.", "")) & companyWebsite.Length >= 4)
                    return true;
                if (companyPhone.Length >= 4)
                    if (results.Contains(companyPhone.Insert(6, "-").Insert(3, ")").Insert(0, "(")))
                        return true;
                if (companyEmail.Length > 4)
                    if (results.Contains(companyEmail))
                        return true;
            }
            else if (webBrowser.Document.Body.InnerHtml.Contains("We're sorry, but something went wrong"))
                return duplicateCheckV2(companyName, companyPhone, companyWebsite, companyEmail);
            return false;
        }

        private void ToCsV(DataGridView dGV, string filename)
        {
            string stOutput = "";
            // Export titles:
            string sHeaders = "";

            for (int j = 0; j < dGV.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
            stOutput += sHeaders + "\r\n";
            // Export data.
            for (int i = 0; i <= dGV.RowCount - 1; i++)
            {
                string stLine = "";
                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); //write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();
        }  

        string getMapList(string searchText)
        {
            // Create a request using a URL that can receive a post. 
            WebRequest request =
                   WebRequest.Create("https://maps.googleapis.com/maps/api/place/textsearch/json?query=" + Uri.EscapeDataString(searchText)
                   + "&key=" + apiGoogleKey);
            // Get the response.
            WebResponse response = request.GetResponse();
            // Get the stream containing content returned by the server.
            Stream dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Clean up the streams.
            reader.Close();
            dataStream.Close();
            response.Close();

            return responseFromServer;
        }

        string getPlaceDetails (string searchText)
        {
            // Create a request using a URL that can receive a post. 
            WebRequest request =
                   WebRequest.Create("https://maps.googleapis.com/maps/api/place/details/json?placeid=" + Uri.EscapeDataString(searchText)
                   + "&key=" + apiGoogleKey);
            // Get the response.
            WebResponse response = request.GetResponse();
            // Get the stream containing content returned by the server.
            Stream dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Clean up the streams.
            reader.Close();
            dataStream.Close();
            response.Close();

            return responseFromServer;
        }



        bool checkIfPhoneNumberFormat(HtmlElement span)
        {
            if( searchCountry.Text == "US" )
                if (span.InnerText.Length == 14 & span.InnerText.Contains("(") & span.InnerText.Contains(")") & span.InnerText.Contains("-"))
                    return true;
            if( searchCountry.Text == "AU" )
            {
                string data = span.InnerText.Replace(" ","");
                //instantiate with this pattern 
                Regex AUphoneRegex = new Regex(@"^\({0,1}((0|\+61)(2|4|3|7|8)){0,1}\){0,1}(\ |-){0,1}[0-9]{2}(\ |-){0,1}[0-9]{2}(\ |-){0,1}[0-9]{1}(\ |-){0,1}[0-9]{3}$",
                    RegexOptions.IgnoreCase);
                //find items that matches with our pattern
                Match AUphoneMatch = AUphoneRegex.Match(data);
                string AUphoneFound = AUphoneMatch.Value.ToString();
                if (AUphoneFound != "")
                    return true;
            }
            else if (searchCountry.Text == "SG")
            {
                string data = span.InnerText.Replace(" ", "");
                //instantiate with this pattern 
                Regex AUphoneRegex = new Regex(@"^(\+|\d)[0-9]{7,16}$",
                    RegexOptions.IgnoreCase);
                //find items that matches with our pattern
                Match AUphoneMatch = AUphoneRegex.Match(data);
                string AUphoneFound = AUphoneMatch.Value.ToString();
                if (AUphoneFound != "")
                    return true;
            }
            else if (searchCountry.Text == "NZ")
            {
                string data = span.InnerText.Replace(" ", "");
                data = Regex.Replace(data, @"[^0-9\s]", string.Empty);
                //instantiate with this pattern 
                if ((data.StartsWith("64") || data.StartsWith("0")) & data.Length >= 8 || (data.StartsWith("9") & data.Length == 8))
                    return true;
            }
            return false;
        }

        string fixPhoneNumberFormat(string phone)
        {
            return phone.Replace("(","").Replace(")","").Replace("-","").Replace(" ","").Replace(".","").Replace("+","");
        }

        string formatCompanyWebsite(string input)
        {
            if (input != null)
            {
                input = input.Replace("http:///","").Replace("http://", "").Replace("https://","").Replace("http:/","").Replace("http/","").Replace("http:","").Replace("http//","").Replace("www.", "");
                input = "http://www." + input;
                if (!input.Contains("yelp.com") & !input.Contains("facebook.com") & !input.Contains("angieslist.com") &
                    !input.Contains("bbbxpress.com") & !input.Contains("bbb.org") & !input.Contains("sites.google.com") &
                    !input.Contains("knoxnews.com") )
                {
                    try
                    {
                        var uri = new Uri(input);
                        var host = uri.GetLeftPart(System.UriPartial.Authority);
                        return host;
                    }
                    catch { }
                }
            }
            return input;
        }

        string getWebsite()
        {
            HtmlElementCollection links = webBrowser.Document.Links;
            foreach (HtmlElement l in links)
            {
                try
                {
                    string Id = l.GetAttribute("compid");
                    if (Id == "Profile_Website")
                    {
                        string websiteFound = l.GetAttribute("href");
                        string rawWebsite = formatCompanyWebsite(websiteFound);
                        rawWebsite = rawWebsite.Replace("http://", "");
                        //rawWebsite = rawWebsite.Replace("www.", "");
                        return rawWebsite;
                    }
                }
                catch { }
            }
            return "No website";
        }

        string getPhoneNumber()
        {
            HtmlElementCollection spanlist = webBrowser.Document.GetElementsByTagName("span");
            foreach (HtmlElement span in spanlist)
            {
                try
                {
                    if (checkIfPhoneNumberFormat(span))
                    {
                        return fixPhoneNumberFormat(span.InnerText);
                    }
                }
                catch { }
            }
            return "No phone number";
        }

        string getAddress()
        {
            HtmlElementCollection spanlist = webBrowser.Document.GetElementsByTagName("span");
            foreach (HtmlElement span in spanlist)
            {
                try
                {
                    if (span.GetAttribute("itemprop") == "streetAddress")
                    {
                        return span.InnerText;
                    }
                }
                catch { }
            }
            return "No address";
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);


        private void ClickOKButton()
        {
            IntPtr hwnd = FindWindow("#32770", "Message from webpage");
            if (hwnd != IntPtr.Zero)
            {
                hwnd = FindWindowEx(hwnd, IntPtr.Zero, "Button", "OK");
                uint message = 0xf5;
                SendMessage(hwnd, message, IntPtr.Zero, IntPtr.Zero);
            }
        }

        bool CheckURLValid(string source)
        {
            Uri uriResult;
            return Uri.TryCreate(source, UriKind.Absolute, out uriResult) && uriResult.Scheme == Uri.UriSchemeHttp;
        }

        bool fixFrozen = false;
        void navigateAndWait(string url)
        {
            int frozentime = 0;
            int waittime = 90000;

            if (CheckURLValid(url))
            {
                webBrowser.Navigate(url);
                while (webBrowser.IsBusy)
                {
                    ClickOKButton();
                    Application.DoEvents();
                }
                try
                {
                    while (webBrowser.ReadyState != WebBrowserReadyState.Complete)
                    {
                        if (webBrowser.ReadyState == WebBrowserReadyState.Interactive & webBrowser.Document != null)
                        {
                            if (webBrowser.Document.Body != null)
                                if (webBrowser.Document.Body.InnerHtml != null)
                                    frozentime++;
                        }
                        if (fixFrozen || (frozentime > waittime & webBrowser.Document != null))
                        {
                            fixFrozen = false;
                            frozentime = 0;
                            textDebugger.AppendText("Detected frozen webpage for 90 seconds. Attempting to continue bot!" + Environment.NewLine);
                            break;
                        }
                        ClickOKButton();
                        Application.DoEvents();
                    }
                }
                catch { }
            }
        }

        string getFirstTwoWords(string input)
        {
            if (input.Contains(" ") & input.Length > 8)
            {
                List<string> allWords = input.Split(' ').ToList<string>();
                try
                {
                    string firstTwoWords = allWords[0] + " " + allWords[1];
                    if (firstTwoWords.Length < 9)
                        firstTwoWords += " " + allWords[2];
                    return firstTwoWords;
                }
                catch { return input; }
            }
            return input;
        }

        List<string> checkedHouzzProfiles = new List<string>();
        string checkIfHouzzDuplicate(string companyHouzzSearch, string companyPhone, string companyAddress, string companyWebsite)
        {
            navigateAndWait(companyHouzzSearch);

            List<string> proUrls = new List<string>();
            List<string> pro2Urls = new List<string>();

            HtmlElementCollection links = webBrowser.Document.Links;
            foreach (HtmlElement l in links)
            {
                try
                {
                    string linkFound = l.GetAttribute("href");
                    if (linkFound.Contains("www.houzz.com/pro/"))
                    {
                        if (!proUrls.Contains(linkFound) & !checkedHouzzProfiles.Contains(linkFound))
                            proUrls.Add(linkFound);
                    }
                    if (linkFound.Contains("www.houzz.com/pro2/"))
                    {
                        if (!pro2Urls.Contains(linkFound) & !checkedHouzzProfiles.Contains(linkFound))
                            pro2Urls.Add(linkFound);
                    }
                    checkedHouzzProfiles.Add(linkFound);
                }
                catch { }
            }


            string result = "No pro found";
            proUrls.ForEach(delegate(string linkUrl)
            {
                if ( result == "No pro found")
                {
                    navigateAndWait(linkUrl);
                    string phoneFound = getPhoneNumber();
                    string websiteFound = getWebsite();
                    string addressFound = getFirstTwoWords(getAddress());
                    if ((phoneFound.Contains(companyPhone) || companyPhone.EndsWith(phoneFound) ) & companyPhone != "")
                    {
                        textDebugger.AppendText("Found Houzz Dupe: " + phoneFound + " equals " + companyPhone + Environment.NewLine);
                        result = "Pro found";
                    }
                    if (formatCompanyWebsite(companyWebsite).Contains(websiteFound) & websiteFound.Length >= 5)
                    {
                        textDebugger.AppendText("Found Houzz Dupe: " + websiteFound + " within " + formatCompanyWebsite(companyWebsite) + Environment.NewLine);
                        result = "Pro found";
                    }
                    if (!addressFound.ToLower().StartsWith("unit") & addressFound.Length > 4 & getFirstTwoWords(companyAddress).Contains(addressFound))
                    {
                        textDebugger.AppendText("Found Houzz Dupe: " + addressFound + " within " + getFirstTwoWords(companyAddress) + Environment.NewLine);
                        result = "Pro found";
                    }
                }
            });

            if (result == "No pro found")
            {
                pro2Urls.ForEach(delegate(string linkUrl)
                {
                    if (result == "No pro found")
                    {
                        navigateAndWait(linkUrl);
                        string phoneFound = getPhoneNumber();
                        string websiteFound = getWebsite();
                        if ((phoneFound == companyPhone || (companyWebsite.Contains(websiteFound) & websiteFound.Length >= 5) ))
                            result = linkUrl;
                    }
                });
            }           

            return result;
        }

        string replaceCompanyNamePTE(string companyName)
        {
            try
            {
                companyName = getFirstTwoWords(companyName);
                companyName = companyName.Replace(",", "");
            }
            catch { }

            try
            {
                if (companyName.EndsWith("."))
                    companyName = companyName.Remove(companyName.Length - 1);
                string companyNameRight = companyName.Substring(companyName.LastIndexOf(' ') + 1).ToLower();
                if (companyNameRight == "co"
                    || companyNameRight == "inc"
                    || companyNameRight == "company"
                    || companyNameRight == "llc")
                    companyName = companyName.Remove(companyName.LastIndexOf(' ') + 1);
                if (companyName.EndsWith(",") || companyName.EndsWith(" ") || companyName.EndsWith("."))
                    companyName = companyName.Remove(companyName.Length - 1);
            }
            catch { }

            return companyName;
        }

        string generateHouzzSearchLink(string criteria, string location, bool includeQuotations)
        {
            string companyHouzzSearch = "none";
            if (criteria.Length >= 4)
            {
                if (includeQuotations)
                    companyHouzzSearch = "http://www.houzz.com/professionals/s/\"" +
                                criteria +
                                "\"/c/" +
                                location;
                else
                    companyHouzzSearch = "http://www.houzz.com/professionals/s/" +
                                criteria +
                                "/c/" +
                                location;
            }

            return companyHouzzSearch;
        }

        bool checkDuplicate(Excel._Worksheet xlWorksheet, string companyName, string companyPhone, string companyEmail, 
            string companyWebsite)
        {
            Excel.Range xlRange = xlWorksheet.UsedRange;
            xlRange.Cells[13, 3].Value2 = companyName;
            xlRange.Cells[13, 4].Value2 = companyPhone;
            xlRange.Cells[13, 5].Value2 = companyEmail;
            xlRange.Cells[13, 6].Value2 = companyWebsite;
            xlWorksheet.Calculate();

            string duplicateCheckName = (string)(xlRange.Cells[15, 3] as Excel.Range).Value2;
            string duplicateCheckPhone = (string)(xlRange.Cells[15, 4] as Excel.Range).Value2;
            string duplicateCheckEmail = (string)(xlRange.Cells[15, 5] as Excel.Range).Value2;
            string duplicateCheckWebsite = (string)(xlRange.Cells[15, 6] as Excel.Range).Value2;

            if (/*duplicateCheckName == "DUPLICATE!!" || */ /* Took out check company name, it sucks */
                duplicateCheckPhone == "DUPLICATE!!" ||
                duplicateCheckEmail == "DUPLICATE!!" ||
                duplicateCheckWebsite == "DUPLICATE!!")
            {
                textDebugger.AppendText("Lead dupe excel found: " + String.Concat(companyName, ",", companyEmail, ",", companyWebsite)
                    + Environment.NewLine);
                return true;
            }
            else return false;
        }

        string findAboutMelink()
        {
            HtmlElementCollection links = webBrowser.Document.Links;
            foreach (HtmlElement l in links)
            {
                try
                {
                    string aboutlink = l.GetAttribute("href");
                    if (aboutlink.ToLower().Contains("about"))
                        return aboutlink;
                    string aboutInnerText = l.InnerText;
                    if (aboutInnerText.ToLower().Contains("about"))
                        return l.GetAttribute("href");
                }
                catch { }
            }
            return "No aboutme link";
        }

        string findLink(string textToFind)
        {
            HtmlElementCollection links = webBrowser.Document.Links;
            foreach (HtmlElement l in links)
            {
                try
                {
                    string aboutlink = l.GetAttribute("href");
                    if (aboutlink.ToLower().Contains(textToFind))
                        return aboutlink;
                    string aboutInnerText = l.InnerText;
                    if (aboutInnerText.ToLower().Contains(textToFind))
                        return l.GetAttribute("href");
                }
                catch { }
            }
            return "No link";
        }

        string findContactlink()
        {
            HtmlElementCollection links = webBrowser.Document.Links;
            foreach (HtmlElement l in links)
            {
                try
                {
                    string contactlink = l.GetAttribute("href");
                    if ( contactlink.ToLower().Contains("contact") )
                        return contactlink;
                    string contactInnerText = l.InnerText;
                    if (contactInnerText.ToLower().Contains("contact"))
                        return contactlink;
                }
                catch { }
            }
            return "No contact link";
        }

        string findEmailFromlink()
        {
            HtmlElementCollection links = webBrowser.Document.Links;
            foreach (HtmlElement l in links)
            {
                try
                {
                    string emaillink = l.GetAttribute("href").ToLower();
                    if (emaillink.Contains("mailto:"))
                    {
                        emaillink = emaillink.Replace("mailto:", "");
                        emaillink = emaillink.Replace("%20", "");
                        if (emaillink.Contains("?subject"))
                            emaillink = emaillink.Remove(emaillink.IndexOf("?subject"));
                        return emaillink;
                    }
                }
                catch { }
            }
            return "No email link";
        }

        string extractEmailfromString( string data)
        {
            //instantiate with this pattern 
            Regex emailRegex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*",
                RegexOptions.IgnoreCase);
            //find items that matches with our pattern
            Match emailMatch = emailRegex.Match(data);
            string emailfoound = emailMatch.Value.ToString();
            if (emailfoound.Contains("?subject"))
                emailfoound = emailfoound.Remove(emailfoound.IndexOf("?subject"));

            return emailfoound;
        }

        string findEmailFromInnerText()
        {
            foreach (HtmlElement span in webBrowser.Document.All)
            {
                try
                {
                    string data = span.InnerText;
                    if (data.Length < 50)
                    {
                        string emailfound = extractEmailfromString(data);
                        if (emailfound != "")
                        {
                            textDebugger.AppendText("Email found from " + webBrowser.Url.ToString() + ": " + emailfound + Environment.NewLine);
                            return emailfound;
                        }
                    }
                }
                catch { }
            }
            return "No email innertext";
        }

        string extractEmailAddressFromPage()
        {
            /* Check if email is in a link */
            string emailFromLink = findEmailFromlink();
            /* check if there is an email in the innertext */
            string emailFromInnerText = findEmailFromInnerText();

            /* If email cannot be found in homepage we find the "Contact Us" page */
            if (emailFromLink != "No email link")
                return emailFromLink;
            else if (emailFromInnerText != "No email innertext")
                return emailFromInnerText;

            return "No email on this page";
        }

        string findEmail(string website)
        {
            if (website.Contains("http://www.yelp.com/"))
                return "";

            navigateAndWait(website);
            string extractedEmail = extractEmailAddressFromPage();

            /* If email cannot be found in homepage we find the "Contact Us" page */
            if (extractedEmail != "No email on this page")
                return extractedEmail + "/" + website;
            else
            {
                string contactsLink = findContactlink();
                string aboutmeLink = findAboutMelink();
                if (contactsLink != "No contact link" )
                {
                    navigateAndWait(contactsLink);
                    extractedEmail = extractEmailAddressFromPage();
                    if (extractedEmail != "No email on this page")
                        return extractedEmail + "/" + contactsLink;

                    return contactsLink;
                }
                else if (aboutmeLink != "No aboutme link")
                {
                    navigateAndWait(aboutmeLink);
                    extractedEmail = extractEmailAddressFromPage();
                    if (extractedEmail != "No email on this page")
                        return extractedEmail + "/" + aboutmeLink;
                }
            }

            return "";
        }

        bool doesYelpCompanyHavePictures()
        {
            foreach (HtmlElement div in webBrowser.Document.All)
            {
                try
                {
                    if (div.GetAttribute("data-media-count") != "")
                        if (Convert.ToInt32(div.GetAttribute("data-media-count")) > 0)
                            return true;
                }
                catch
                {
                    textDebugger.AppendText("Getting yelp media pictures fail: " + webBrowser.Url.ToString() + ": " + div.GetAttribute("data-media-count") + Environment.NewLine);
                }
            }
            return false;
        }

        string findYelpCompanyWebsite(string yelpLink)
        {
            navigateAndWait(yelpLink);
            HtmlElementCollection links = webBrowser.Document.Links;
            foreach (HtmlElement l in links)
            {
                try
                {
                    string companylink = l.GetAttribute("href").ToLower();
                    if (companylink.Contains("/biz_redir"))
                    {
                        string href = l.GetAttribute("href");
                        try
                        {
                            href = href.Substring(href.IndexOf("url=http"));
                            href = href.Remove(href.IndexOf("&src_biz"));
                            href = href.Replace("url=", "");
                            href = System.Web.HttpUtility.UrlDecode(href);
                        }catch{}
                        string companyWebsite = href;
                        textDebugger.AppendText("Yelp website found: " + companyWebsite + Environment.NewLine);
                        return companyWebsite;
                    }
                }
                catch { }
            }
            return "No company website";
        }

        bool checkIfInIgnoreList(string companyName, string companyPhone, string companyWebsite, string companyEmail)
        {
            string _currentList = ignoreList.Text;
            companyWebsite = companyWebsite.Replace("http://", "").Replace("www.", "");
            companyName = companyName == "" ? "NO COMPANY NAME" : companyName;
            companyPhone = companyPhone == "" ? "NO COMPANY PHONE" : companyPhone;
            companyWebsite = companyWebsite == "" ? "NO COMPANY WEBSITE" : companyWebsite;
            companyEmail = companyEmail == "" ? "NO COMPANY EMAIL" : companyEmail;

            if (_currentList.Contains(companyName) || _currentList.Contains(companyPhone)
                || _currentList.Contains(companyWebsite)
                || (_currentList.Contains(companyEmail) & companyEmail != ""))
                return true;
            return false;
        }

        bool checkIfExistInCurrentList(string companyName, string companyPhone, string companyWebsite, string companyEmail)
        {
            if (checkIfInIgnoreList(companyName, companyPhone, companyWebsite, companyEmail))
                return true;
            foreach (DataGridViewRow row in dataCompanyList.Rows)
            {
                if ((string)row.Cells[0].EditedFormattedValue == companyName)
                    return true;  
                if ((string)row.Cells[8].EditedFormattedValue == companyPhone & companyPhone != "")
                    return true;
                if (row.Cells[11].EditedFormattedValue.ToString().Contains(companyWebsite))
                    return true;
                if ((string)row.Cells[9].EditedFormattedValue == companyEmail & companyEmail != "")
                    return true;
            }
            return false;
        }

        void AddtoIgnoreList(string companyName, string companyPhone, string companyWebsite, string companyEmail)
        {
            /*if (companyName != "")
                ignoreList.Text += Environment.NewLine + companyName; */
            if(ignoreList.Text != "" ) ignoreList.AppendText(Environment.NewLine);
            if (companyPhone != "")
                ignoreList.AppendText(companyPhone + Environment.NewLine);
            if (companyEmail != "")
                ignoreList.AppendText(companyEmail + Environment.NewLine);
            if (companyWebsite != "")
                ignoreList.AppendText(companyWebsite + Environment.NewLine);
        }

        void AddtoList(string companyName, string companyAddress, string companyCity, string companyState, string companyZip, 
            string companyCountry, string companyPhone, string companyEmail, string companyContactUs, string companyWebsite, 
            string companyHouzzSearch, List<string> companyCategories)
        {
            companyCountry = companyCountry == "New Zealand" ? "NZ" : companyCountry;
            companyCountry = companyCountry == "Singapore" ? "SG" : companyCountry;
            companyCountry = companyCountry == "United States" ? "US" : companyCountry;
            companyCountry = companyCountry == "Canada" ? "CA" : companyCountry;
            companyCountry = companyCountry == "Australia" ? "AU" : companyCountry;
            if (companyWebsite != null & (companyCountry.ToUpper() == searchCountry.Text.ToUpper() || companyCountry == "")) /* Must perform check alone to see if it's null */
            {
                companyWebsite = formatCompanyWebsite(companyWebsite);
                if (!checkIfExistInCurrentList(companyName, companyPhone, companyWebsite, "No email on 1st check"))
                {
                    if ( !companyEmail.Contains("@") & companyContactUs == "" )
                    { /* If there is no emails passed through parameters then we find email on website */
                        companyEmail = findEmail(companyWebsite);
                        if (companyEmail != "")
                        {
                            if (!companyEmail.Contains("@"))
                            {
                                companyContactUs = companyEmail;  /* If email not found, we'll accept the contact page */
                                companyEmail = "";
                            }
                            else
                            {
                                int splitCharacter = companyEmail.IndexOf("/");
                                companyContactUs = companyEmail.Substring(splitCharacter + 1);
                                companyEmail = companyEmail.Substring(0, splitCharacter);
                            }
                        }
                    }

                    if (!checkDuplicate(xlWorksheet, companyName, companyPhone, companyEmail, companyWebsite))
                    {
                        if (!checkIfExistInCurrentList(companyName, companyPhone, companyWebsite, companyEmail))
                        {
                            /* string duplicateHouzzSearch = checkIfHouzzDuplicate(companyHouzzSearch, companyPhone, companyAddress, companyWebsite);
                            if (duplicateHouzzSearch != "Pro found")
                            {
                                string pro2Url = "";
                                if (duplicateHouzzSearch.Contains("houzz.com/pro2"))
                                    pro2Url = duplicateHouzzSearch;
                            */

                            foreach (string Category in companyCategories)
                            {
                                if (!checkListCategories.Items.Contains(Category))
                                {
                                    checkListCategories.Items.Add(Category);
                                    checkListCategories.SetItemChecked(checkListCategories.Items.Count - 1, true);
                                }
                            }

                                dataCompanyList.Rows.Add(
                                    companyName,
                                    "", /*First Name */
                                    "", /*Last Name */
                                    companyAddress,
                                    companyCity, /*City*/
                                    companyState, /*State*/
                                    companyZip, /*Zip*/
                                    companyCountry.ToUpper(), /* Country */
                                    companyPhone,
                                    companyEmail, /*Email*/
                                    companyContactUs, /*Email Source*/
                                    companyWebsite,
                                    "", /*Pro2 URL */
                                    "", /*Removed company houzz search link*/
                                    "", /* Notes */
                                    String.Join(";",companyCategories) /* Categories*/ 
                                    );
                            /* }
                            // If it's a houzz duplicate, we'll add this company into ignore list 
                            else
                                AddtoIgnoreList(companyName, companyPhone, companyWebsite, companyEmail); */
                        }
                    }
                }
            }
        }


        string getYPList(string searchCriteria, string searchArea, int page)
        {
            string url = "http://pubapi.yp.com/search-api/search/devapi/search?searchloc=" + searchArea
                   + "&term=" + searchCriteria
                   + "&listingcount=50"
                   + "&format=json"
                   + "&sort=distance"
                   + "&pagenum=" + page.ToString()
                   + "&key=" + apiYPKey;

            WebClient request = new WebClient();
            WebHeaderCollection headers = new WebHeaderCollection();
            headers[HttpRequestHeader.UserAgent] = "Mozilla/5.0 (X11; Linux x86_64; rv:10.0) Gecko/20100101 Firefox/10.0 (Chrome)";
            request.Headers = headers;
            var response = request.DownloadString(url);
            request.Dispose();

            return response;
        }

        string getYPListingDetails(string companyListId)
        {
            string url = "http://pubapi.yp.com/search-api/search/devapi/details?"
                            + "listingid=" + companyListId
                            + "&key=" + apiYPKey
                            + "&format=json";

            WebClient request = new WebClient();
            WebHeaderCollection headers = new WebHeaderCollection();
            headers[HttpRequestHeader.UserAgent] = "Mozilla/5.0 (X11; Linux x86_64; rv:10.0) Gecko/20100101 Firefox/10.0 (Chrome)";
            request.Headers = headers;
            var response = request.DownloadString(url);
            request.Dispose();

            return response;
        }

        string getFacebookSearchList(string Criterea, string searchArea, string type, bool switchCriteriaArea)
        {
            string searchQuery = switchCriteriaArea ? searchArea + " " + Criterea : Criterea + " " + searchArea;
            string url = "https://graph.facebook.com/v2.4/search?" 
                            + "access_token=" + apiFacebookKey
                            + "&q=" + searchQuery
                            + "&type=" + type
                            + "&limit=1000"
                            + "&fields=name,id,category,emails,link,website,location,phone,is_unclaimed";

            WebClient request = new WebClient();
            WebHeaderCollection headers = new WebHeaderCollection();
            headers[HttpRequestHeader.UserAgent] = "Mozilla/5.0 (X11; Linux x86_64; rv:10.0) Gecko/20100101 Firefox/10.0 (Chrome)";
            request.Headers = headers;
            var response = request.DownloadString(url);
            request.Dispose();

            return response;
        }

        void populateYPCompanies(RootObjectYP array, Excel._Worksheet xlWorksheet)
        {
            SearchListings searchListings = array.searchResult.searchListings;
                for (int i = 0; i != searchListings.searchListing.Count; i++)
                {
                    string companyName = searchListings.searchListing[i].businessName == null ? "" : searchListings.searchListing[i].businessName;
                    string companyAddress = searchListings.searchListing[i].street == null ? "" : searchListings.searchListing[i].street;
                    string companyState = searchListings.searchListing[i].state == null ? "" : searchListings.searchListing[i].state;
                    string companyCity = searchListings.searchListing[i].city == null ? "" : searchListings.searchListing[i].city;
                    string companyZip = searchListings.searchListing[i].zip == null ? "" : searchListings.searchListing[i].zip;
                    string companyCountry = "US";

                    string companyPhone = searchListings.searchListing[i].phone == null ? "" : fixPhoneNumberFormat(searchListings.searchListing[i].phone);
                    string companyListingId = searchListings.searchListing[i].listingId.ToString();
                    string companyWebsite = "";
                    List<string> companyCategories = searchListings.searchListing[i].categories.category;
                    string companyHouzzSearch = generateHouzzSearchLink(replaceCompanyNamePTE(companyName), companyState, true);
                    /* Search for company details with companyPlaceId */
                    string json = getYPListingDetails(companyListingId);
                    var jarray_Company = JsonConvert.DeserializeObject<RootObjectYPListingId>(json);
                    if (jarray_Company.listingsDetailsResult.listingsDetails.listingDetail[0].extraWebsiteURLs != null)
                    {
                        companyWebsite = jarray_Company.listingsDetailsResult.listingsDetails.listingDetail[0].extraWebsiteURLs.extraWebsiteURL[0];
                        /* Format the website url */
                        companyWebsite = companyWebsite.Remove(0, companyWebsite.LastIndexOf("%2F") + 3);
                        AddtoList(companyName, companyAddress, companyCity, companyState, companyZip, companyCountry, companyPhone, 
                            "No email", "", companyWebsite, companyHouzzSearch, companyCategories);
                    }
                }
        }

        void populateFacebookCompanies(RootObjectFacebook array)
        {
            List<DatumFacebook> searchListings = array.data;
            for (int i = 0; i != searchListings.Count; i++)
            {
                string companyName = searchListings[i].name == null ? "" : searchListings[i].name;
                string companyAddress = "", companyState = "", companyCity = "", companyZip = "", companyCountry = "";
                if (searchListings[i].location != null)
                {
                    companyAddress = searchListings[i].location.street == null ? "" : searchListings[i].location.street;
                    companyState = searchListings[i].location.state == null ? "" : searchListings[i].location.state;
                    companyCity = searchListings[i].location.city == null ? "" : searchListings[i].location.city;
                    companyZip = searchListings[i].location.zip == null ? "" : searchListings[i].location.zip;
                    companyCountry = searchListings[i].location.country == null ? "" : searchListings[i].location.country;
                }
                string companyEmail = searchListings[i].emails == null ? "" : searchListings[i].emails[0];

                string companyPhone = searchListings[i].phone == null ? "" : Regex.Replace(searchListings[i].phone, @"[^0-9\s]", ",");
                while (companyPhone.StartsWith(",")) companyPhone = companyPhone.Substring(1);
                companyPhone = companyPhone.Contains(",") ? companyPhone.Substring(0, companyPhone.IndexOf(",")) : companyPhone;
                string companyWebsite = searchListings[i].website != null ? searchListings[i].website :
                    searchListings[i].is_unclaimed ? null : searchListings[i].link;
                List<string> companyCategories = new List<string>(new string[] {searchListings[i].category});
                string companyHouzzSearch = generateHouzzSearchLink(replaceCompanyNamePTE(companyName), companyState, true);
                AddtoList(companyName, companyAddress, companyCity, companyState, companyZip, companyCountry, companyPhone,
                    companyEmail, "", companyWebsite, companyHouzzSearch, companyCategories);
            }
        }

        string getGoogleAddressComponents(List <AddressComponent> address_components, string component)
        {
            foreach (var item in address_components)
            {
                if (item.types.Contains(component))
                    return item.short_name;
            }
            return "";
        }

        void populateGoogleCompanies(RootObjectPlaces array, Excel._Worksheet xlWorksheet)
        {
            /* Populate the company list */
            for (int i = 0; i != array.results.Count; i++)
            {
                string companyName = array.results[i].name;
                //string companyAddress = array.results[i].formatted_address.Replace(", United States", "");//.Replace(", ", ",");
                string companyPlaceId = array.results[i].place_id;

                /* Search for company details with companyPlaceId */
                string json = getPlaceDetails(companyPlaceId);
                List<string> companyCategories = array.results[i].types;
                companyCategories.RemoveAll(x => x.Equals("point_of_interest"));
                companyCategories.RemoveAll(x => x.Equals("establishment"));
                var jarray_Company = JsonConvert.DeserializeObject<RootObjectCompany>(json);
                if (jarray_Company.status == "OK")
                {
                    string companyPhone = jarray_Company.result.formatted_phone_number == null ? "" : fixPhoneNumberFormat(jarray_Company.result.formatted_phone_number);
                    if (searchCountry.Text != "US")
                        companyPhone = jarray_Company.result.international_phone_number == null ? "" : fixPhoneNumberFormat(jarray_Company.result.international_phone_number);
                    string companyWebsite = jarray_Company.result.website;
                    string companyAddress = getGoogleAddressComponents(jarray_Company.result.address_components, "street_number") + " ";
                    companyAddress += getGoogleAddressComponents(jarray_Company.result.address_components, "route");
                    string companyCity = getGoogleAddressComponents(jarray_Company.result.address_components, "locality");
                    string companyState = getGoogleAddressComponents(jarray_Company.result.address_components, "administrative_area_level_1");
                    string companyZip = getGoogleAddressComponents(jarray_Company.result.address_components, "postal_code");
                    string companyCountry = getGoogleAddressComponents(jarray_Company.result.address_components, "country");
                    string companyHouzzSearch = generateHouzzSearchLink(replaceCompanyNamePTE(companyName),
                        companyState, true);
                    /*Check if duplicate from excel table & houzz website */
                    AddtoList(companyName, companyAddress, companyCity, companyState, companyZip, companyCountry, companyPhone, 
                        "No email", "", companyWebsite, companyHouzzSearch, companyCategories);
                }
                else MessageBox.Show("Error:" + jarray_Company.error_message);
            }
        }

        Excel.Application xlApp;
        Excel.Workbook xlWorkbook = null;
        Excel._Worksheet xlWorksheet;
        
        void SaveSettings()
        {
            Settings.Default.searchCategory = searchCategory.Text;
            Settings.Default.searchCity = searchCity.Text;
            Settings.Default.searchState = searchState.Text;
            Settings.Default.searchCountry = searchCountry.Text;
            Settings.Default.currentList = ignoreList.Text;
            Settings.Default.auditWords = textAuditWords.Text;
            Settings.Default.checkListCategories = new StringCollection();
            foreach (string item in checkListCategories.Items)
                Settings.Default.checkListCategories.Add(item);
            Settings.Default.Save();
        }

        string generateSearchQueryCategory(int categoryline)
        {
            return searchCategory.Lines[categoryline];
        }

        string generateSearchQueryArea(int locationline)
        {
            string searchArea = "";
            searchArea = searchCity.Lines[locationline];
            if (searchState.Text != "")
                searchArea += ", " + searchState.Text;
            return searchArea;
        }

        string generateSearchQuery(int categoryline, int locationline)
        {
            string searchQuery = "";
            searchQuery = generateSearchQueryCategory(categoryline) + " " + generateSearchQueryArea(locationline);
            return searchQuery;
        }
        
        private void buttonSearch_Click(object sender, EventArgs e)
        {
            /* initiate duplicate checker with excel file */
            if (xlWorkbook == null)
            {
                xlApp = new Excel.Application();
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                ofd.ShowDialog();

                xlWorkbook = xlApp.Workbooks.Open(ofd.FileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            }
            /* Save settings */
            SaveSettings();
            checkListCategories.Items.Clear();

            statusBotStrip.Text = "Status: Searching";

            string[] countriesArray = new string[]{"United States", "Singapore", "Australia", "New Zealand"};

            if (searchCity.Lines.Length == 0)
                searchCity.Text = countriesArray[searchCountry.SelectedIndex];

            for (int locationline = 0; locationline < searchCity.Lines.Length; locationline++)
            {
                for (int categoryline = 0; categoryline < searchCategory.Lines.Length; categoryline++)
                {
                    if (searchGoogle.Checked)
                    {
                        /* Conduct first page search */
                        string searchCritera = generateSearchQuery(categoryline, locationline);
                        statusBotStrip.Text = "Status: Searching " + searchCritera + " on Google";
                        searchCritera = searchCritera.Replace(" ", "+");
                        string json = getMapList(searchCritera);
                        var jarray = JsonConvert.DeserializeObject<RootObjectPlaces>(json);

                        if (jarray.status == "OK")
                        {
                            /* Populate the company list */
                            populateGoogleCompanies(jarray, xlWorksheet);

                            /* Conduct next page search until reached end of search */
                            string nextPageToken = jarray.next_page_token;
                            while (nextPageToken != null)
                            {
                                json = getMapList(searchCritera + "&pagetoken=" + nextPageToken);
                                jarray = JsonConvert.DeserializeObject<RootObjectPlaces>(json);
                                /* Populate again */
                                populateGoogleCompanies(jarray, xlWorksheet);
                                nextPageToken = jarray.next_page_token;
                            }
                        }
                        else if (jarray.status == "ZERO_RESULTS")
                            textDebugger.AppendText("Google found no results searching:" + searchCritera + Environment.NewLine);
                        else
                            textDebugger.AppendText("Google Search Error: " + jarray.error_message + " when searching: "
                                + searchCritera + Environment.NewLine);
                    }
                    saveDataSet();  
                    if (searchYelp.Checked)
                    {
                        string searchCriteria = searchCategory.Lines[categoryline];
                        string search_Country = searchCountry.Text;
                        string searchArea = "";
                        searchArea = generateSearchQueryArea(locationline);
                        statusBotStrip.Text = "Status: Searching " + searchArea + " " + searchCriteria + " on Yelp";
                        var yelp = new Yelp(Config.Options);
                        var searchOpt = new SearchOptions();
                        searchOpt.GeneralOptions = new GeneralOptions() { term = searchCriteria, radius_filter = 50000, category_filter = "homeservices", sort = 1 };
                        searchOpt.LocationOptions = new LocationOptions() { location = searchArea };
                        searchOpt.LocaleOptions = new LocaleOptions() { cc = search_Country };
                        var task = yelp.Search(searchOpt);

                        int pages = (int)Math.Ceiling(task.Result.total / 20.0);
                        for (int p = 1; p <= pages; p++)
                        {
                            for (int i = 0; i <= 19; i++)
                            {
                                try
                                {
                                    Business business = task.Result.businesses[i];
                                    string companyName = business.name;
                                    string companyAddress = String.Join(" ", business.location.address);
                                    string companyCity = business.location.city == null ? "" : business.location.city;
                                    string companyState = business.location.state_code == null ? "" : business.location.state_code;
                                    string companyZip = business.location.postal_code == null ? "" : business.location.postal_code;
                                    string companyCountry = business.location.country_code == null ? "" : business.location.country_code;
                                    string companyPhone = business.phone == null ? "" : fixPhoneNumberFormat(business.phone);
                                    if (!checkIfExistInCurrentList(companyName, companyPhone, "No website yet", "No email on 1st check"))
                                    {
                                        string companyWebsite = findYelpCompanyWebsite(task.Result.businesses[i].url);
                                        bool companyHasPicturesOrPersonalWebsite = true;
                                        if (companyWebsite == "No company website")
                                        {
                                            companyHasPicturesOrPersonalWebsite = doesYelpCompanyHavePictures();
                                            companyWebsite = task.Result.businesses[i].url;
                                        }
                                        if (companyHasPicturesOrPersonalWebsite)
                                        {
                                            List<string> companyCategories = new List<string>();
                                            foreach (string[] category in task.Result.businesses[i].categories)
                                                companyCategories.Add(category[0]);
                                            string companyHouzzSearch = generateHouzzSearchLink(replaceCompanyNamePTE(companyName),
                                                searchArea.Remove(0, searchArea.LastIndexOfAny(new char[] { ',', ' ' }) + 1).Replace(" ", ""), true);
                                            AddtoList(companyName, companyAddress, companyCity, companyState, companyZip, companyCountry,
                                                companyPhone, "No email", "", companyWebsite, companyHouzzSearch, companyCategories);
                                        }
                                    }
                                }
                                catch { }
                            }
                            var searchOptions = new SearchOptions();
                            searchOptions.GeneralOptions = new GeneralOptions()
                            {
                                term = searchCriteria,
                                offset = (p * 20),
                                radius_filter = 50000,
                                category_filter = "homeservices",
                                sort = 1 //distance
                            };
                            searchOptions.LocationOptions = new LocationOptions() { location = searchArea };
                            searchOptions.LocaleOptions = new LocaleOptions() { cc = search_Country };
                            task = yelp.Search(searchOptions);
                        }
                    }
                    saveDataSet();
                    if (searchFacebook.Checked)
                    {
                        string searchCriteria = searchCategory.Lines[categoryline];
                        string searchArea = generateSearchQueryArea(locationline);

                        statusBotStrip.Text = "Status: Searching " + searchArea + " " + searchCriteria + " on Facebook";

                        string[] jsonFacebookSearchResults = new string[5];
                        jsonFacebookSearchResults[1] = getFacebookSearchList(searchCriteria, searchArea, "page", false);
                        jsonFacebookSearchResults[2] = getFacebookSearchList(searchCriteria, searchArea, "page", true);
                        jsonFacebookSearchResults[3] = getFacebookSearchList(searchCriteria, searchArea, "place", false);
                        jsonFacebookSearchResults[4] = getFacebookSearchList(searchCriteria, searchArea, "place", true);

                        for (int i = 1; i <= 4; i++)
                        {
                            var jarray = JsonConvert.DeserializeObject<RootObjectFacebook>(jsonFacebookSearchResults[i]);
                            if( jarray.data != null & jarray.data.Count > 0 )
                                populateFacebookCompanies(jarray);
                            else if( jarray.error != null)
                                textDebugger.AppendText("Facebook search error: " + jarray.error.message + Environment.NewLine);
                        }
                    }
                    saveDataSet();
                    if (searchYellowPages.Checked)
                    {
                        /* Conduct first page search */
                        string searchCriteria = searchCategory.Lines[categoryline].Replace(" ", "+");
                        string search_Country = searchCountry.Text;
                        string searchArea = generateSearchQueryArea(locationline);
                        statusBotStrip.Text = "Status: Searching " + searchArea + " " + searchCriteria + " on YellowPages";
                        for (int i = 1; i <= 30; i++) /* 30 pages even if they do not reach 30 */
                        {
                            string json = getYPList(searchCriteria, searchArea, i);
                            var jarray = JsonConvert.DeserializeObject<RootObjectYP>(json);

                            if (jarray.searchResult.metaProperties.errorCode == "")
                            {
                                // Populate the company list 
                                if (jarray.searchResult.metaProperties.listingCount > 0)
                                {
                                    populateYPCompanies(jarray, xlWorksheet);
                                }
                                else
                                {
                                    textDebugger.AppendText("Error on YP page " + i.ToString() + ": No results found searching: " 
                                        + searchArea + " " + searchCriteria + Environment.NewLine);
                                    break;
                                }
                            }
                            else
                                textDebugger.AppendText("Error on YP page " + i.ToString() + ": " + jarray.searchResult.metaProperties.message
                                    + Environment.NewLine);
                        }
                    }
                    saveDataSet();  
                    if( searchFactual.Checked )
                    {
                        string searchCriteria = searchCategory.Lines[categoryline];
                        string search_Country = searchCountry.Text;
                        string search_City = "", search_State = "";
                        Factual factual = new Factual(apiFactualKey, apiFactualSecret);
                        Query q = new Query().SearchExact(searchCriteria);
                        if (search_Country != "US")
                        {
                            q.And(q.Field("country").Equal(search_Country), q.Limit(50));
                            statusBotStrip.Text = "Status: Searching" + search_Country + " " + searchCriteria + " on Factual";
                        }
                        if (search_Country == "US")
                        {
                            q.And(q.Field("country").Equal(search_Country), q.Field("region").Equal(search_State), q.Field("locality").Equal(search_City), q.Limit(50));
                            search_City = searchCity.Lines[locationline];
                            search_State = searchState.Text;
                            statusBotStrip.Text = "Status: Searching " + search_Country + " " + search_City
                                + " " + searchCriteria + " on Factual";
                        }
                            /* if (searchCategoryFilters != "")
                            q.Field("category_labels").Includes(searchCategoryFilters); */
                        int page = 1;
                        try
                        {
                            var json = factual.Fetch("places", q);
                            var jarray = JsonConvert.DeserializeObject<RootObjectFactual>(json);

                            while (jarray.status == "ok" & jarray.response.data.Count > 0)
                            {
                                /* Populate the company list */
                                foreach (Datum item in jarray.response.data)
                                {
                                    string companyName = item.name;
                                    string companyAddress = item.address == null ? "" : item.address;
                                    string companyCity = item.locality == null ? "" : item.locality;
                                    string companyState = item.region == null ? "" : item.region;
                                    string companyCountry = item.country == null ? "" : item.country;
                                    string companyZip = item.postcode == null ? "" : item.postcode;
                                    string companyPhone = item.tel == null ? "" : fixPhoneNumberFormat(item.tel);
                                    string companyEmail = item.email == null ? "no email" : item.email;
                                    string companyWebsite = item.website;
                                    string companyHouzzSearch = generateHouzzSearchLink(replaceCompanyNamePTE(companyName), companyState, true);
                                    List<string> companyCategories = item.category_labels == null ? new List<string>() : item.category_labels[0];
                                    //textDebugger.AppendText(companyName + Environment.NewLine);
                                    AddtoList(companyName, companyAddress, companyCity, companyState, companyZip, companyCountry,
                                        companyPhone, companyEmail, "", companyWebsite, companyHouzzSearch, companyCategories);
                                }
                                page++;
                                int pageOffset = (page - 1) * 50;
                                Query qnew = new Query().SearchExact(searchCriteria);
                                if ( searchCountry.Text == "US" || searchCountry.Text == "AU")
                                {
                                    json = factual.Fetch("places", new Query()
                                        .SearchExact(searchCriteria)
                                        .Field("country").Equal(search_Country)
                                        .Field("region").Equal(search_State)
                                        .Field("locality").Equal(search_City)
                                        .Limit(50)
                                        .Offset(pageOffset));
                                }
                                else if (searchCountry.Text != "US" & searchCountry.Text != "AU")
                                {
                                    json = factual.Fetch("places", new Query()
                                        .SearchExact(searchCriteria)
                                        .Field("country").Equal(search_Country)
                                        .Limit(50)
                                        .Offset(pageOffset));
                                }
                                jarray = JsonConvert.DeserializeObject<RootObjectFactual>(json);
                            }
                        }
                        catch (FactualApiException ex)
                        {
                            textDebugger.AppendText("Factual Requested URL: " + ex.Url + Environment.NewLine);
                            textDebugger.AppendText("Factual Error Status Code: " + ex.StatusCode + Environment.NewLine); ;
                            textDebugger.AppendText("Factual Error Response Message: " + ex.Response + Environment.NewLine); ;
                            if (ex.StatusCode.ToString().Contains("RequestedRangeNotSatisfiable"))
                                textDebugger.AppendText("Factual reached end of results on page " + (page - 1).ToString() + Environment.NewLine);
                        }
                    }
                    saveDataSet();
                    if( searchSensis.Checked )
                    {
                        var searcher = new SsapiSearcher(searchEndPoint, apiSensisKey);
                        // Perform a search and check the response
                        var searchResponse = searcher.SearchFor(searchCategory.Text, searchCity.Text, searchState.Text, 1); /* page 1 */
                        if (searchResponse.code < 200 || searchResponse.code > 299)
                            textDebugger.AppendText("Search failed - Error " + searchResponse.code + ": " + searchResponse.message + Environment.NewLine);
                        else
                        {
                            textDebugger.AppendText("Total results found: " + searchResponse.totalResults.ToString() + Environment.NewLine);
                            textDebugger.AppendText("Total pages: " + searchResponse.totalPages.ToString() + Environment.NewLine);

                            for (int page = 1; page < searchResponse.totalPages; page++)
                            {
                                // Display the results
                                foreach (var result in searchResponse.results)
                                {
                                    string companyWebsite = getSensisContactComponents(result.primaryContacts, "URL");
                                    if (companyWebsite != "")
                                    {
                                        string companyName = result.name;
                                        string companyAddress = "", companyCity = "", companyState = "", companyZip = "";
                                        if (result.primaryAddress != null)
                                        {
                                            companyAddress = result.primaryAddress.addressLine == null ? "" : result.primaryAddress.addressLine;
                                            companyCity = result.primaryAddress.suburb == null ? "" : result.primaryAddress.suburb;
                                            companyState = result.primaryAddress.state == null ? "" : result.primaryAddress.state;
                                            companyZip = result.primaryAddress.postcode == null ? "" : result.primaryAddress.postcode;
                                        }
                                        string companyCountry = "AU";
                                        string companyEmail = getSensisContactComponents(result.primaryContacts, "EMAIL");
                                        string companyPhone = fixPhoneNumberFormat(getSensisContactComponents(result.primaryContacts, "PHONE"));
                                        if (companyPhone == "") companyPhone = fixPhoneNumberFormat(getSensisContactComponents(result.primaryContacts, "MOBILE"));
                                        string companyContactUs = companyEmail.Contains("@") ? "" : getSensisExternalLinksComponents(result.externalLinks, "Contact Us");
                                        string companyHouzzSearch = generateHouzzSearchLink(replaceCompanyNamePTE(companyName), companyState, true);
                                        List<string> companyCategories = result.categories == null ? new List<string>() : new List<string>(new string[] { result.categories[0].name });

                                        //textDebugger.AppendText(companyName + companyAddress + companyCity + companyState + companyZip + companyEmail + companyPhone + companyWebsite + Environment.NewLine);

                                        AddtoList(companyName, companyAddress, companyCity, companyState, companyZip, companyCountry, companyPhone,
                                            companyEmail, companyContactUs, companyWebsite, companyHouzzSearch, companyCategories);
                                    }
                                }
                                searchResponse = searcher.SearchFor(searchCategory.Text, searchCity.Text, searchState.Text, page);
                            }
                        }
                    }
                    /* At the end of each category, we save the data found into an xml file*/
                    saveDataSet();  
                }
            }

            statusBotStrip.Text = "Status: Done";
            /* Save settings */
            SaveSettings();
        }

        private void formMain_Load(object sender, EventArgs e)
        {
           /* dataCompanyList.Rows.Add("Corsi Company Co", "6111 Churchman Bypass, Indianapolis, IN 46203, United States",
                                   "13177861434", "http://www.corsicabinets.com/");
            */
            searchCategory.Text = Settings.Default.searchCategory;
            searchCity.Text = Settings.Default.searchCity;
            ignoreList.Text = Settings.Default.currentList;
            searchState.Text = Settings.Default.searchState;
            searchCountry.Text = Settings.Default.searchCountry;
            textAuditWords.Text = Settings.Default.auditWords;
            if (Settings.Default.checkListCategories != null)
            {
                foreach (string item in Settings.Default.checkListCategories)
                {
                    checkListCategories.Items.Add(item);
                    checkListCategories.SetItemChecked(checkListCategories.Items.Count - 1, true);
                }
            }
            loadDataset();
        }

        string createHttpWWW_Website(string input)
        {
            input = input.Replace("http:///", "").Replace("http://", "").Replace("https://", "").Replace("http:/", "").Replace("http/", "").Replace("http:", "").Replace("http//", "").Replace("www.", "");
            input = "http://www." + input;
            return input;
        }

        //IWebDriver driver;
        private void dataCompanyList_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                var hit = dataCompanyList.HitTest(e.X, e.Y);
                if (hit.ColumnIndex != -1 && hit.RowIndex != -1 && (hit.ColumnIndex == 11 || hit.ColumnIndex == 12 || hit.ColumnIndex == 13))
                {
                    try
                    {
                        string url = dataCompanyList.Rows[hit.RowIndex].Cells[hit.ColumnIndex].Value.ToString();
                        dataCompanyList.Rows[hit.RowIndex].Cells[hit.ColumnIndex].Style.ForeColor = Color.Purple;
                        if (url.Length > 4)
                        {
                            if( languageChinese.Checked )
                                url = "https://translate.google.com/translate?hl=en&sl=auto&tl=zh-TW&u=" + url;
                            /* if (driver != null)
                                driver.Navigate().GoToUrl(createHttpWWW_Website(url)); 
                            else */
                                webBrowser2.Navigate(url);
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void dataCompanyList_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                var hit = dataCompanyList.HitTest(e.X, e.Y);
                if (hit.ColumnIndex != -1 && hit.RowIndex != -1 && (hit.ColumnIndex == 11 || hit.ColumnIndex == 12 || hit.ColumnIndex == 13))
                {
                    this.Cursor = Cursors.Hand;
                }
                else
                {
                    this.Cursor = Cursors.Default;
                }
            }
            catch { }
        }


        private void buttonClear_Click(object sender, EventArgs e)
        {
            DialogResult dlgPrompt = MessageBox.Show("Are you sure you want to clear your entire list?\nSave your work first!",
                "Are you sure?",
                MessageBoxButtons.YesNo);
            if (dlgPrompt == DialogResult.Yes)
            {
                try
                {
                    while (dataCompanyList.Rows.Count > 0)
                    {
                        dataCompanyList.Rows.RemoveAt(0);
                    }
                }
                catch { }
            }
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "Houzz Company List Export (" + DateTime.Now.ToString("MMMM d, yyyy") + ").xls";
            if (sfd.ShowDialog() == DialogResult.OK)
                ToCsV(dataCompanyList, sfd.FileName);
        }

        private void webBrowser_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            textAddressBar.Text = webBrowser.Url.ToString();
        }

        private void textAddressBar_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                webBrowser.Navigate(textAddressBar.Text);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            webBrowser.GoBack();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            webBrowser.GoForward();
        }

        private void dataCompanyList_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try
            {
                /* Row #'s extremely laggy */
                foreach (DataGridViewRow row in dataCompanyList.Rows)
                {
                    row.HeaderCell.Value = (row.Index + 1).ToString();
                }
                dataCompanyList.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            catch { }
            fixDataSetNullsToEmptyStrings();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void dataCompanyList_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            /* Row #'s extremely laggy */
            foreach (DataGridViewRow row in dataCompanyList.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
            dataCompanyList.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void formMain_FormClosing(object sender, FormClosingEventArgs e)
        {
			object misValue = System.Reflection.Missing.Value;
            if( xlWorkbook != null)
                xlWorkbook.Close(false, misValue, misValue);
            SaveSettings();
            saveDataSet();
            try
            {
                if (xlWorkbook != null)
                    xlApp.Quit();
            }
            catch { }
            /* try
            {
                if (driver != null)
                    driver.Quit();
            }
            catch { } */
        }

        private void webBrowser_NewWindow(object sender, CancelEventArgs e)
        {
            try
            {
                webBrowser.Navigate(webBrowser.StatusText);
                e.Cancel = true;
            }
            catch 
            {
                navigateAndWait(webBrowser.StatusText);
                e.Cancel = true;
            }
        }

        private void textAddressBar_Enter(object sender, EventArgs e)
        {
            textAddressBar.SelectAll();
        }

        void finalDuplicateCheck(int start, int end)
        {
            statusBotStrip.Text = "Status: Check Dupes in lead dupe checker excel table";
            auditDuplicatesProgress.Maximum = end;
            auditDuplicatesProgress.Minimum = start;
            auditDuplicatesProgress.Value = start;
            for (int i = start; i <= end & dataCompanyList.Rows[i].Cells[0].Value != null & dataCompanyList.Rows[i].Visible; i++)
            {
                auditDuplicatesProgress.Value = i;
                auditDuplicatesProgress.Refresh();
                double percent = (((double)auditDuplicatesProgress.Value / (double)auditDuplicatesProgress.Maximum) * 100.0);
                try
                {
                    auditDuplicatesProgress.CreateGraphics().DrawString(Convert.ToInt32(percent).ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(auditDuplicatesProgress.Width / 2 - 10, auditDuplicatesProgress.Height / 2 - 7));
                }
                catch { }

                if (dataCompanyList.Rows[i].Visible)
                {
                    if (dataCompanyList.Rows[i].Cells[14].Value == null || dataCompanyList.Rows[i].Cells[14].Value.ToString() != "bad lead")
                    {
                        string companyName = dataCompanyList.Rows[i].Cells[0].Value.ToString();
                        string companyPhone = dataCompanyList.Rows[i].Cells[8].Value == null ? "NO COMPANY PHONE" : dataCompanyList.Rows[i].Cells[8].Value.ToString();
                        string companyEmail = dataCompanyList.Rows[i].Cells[9].Value == null ? "NO COMPANY EMAIL" : dataCompanyList.Rows[i].Cells[9].Value.ToString();
                        string companyWebsite = dataCompanyList.Rows[i].Cells[11].Value == null ? "NO COMPANY WEBSITE" : dataCompanyList.Rows[i].Cells[11].Value.ToString();

                        if (checkDuplicate(xlWorksheet, "", companyPhone, companyEmail, companyWebsite))
                        {
                            dataCompanyList.Rows[i].DefaultCellStyle.BackColor = Color.Maroon;
                            dataCompanyList.Rows[i].Cells[14].Value = "bad lead";
                            dataCompanyList.Rows[i].Cells[12].Value = "Lead Dupe checker found duplicate";
                        }
                        else if (checkIfInIgnoreList(companyName, companyPhone, companyWebsite, companyEmail))
                        {
                            dataCompanyList.Rows[i].DefaultCellStyle.BackColor = Color.Maroon;
                            dataCompanyList.Rows[i].Cells[14].Value = "bad lead";
                            dataCompanyList.Rows[i].Cells[12].Value = "Found in ignore list";
                        }
                    }
                }
            }
            auditDuplicatesProgress.Value = end ;
            statusBotStrip.Text = "Status: Done";
        }

        private void buttonDupCheck_Click(object sender, EventArgs e)
        {
            fixDataSetNullsToEmptyStrings();
            if (xlWorkbook == null)
            {
                xlApp = new Excel.Application();
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                ofd.ShowDialog();

                xlWorkbook = xlApp.Workbooks.Open(ofd.FileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            }

            int checkStart = checkBoxDupeCheckAll.Checked ? 0 : Convert.ToInt32(textDupeCheckRowFrom.Text) - 1;
            int checkEnd = checkBoxDupeCheckAll.Checked ? dataCompanyList.Rows.Count - 1 : Convert.ToInt32(textDupeCheckRowTo.Text) - 1;
            fixDataSetNullsToEmptyStrings();

            finalDuplicateCheck(checkStart, checkEnd);
        }

        private void webBrowser_ProgressChanged(object sender, WebBrowserProgressChangedEventArgs e)
        {
            statusStripLabel.Text = webBrowser.StatusText;
            try
            {
                statusProgressBar.Maximum = (int)e.MaximumProgress;
                statusProgressBar.Value = (int)e.CurrentProgress;
            }
            catch { }
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
        }

        private void webBrowser_LocationChanged(object sender, EventArgs e)
        {
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                /* Row #'s extremely laggy */
                foreach (DataGridViewRow row in dataCompanyList.Rows)
                {
                    string companyName = row.Cells[0].Value.ToString();
                    string companyPhone = row.Cells[8].Value.ToString();
                    string companyEmail = row.Cells[9].Value.ToString();
                    string companyWebsite = row.Cells[11].Value.ToString();
                    if (companyName != "")
                        ignoreList.Text += Environment.NewLine + companyName;
                    if (companyPhone != "")
                        ignoreList.Text += Environment.NewLine + companyPhone;
                    if (companyEmail != "")
                        ignoreList.Text += Environment.NewLine + companyEmail;
                    if (companyWebsite != "")
                        ignoreList.Text += Environment.NewLine + companyWebsite;
                }
            }
            catch { }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ignoreList.SelectAll();
            Clipboard.SetText(ignoreList.Text);
        }


        private void button3_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show(formatCompanyWebsite(textBox1.Text));
        }

        private void button6_Click(object sender, EventArgs e)
        {
            MessageBox.Show("WebBrowser ReadyState: " + webBrowser.ReadyState.ToString());
            MessageBox.Show("WebBrowser busy: " + webBrowser.IsBusy.ToString());
            string webdoc = webBrowser.Document == null ? "" : webBrowser.Document.ToString();
            MessageBox.Show("WebBrowser doc: " + webdoc);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            fixDataSetNullsToEmptyStrings();
            /* Dunno why but must establish connection first by going to website once */
            navigateAndWait("http://waypoint.houzz.net/");
            /* Loop through all data and check if it's a duplicate */
            statusBotStrip.Text = "Status: Check waypoint Dupes";
            int checkStart = checkBoxDupeCheckAll.Checked ? 0 : Convert.ToInt32(textDupeCheckRowFrom.Text) - 1;
            int checkEnd = checkBoxDupeCheckAll.Checked ? dataCompanyList.Rows.Count - 1 : Convert.ToInt32(textDupeCheckRowTo.Text) - 1;

            auditDuplicatesProgress.Maximum = checkEnd;
            auditDuplicatesProgress.Minimum = checkStart;
            auditDuplicatesProgress.Value = checkStart;

            for (int i = checkStart; i <= checkEnd & dataCompanyList.Rows[i].Cells[0].Value != null & dataCompanyList.Rows[i].Visible; i++)
            {
                auditDuplicatesProgress.Value = i;
                auditDuplicatesProgress.Refresh();
                double percent = (((double)auditDuplicatesProgress.Value / (double)auditDuplicatesProgress.Maximum) * 100.0);
                auditDuplicatesProgress.CreateGraphics().DrawString(Convert.ToInt32(percent).ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(auditDuplicatesProgress.Width / 2 - 10, auditDuplicatesProgress.Height / 2 - 7));

                if (dataCompanyList.Rows[i].Cells[0].Value.ToString().Length > 0)
                {
                    if (dataCompanyList.Rows[i].Cells[14].Value == null || dataCompanyList.Rows[i].Cells[14].Value.ToString() != "bad lead")
                    {
                        string companyName = dataCompanyList.Rows[i].Cells[0].Value.ToString();
                        string companyPhone = dataCompanyList.Rows[i].Cells[8].Value == null ? "" : fixPhoneNumberFormat(dataCompanyList.Rows[i].Cells[8].Value.ToString());
                        string companyEmail = dataCompanyList.Rows[i].Cells[9].Value == null ? "" : dataCompanyList.Rows[i].Cells[9].Value.ToString();
                        string companyWebsite = dataCompanyList.Rows[i].Cells[11].Value == null ? "" : dataCompanyList.Rows[i].Cells[11].Value.ToString();

                        if (duplicateCheckV2(companyName, companyPhone, companyWebsite, companyEmail))
                        {
                            dataCompanyList.Rows[i].DefaultCellStyle.BackColor = Color.Maroon;
                            dataCompanyList.Rows[i].Cells[14].Value = "bad lead";
                            dataCompanyList.Rows[i].Cells[12].Value = "Waypoint Dupe checker found duplicate";
                            AddtoIgnoreList(companyName, companyPhone, companyWebsite, companyEmail);
                        }
                    }
                }
            }
            auditDuplicatesProgress.Value = checkEnd;
            statusBotStrip.Text = "Status: Done";
            fixDataSetNullsToEmptyStrings();
            saveDataSet();  
        }

        private void saveCompanies_Click(object sender, EventArgs e)
        {
            fixDataSetNullsToEmptyStrings();
            saveDataSet();  
        }

        private void button8_Click(object sender, EventArgs e)
        {
        }

        private void buttonFrozenFix_Click(object sender, EventArgs e)
        {
            fixFrozen = true;
        }

        string getSensisContactComponents(List<PrimaryContact> primaryContacts, string component)
        {
            foreach (var item in primaryContacts)
            {
                if (item.type.Contains(component))
                    return item.value;
            }
            return "";
        }

        string getSensisExternalLinksComponents(List<ExternalLink> externalLinks, string component)
        {
            if (externalLinks != null)
            {
                foreach (var item in externalLinks)
                {
                    if (item.label.ToLower().Contains(component.ToLower()) || item.displayValue.ToLower().Contains(component.ToLower()))
                        return item.url;
                }
            }
            return "";
        }

        [DllImport("user32.dll")]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

        //Sets window attributes
        [DllImport("USER32.DLL")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        //Gets window attributes
        [DllImport("USER32.DLL")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        //assorted constants needed
        const int WS_BORDER = 8388608;
        const int WS_DLGFRAME = 4194304;
        const int WS_CAPTION = WS_BORDER | WS_DLGFRAME;
        const int WS_SYSMENU = 524288;
        const int WS_THICKFRAME = 262144;
        const int WS_MINIMIZE = 536870912;
        const int WS_MAXIMIZEBOX = 65536;
        const int GWL_STYLE = (int)-16L;
        const int GWL_EXSTYLE = (int)-20L;
        const int WS_EX_DLGMODALFRAME = (int)0x1L;
        const int SWP_NOMOVE = 0x2;
        const int SWP_NOSIZE = 0x1;
        const int SWP_FRAMECHANGED = 0x20;
        const uint MF_BYPOSITION = 0x400;
        const uint MF_REMOVE = 0x1000;

        public void MakeExternalWindowBorderless(IntPtr MainWindowHandle)
        {
            int Style = 0;
            Style = GetWindowLong(MainWindowHandle, GWL_STYLE);
            Style = Style & ~WS_CAPTION;
            Style = Style & ~WS_SYSMENU;
            Style = Style & ~WS_THICKFRAME;
            Style = Style & ~WS_MINIMIZE;
            Style = Style & ~WS_MAXIMIZEBOX;
            SetWindowLong(MainWindowHandle, GWL_STYLE, Style);
            Style = GetWindowLong(MainWindowHandle, GWL_EXSTYLE);
            SetWindowLong(MainWindowHandle, GWL_EXSTYLE, Style | WS_EX_DLGMODALFRAME);
            SetWindowPos(MainWindowHandle, new IntPtr(0), 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED);
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            for (int i = 1; i < 4; i++)
            {
                MessageBox.Show("test");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (ignoreList.Text.Contains(textIgnoreSearch.Text))
                button9.Text = "Found";
            else button9.Text = "Not found.";
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (xlWorkbook == null)
            {
                xlApp = new Excel.Application();
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                ofd.ShowDialog();

                xlWorkbook = xlApp.Workbooks.Open(ofd.FileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            }
            finalDuplicateCheck(Convert.ToInt32(textDupeCheckRowFrom.Text)-1, Convert.ToInt32(textDupeCheckRowFrom.Text)-1);
        }

        private void dataCompanyList_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            try
            {
                if (e.RowIndex == this.dataCompanyList.CurrentCell.RowIndex)
                {
                    e.Paint(dataCompanyList.GetRowDisplayRectangle(e.RowIndex, true), DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);
                    using (Pen p = new Pen(Color.Black, 2))
                    {
                        Rectangle rect = dataCompanyList.GetRowDisplayRectangle(e.RowIndex, true);
                        rect.Width -= 1;
                        rect.Height -= 1;
                        e.Graphics.DrawRectangle(p, rect);
                    }
                    e.Handled = true;
                }
            }
            catch { }
        }

        private void dataCompanyList_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dataCompanyList.Invalidate();
        }

        private void dataCompanyList_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            dataCompanyList.Invalidate();
        }

        private void dataCompanyList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Delete & dataCompanyList.SelectedRows.Count == 0)
            {
                foreach (DataGridViewCell selectedcell in dataCompanyList.SelectedCells)
                {
                    dataCompanyList.CurrentCell.Value = "";
                    selectedcell.Value = "";
                    e.Handled = true;
                }
            }
        }

        bool doMultipleCompanyHouzzSearch(string companyHouzzSearch, string companyPhone, string companyAddress, string companyWebsite, int iRow)
        {
            if (companyHouzzSearch != "none")
            {
                string duplicateHouzzResults = checkIfHouzzDuplicate(companyHouzzSearch, companyPhone, companyAddress, companyWebsite);
                if (duplicateHouzzResults == "Pro found")
                {
                    dataCompanyList.Rows[iRow].Cells[14].Value = "bad lead";
                    dataCompanyList.Rows[iRow].DefaultCellStyle.BackColor = Color.Maroon;
                    dataCompanyList.Rows[iRow].Cells[12].Value = webBrowser.Url.ToString();
                    return true;
                }
                else if (duplicateHouzzResults.Contains("houzz.com/pro2"))
                    dataCompanyList.Rows[iRow].Cells[12].Value = duplicateHouzzResults;
            }
            return false;
        }

        void searchForHouzzDupe(int start, int end)
        {
            fixDataSetNullsToEmptyStrings();
            statusBotStrip.Text = "Status: Check Houzz Pro Dupes";
            auditDuplicatesProgress.Maximum = end;
            auditDuplicatesProgress.Minimum = start;
            auditDuplicatesProgress.Value = start;
            for (int i = start; i <= end & dataCompanyList.Rows[i].Cells[0].Value != null & dataCompanyList.Rows[i].Visible; i++)
            {
                auditDuplicatesProgress.Value = i;
                auditDuplicatesProgress.Refresh();
                double percent = (((double)auditDuplicatesProgress.Value / (double)auditDuplicatesProgress.Maximum) * 100.0);
                auditDuplicatesProgress.CreateGraphics().DrawString(Convert.ToInt32(percent).ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(auditDuplicatesProgress.Width / 2 - 10, auditDuplicatesProgress.Height / 2 - 7));

                if (dataCompanyList.Rows[i].Cells[0].Value.ToString().Length > 0)
                {
                    if (dataCompanyList.Rows[i].Cells[14].Value == null || dataCompanyList.Rows[i].Cells[14].Value.ToString() != "bad lead")
                    {
                        string companyName = dataCompanyList.Rows[i].Cells[0].Value.ToString();
                        string companyPhone = dataCompanyList.Rows[i].Cells[8].Value == null ? "" : 
                            dataCompanyList.Rows[i].Cells[8].Value.ToString() == "" ? "No Phone" : 
                            fixPhoneNumberFormat(dataCompanyList.Rows[i].Cells[8].Value.ToString());
                        string companyWebsite = dataCompanyList.Rows[i].Cells[11].Value == null ? "" : dataCompanyList.Rows[i].Cells[11].Value.ToString();
                        string companyAddress = dataCompanyList.Rows[i].Cells[3].Value == null ? "" : dataCompanyList.Rows[i].Cells[3].Value.ToString();
                        string companyCountry = dataCompanyList.Rows[i].Cells[7].Value == null ? "" : dataCompanyList.Rows[i].Cells[7].Value.ToString();
                        if (companyPhone.StartsWith("64") & companyCountry.ToUpper() == "NZ")
                            companyPhone = companyPhone.Substring(2);
                        string companyLocation = companyCountry == "SG" ? "Singapore" : 
                                                companyCountry == "NZ" ? "New Zealand" : 
                                                dataCompanyList.Rows[i].Cells[6].Value.ToString();
                        string companyEmail = dataCompanyList.Rows[i].Cells[9].Value == null ? "" : dataCompanyList.Rows[i].Cells[9].Value.ToString();
                        const int SEARCH_NUM = 6;
                        string[] companyHouzzSearch = new string[SEARCH_NUM];
                        companyHouzzSearch[0] = generateHouzzSearchLink(formatCompanyWebsite(companyWebsite).Replace("http://", "").Replace("https://", "").Replace("www.", "").Replace("/", "-"), "", false);
                        companyHouzzSearch[1] = generateHouzzSearchLink(companyPhone, "", false);
                        companyHouzzSearch[2] = generateHouzzSearchLink(companyAddress, companyLocation, true);
                        companyHouzzSearch[3] = generateHouzzSearchLink(replaceCompanyNamePTE(companyName).ToLower(), "", true);
                        companyHouzzSearch[4] = generateHouzzSearchLink(replaceCompanyNamePTE(companyName).ToLower(), companyLocation, false);
                        companyHouzzSearch[5] = generateHouzzSearchLink(companyName.ToLower(), companyLocation, false);
                        bool HouzzSearch = false;
                        checkedHouzzProfiles.Clear();
                        for (int searchtimes = 0; searchtimes <= SEARCH_NUM - 1; searchtimes++)
                        {
                            if (HouzzSearch)
                            {
                                AddtoIgnoreList(companyName, companyPhone, companyWebsite, companyEmail);
                                break;
                            }
                            HouzzSearch = doMultipleCompanyHouzzSearch(companyHouzzSearch[searchtimes], companyPhone, companyAddress, companyWebsite, i);
                        }
                    }
                }
            }
            auditDuplicatesProgress.Value = end;
            statusBotStrip.Text = "Status: Done";
        }

        private void buttonHouzzDupeCheck_Click(object sender, EventArgs e)
        {
            int checkStart = checkBoxDupeCheckAll.Checked ? 0 : Convert.ToInt32(textDupeCheckRowFrom.Text) - 1;
            int checkEnd = checkBoxDupeCheckAll.Checked ? dataCompanyList.Rows.Count - 1 : Convert.ToInt32(textDupeCheckRowTo.Text) - 1;
            
            searchForHouzzDupe(checkStart, checkEnd);
            fixDataSetNullsToEmptyStrings();
            saveDataSet();  
        }


        string matchRegex(Regex regex, string data)
        {
            Match match = regex.Match(data);
            string found = match.Value.ToString();
            return found;
        }

        string getPhoneNumberFromWebsite()
        {
            foreach (HtmlElement span in webBrowser.Document.All)
            {
                try
                {
                    string data = span.InnerText;
                    if (data.Length < 50)
                    {
                        //instantiate with this pattern 
                        //Regex phoneRegex = new Regex(@"\(?\b(\d{3})\D?\D?(\d{3})\D?(\d{4})\b", RegexOptions.IgnoreCase);
                        Regex phoneRegex = null;
                        if( searchCountry.Text == "US" )
                            phoneRegex = new Regex(@"\(?\b(\d{3})\D?\D?(\d{3})\D?(\d{4})\b", RegexOptions.IgnoreCase);
                        if (searchCountry.Text == "AU")
                        {
                            data = Regex.Replace(data, @"[^0-9\s]", string.Empty);
                            while (data.StartsWith(" ")) data = data.Remove(0, 1);
                            if (data.Contains(" "))
                                data = data.Split(' ')[0];
                            phoneRegex = new Regex(@"^\({0,1}((0|\+61)(2|4|3|7|8)){0,1}\){0,1}(\ |-){0,1}[0-9]{2}(\ |-){0,1}[0-9]{2}(\ |-){0,1}[0-9]{1}(\ |-){0,1}[0-9]{3}$");
                        }
                        if (searchCountry.Text == "SG")
                        {
                            data = data.Replace(" ", "");
                            phoneRegex = new Regex(@"^(\+|\d)[0-9]{7,16}$", RegexOptions.IgnoreCase);
                        }

                        //find items that matches with our pattern
                        string phonefound = matchRegex(phoneRegex, data);
                        if (phonefound != "" & phonefound != "1800000000" & !phonefound.Contains("%20"))
                        {
                            textDebugger.AppendText("Phone found from " + webBrowser.Url.ToString() + ": " + phonefound + Environment.NewLine);
                            return fixPhoneNumberFormat(phonefound);
                        }
                        else if( phonefound == "" & searchCountry.Text == "AU")
                        {
                            data = data.StartsWith("0") ? data.Remove(0, 1) : "0" + data;
                            phonefound = matchRegex(phoneRegex, data);
                            if (phonefound != "" & phonefound != "1800000000" & !phonefound.Contains("%20"))
                            {
                                textDebugger.AppendText("Phone found from " + webBrowser.Url.ToString() + ": " + phonefound + Environment.NewLine);
                                return fixPhoneNumberFormat(phonefound);
                            }
                        }
                        if( phonefound == "")
                        {
                            /* 1800 numbers */
                            Regex phoneRegex2 = new Regex(@"^[0-9]{10}$|^\(0[1-9]{1}\)[0-9]{8}$|^[0-9]{8}$|^[0-9]{4}[ ][0-9]{3}[ ][0-9]{3}$|^\(0[1-9]{1}\)[ ][0-9]{4}[ ][0-9]{4}$|^[0-9]{4}[ ][0-9]{4}$");
                            phonefound = matchRegex(phoneRegex2, data);
                            if (phonefound != "" & phonefound != "1800000000" & !phonefound.Contains("%20"))
                            {
                                textDebugger.AppendText("Phone found from " + webBrowser.Url.ToString() + ": " + phonefound + Environment.NewLine);
                                return fixPhoneNumberFormat(phonefound);
                            }
                        }
                    }
                }
                catch { }
            }
            return "No phone number";
        }

        private void buttonFindAllPhones_Click(object sender, EventArgs e)
        {
            statusBotStrip.Text = "Status: Finding all phones";
            for (int i = 0; i < dataCompanyList.Rows.Count - 1 & dataCompanyList.Rows[i].Cells[11].Value != null; i++)
            {
                navigateAndWait(dataCompanyList.Rows[i].Cells[11].Value.ToString());
                string phoneFound = getPhoneNumberFromWebsite();
                if (phoneFound != "No phone number")
                    dataCompanyList.Rows[i].Cells[8].Value = phoneFound;
                else if (phoneFound == "No phone number")
                {
                    string contactsLink = findContactlink();
                    string aboutmeLink = findAboutMelink();
                    if (contactsLink != "No contact link")
                    {
                        navigateAndWait(contactsLink);
                        phoneFound = getPhoneNumberFromWebsite();
                        if (phoneFound != "No phone number")
                            dataCompanyList.Rows[i].Cells[8].Value = phoneFound;
                        else if (aboutmeLink != "No aboutme link")
                        {
                            navigateAndWait(aboutmeLink);
                            phoneFound = getPhoneNumberFromWebsite();
                            if (phoneFound != "No phone number")
                                dataCompanyList.Rows[i].Cells[8].Value = phoneFound;
                        }
                    }
                }
            }
            statusBotStrip.Text = "Status: Done";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            searchForHouzzDupe(Convert.ToInt32(textDupeCheckRowFrom.Text) - 1, Convert.ToInt32(textDupeCheckRowFrom.Text) - 1);
        }

        string checkIfPriorityHouzzDuplicate(string companyHouzzSearch, string companyPhone, string companyWebsite)
        {
            navigateAndWait(companyHouzzSearch);

            List<string> proUrls = new List<string>();
            List<string> pro2Urls = new List<string>();

            HtmlElementCollection links = webBrowser.Document.Links;
            foreach (HtmlElement l in links)
            {
                try
                {
                    string linkFound = l.GetAttribute("href");
                    if (linkFound.Contains("www.houzz.com/pro/"))
                    {
                        if (!proUrls.Contains(linkFound))
                            proUrls.Add(linkFound);
                    }
                    if (linkFound.Contains("www.houzz.com/pro2/"))
                    {
                        if (!pro2Urls.Contains(linkFound))
                            pro2Urls.Add(linkFound);
                    }
                }
                catch { }
            }


            string result = "No pro found";
            proUrls.ForEach(delegate(string linkUrl)
            {
                if (!result.Contains("http://www.houzz.com/pro/"))
                {
                    navigateAndWait(linkUrl);
                    string phoneFound = getPhoneNumber();
                    string websiteFound = getWebsite();
                    if (phoneFound.EndsWith(companyPhone) || phoneFound == companyPhone)
                    {
                        textDebugger.AppendText("Found Houzz Dupe " + phoneFound + " equals " + companyPhone + Environment.NewLine);
                        result = linkUrl;
                    }
                    if (formatCompanyWebsite(companyWebsite).Contains(websiteFound) & websiteFound.Length >= 5)
                    {
                        textDebugger.AppendText("Found Houzz Dupe " + websiteFound + " within " + formatCompanyWebsite(companyWebsite) + Environment.NewLine);
                        result = linkUrl;
                    }
                }
            });

            if (result == "No pro found")
            {
                pro2Urls.ForEach(delegate(string linkUrl)
                {
                    if (!result.Contains("www.houzz.com/pro2/"))
                    {
                        navigateAndWait(linkUrl);
                        string phoneFound = getPhoneNumber();
                        string websiteFound = getWebsite();
                        if (phoneFound == companyPhone || (companyWebsite.Contains(websiteFound) & websiteFound.Length >= 5))
                            result = linkUrl;
                    }
                });
            }

            return result;
        }

        bool doMultiplePriorityCompanyHouzzSearch(string companyHouzzSearch, string companyPhone, string companyWebsite, int iRow)
        {
            if (companyHouzzSearch != "none")
            {
                string duplicateHouzzResults = checkIfPriorityHouzzDuplicate(companyHouzzSearch, companyPhone, companyWebsite);
                if (duplicateHouzzResults.Contains("houzz.com/pro/"))
                {
                    dataCompanyList.Rows[iRow].Cells[12].Value = "Yes";
                    dataCompanyList.Rows[iRow].Cells[13].Value = duplicateHouzzResults;
                    return true;
                }
                else if (duplicateHouzzResults.Contains("houzz.com/pro2"))
                {
                    dataCompanyList.Rows[iRow].Cells[12].Value = "No";
                    dataCompanyList.Rows[iRow].Cells[13].Value = duplicateHouzzResults;
                }
                else
                {
                    dataCompanyList.Rows[iRow].Cells[12].Value = "No";
                }
            }
            return false;
        }

        void PRIORITYsearchForHouzzDupe(int start, int end)
        {
            statusBotStrip.Text = "Status: PRIORITY Check Houzz Pro Dupes";
            for (int i = start; i <= end & dataCompanyList.Rows[i].Cells[0].Value != null; i++)
            {
                if (dataCompanyList.Rows[i].Cells[12].Value.ToString() != "YES" & dataCompanyList.Rows[i].Cells[12].Value.ToString() != "Yes")
                {
                    string companyName = dataCompanyList.Rows[i].Cells[0].Value.ToString();
                    string companyPhone = dataCompanyList.Rows[i].Cells[8].Value.ToString() == "" ? "No Phone" : dataCompanyList.Rows[i].Cells[8].Value.ToString();
                    string companyWebsite = dataCompanyList.Rows[i].Cells[11].Value.ToString();
                    const int SEARCH_NUM = 6;
                    string[] companyHouzzSearch = new string[SEARCH_NUM];
                    companyHouzzSearch[0] = generateHouzzSearchLink(formatCompanyWebsite(companyWebsite).Replace("http://", "").Replace("https://", "").Replace("www.", "").Replace("/", "-"), "", false);
                    companyHouzzSearch[1] = generateHouzzSearchLink(companyPhone, "", false);
                    companyHouzzSearch[3] = generateHouzzSearchLink(replaceCompanyNamePTE(companyName).ToLower(), "", true);
                    companyHouzzSearch[5] = generateHouzzSearchLink(companyName.ToLower(), "", false);
                    bool HouzzSearch = false;
                    for (int searchtimes = 0; searchtimes <= SEARCH_NUM - 1; searchtimes++)
                    {
                        if (HouzzSearch)
                            break;
                        HouzzSearch = doMultiplePriorityCompanyHouzzSearch(companyHouzzSearch[searchtimes], companyPhone, companyWebsite, i);
                    }
                }
            }
            statusBotStrip.Text = "Status: Done";
        }

        private void button12_Click(object sender, EventArgs e)
        {
            PRIORITYsearchForHouzzDupe(0, (dataCompanyList.Rows.Count - 1));
        }

        private void dataCompanyList_Sorted(object sender, EventArgs e)
        {
            try
            {
                /* Row #'s extremely laggy */
                foreach (DataGridViewRow row in dataCompanyList.Rows)
                {
                    row.HeaderCell.Value = (row.Index + 1).ToString();
                }
                dataCompanyList.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            catch { }
        }

        private void buttonFindAllEmails_Click(object sender, EventArgs e)
        {
            statusBotStrip.Text = "Status: Finding all emails";
            for (int i = 0; i < dataCompanyList.Rows.Count - 1 & dataCompanyList.Rows[i].Cells[11].Value != null; i++)
            {
                string companyWebsite = dataCompanyList.Rows[i].Cells[11].Value.ToString();
                navigateAndWait(companyWebsite);
                string companyEmail = findEmail(companyWebsite);
                string companyContactUs = "";
                if (companyEmail != "")
                {
                    if (!companyEmail.Contains("@"))
                    {
                        companyContactUs = companyEmail;  /* If email not found, we'll accept the contact page */
                        companyEmail = "";
                    }
                }
                dataCompanyList.Rows[i].Cells[9].Value = companyEmail;
                dataCompanyList.Rows[i].Cells[10].Value = companyContactUs;
                saveDataSet();
            }
            saveDataSet();
            statusBotStrip.Text = "Status: Done";
        }

        private void PasteClipboard()
        {
            try
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');
                int iRow = dataCompanyList.CurrentCell.RowIndex;
                int iCol = dataCompanyList.CurrentCell.ColumnIndex;
                DataGridViewCell oCell;
                foreach (string line in lines)
                {
                    if (iRow < dataCompanyList.RowCount && line.Length > 0)
                    {
                        List<string> sCells = line.Split('\t').ToList<string>();
                        if (dataCompanyList.RowCount <= lines.GetLength(0))
                            dataCompanyList.Rows.Add();
                        for (int i = 0; i < sCells.Count; ++i)
                        {
                            if (iCol + i < this.dataCompanyList.ColumnCount)
                            {
                                oCell = dataCompanyList[iCol + i, iRow];
                                if (!oCell.ReadOnly)
                                {
                                    if (oCell.Value == null || oCell.Value.ToString() != sCells[i])
                                    {
                                        oCell.Value = Convert.ChangeType(sCells[i], oCell.ValueType) == null ? "" : Convert.ChangeType(sCells[i], oCell.ValueType);
                                        //oCell.Style.BackColor = Color.Tomato;
                                    }
                                }
                            }
                        }
                        iRow++;
                    }
                    else
                    { break; }
                }
            }
            catch (FormatException)
            {
                MessageBox.Show("The data you pasted is in the wrong format for the cell");
                return;
            }
        }

        private void buttonImportData_Click(object sender, EventArgs e)
        {
            /* 
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
            ofd.ShowDialog();

            //MessageBox.Show(ofd.FileName);
            String name = "Items";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            ofd.FileName +
                            ";Extended Properties=\"Excel 8.0 Xml;HDR=NO;IMEX=1;\"";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            dataCompanyList.DataSource = data; */

            PasteClipboard();
        }

        private void dataCompanyList_KeyUp(object sender, KeyEventArgs e)
        {
             if(e.Control && e.KeyCode == System.Windows.Forms.Keys.C){
				 Clipboard.SetDataObject(dataCompanyList.GetClipboardContent());
			 }
             else if (e.Control && e.KeyCode == System.Windows.Forms.Keys.V)
             {
                 if( dataCompanyList.SelectedCells.Count == 1)
                    PasteClipboard();
                 else if (dataCompanyList.SelectedCells.Count > 1)
                 {
                     foreach (DataGridViewCell selectedcell in dataCompanyList.SelectedCells)
                     {
                         selectedcell.Value = Clipboard.GetText();
                     }
                 }
             }
        }

        bool checkIfEmailExistInPage(string Email)
        {
            try
            {
                string data = webBrowser.Document.Body.InnerHtml.ToLower();
                if (data.Contains(Email))
                    return true;
            }
            catch { }
            foreach (HtmlElement span in webBrowser.Document.All)
            {
                try
                {
                    string data = span.InnerText.ToLower();
                    if (data.Length < 50)
                    {
                        if (data.Contains(Email))
                        {
                            textDebugger.AppendText("Matched emails: " + data + "(extracted) and " + Email + "(datalist)" + Environment.NewLine);
                            return true;
                        }
                    }
                }
                catch { }
            }

            HtmlElementCollection links = webBrowser.Document.Links;
            foreach (HtmlElement l in links)
            {
                try
                {
                    string emaillink = l.GetAttribute("href").ToLower();
                    if (emaillink.Contains(Email))
                    {
                        textDebugger.AppendText("Matched emails: " + emaillink + "(extracted) and " + Email + "(datalist)" + Environment.NewLine);
                        return true;
                    }
                }
                catch { }
            }
            return false;
        }

        bool checkIfPhoneExistInPage(string Phone)
        {
            try
            {
                string data = webBrowser.Document.Body.InnerHtml.ToLower();
                data = Regex.Replace(data, @"[^0-9\s]", string.Empty);
                if (data.Contains(Phone))
                    return true;
            }
            catch { }
            foreach (HtmlElement span in webBrowser.Document.All)
            {
                try
                {
                    string data = span.InnerHtml;
                    data = Regex.Replace(data, @"[^0-9]", string.Empty);
                    if (data.Length < 40)
                    {
                        if (data.Contains(Phone))
                        {
                            //textDebugger.AppendText("Matched Phones: " + data + "(extracted) and " + Phone + "(datalist)" + Environment.NewLine);
                            return true;
                        }
                    }
                }
                catch { }
            }
            return false;
        }

        bool checkIfAddressExistInPage(string Address)
        {
            try
            {
                string data = webBrowser.Document.Body.InnerHtml.ToLower();
                if (data.Contains(Address.ToLower()))
                {
                    //textDebugger.AppendText("Matched addresses:" + data + "(extracted) and " + Address + "(datalist)" + Environment.NewLine);
                    return true;
                }
            }
            catch { }
            foreach (HtmlElement span in webBrowser.Document.All)
            {
                try
                {
                    string data = span.InnerHtml.ToLower();
                    if (data.Length < 50)
                    {
                        if (data.Contains(Address.ToLower()) )
                        {
                            textDebugger.AppendText("Matched addresses:" + data + "(extracted) and " + Address + "(datalist)" + Environment.NewLine);
                            return true;
                        }
                    }
                }
                catch { }
            }
            return false;
        }

        bool checkIfKeywordExistInPage(string Keyword)
        {
            try
            {
                string data = webBrowser.Document.Body.InnerHtml.ToLower();
                if (data.Contains(Keyword.ToLower()))
                {
                    //textDebugger.AppendText("Matched addresses:" + data + "(extracted) and " + Address + "(datalist)" + Environment.NewLine);
                    return true;
                }
            }
            catch { }
            foreach (HtmlElement span in webBrowser.Document.All)
            {
                try
                {
                    string data = span.InnerHtml.ToLower();
                    if (data.Contains(Keyword.ToLower()))
                    {
                        if( data.Length < 50 )
                            textDebugger.AppendText("Matched addresses:" + data + "(extracted) and " + Keyword + "(datalist)" + Environment.NewLine);
                        return true;
                    }
                }
                catch { }
            }
            return false;
        }

        private void buttonAuditCompanyInfo_Click(object sender, EventArgs e)
        {
            fixDataSetNullsToEmptyStrings();
            statusBotStrip.Text = "Status: Auditing Company Info";
            int checkStart = checkBoxAuditCompanyAll.Checked ? 0 : Convert.ToInt32(textCompAuditRowFrom.Text) - 1;
            int checkEnd = checkBoxAuditCompanyAll.Checked ? dataCompanyList.Rows.Count - 1 : Convert.ToInt32(textCompAuditRowTo.Text) - 1;

            auditCompanyInfoProgress.Maximum = checkEnd;
            auditCompanyInfoProgress.Minimum = checkStart;
            auditCompanyInfoProgress.Value = checkStart;
            for (int i = checkStart; i <= checkEnd & dataCompanyList.Rows[i].Cells[11].Value != null; i++)
            {
                auditCompanyInfoProgress.Value = i;
                int percent = (int)(((double)auditCompanyInfoProgress.Value / (double)auditCompanyInfoProgress.Maximum) * 100.0);
                auditCompanyInfoProgress.CreateGraphics().DrawString(percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(auditCompanyInfoProgress.Width / 2 - 10, auditCompanyInfoProgress.Height / 2 - 7));

                /* Get Company information */
                string companyWebsite = dataCompanyList.Rows[i].Cells[11].Value.ToString();
                string companyPhone = dataCompanyList.Rows[i].Cells[8].Value == null ? "" : dataCompanyList.Rows[i].Cells[8].Value.ToString();
                companyPhone = companyPhone.StartsWith("0") ? companyPhone.Remove(0, 1) : companyPhone;
                companyPhone = companyPhone.Replace(" ", "");
                string companyEmail = dataCompanyList.Rows[i].Cells[9].Value == null ? "" : dataCompanyList.Rows[i].Cells[9].Value.ToString().ToLower();
                string companyAddress = dataCompanyList.Rows[i].Cells[3].Value == null ? "" : dataCompanyList.Rows[i].Cells[3].Value.ToString().ToLower();
                string companyCountry = dataCompanyList.Rows[i].Cells[7].Value == null ? "" : dataCompanyList.Rows[i].Cells[7].Value.ToString().ToLower();
                

                bool matchedEmail = auditEmail.Checked ? false : true,
                     matchedPhone = auditPhone.Checked ? false : true,
                     matchedAddress = auditAddress.Checked ? false : true,
                     matchedKeyword = auditKeyword.Checked ? false: true;

                /* Navigate to the current checking website homepage */
                navigateAndWait(companyWebsite);
                string contactsLink = findContactlink();
                string aboutmeLink = findAboutMelink();
                string servicesLink = findLink("service");

                for (int pageCount = 1; pageCount <= 4 & (!matchedPhone || !matchedEmail || !matchedAddress || !matchedKeyword) ; pageCount++)
                {
                    if (pageCount == 2)
                    {
                        if (contactsLink == "No contact link" || contactsLink.Contains("mailto:") )
                            pageCount = 3;
                        else navigateAndWait(contactsLink);
                    }
                    if (pageCount == 3 || aboutmeLink.Contains("mailto:") )
                    {
                        if (aboutmeLink == "No aboutme link")
                            break;
                        else navigateAndWait(aboutmeLink);
                    }
                    if( pageCount == 4 & !matchedKeyword)
                    {
                        if (servicesLink == "No link")
                            break;
                        else navigateAndWait(servicesLink);
                    }
                    /* If there is a phone number for the current company we'll check for a phone from the website to compare */
                    if (companyPhone != "" & auditPhone.Checked & !matchedPhone)
                    {
                        if (companyCountry != "NZ" & companyPhone.StartsWith("64"))
                            companyPhone = companyPhone.Substring(2);
                        if (checkIfPhoneExistInPage(companyPhone))
                        {
                            dataCompanyList.Rows[i].Cells[8].Style.BackColor = Color.Yellow;
                            matchedPhone = true;
                        }
                    }
                
                    /* If there is an email for the current company we'll check for an email from the website to compare */
                    if (companyEmail != "" & auditEmail.Checked & !matchedEmail)
                    {
                        if( checkIfEmailExistInPage(companyEmail) )
                        {
                            dataCompanyList.Rows[i].Cells[9].Style.BackColor = Color.Yellow;
                            matchedEmail = true;
                        }
                    }

                    /* If there is an address for the current company we'll check for an address from the website to compare */
                    if( companyAddress.Length >= 4 & auditAddress.Checked & !matchedAddress)
                    {
                        companyAddress = companyAddress.Replace("street", "st");
                        companyAddress = companyAddress.Replace("road", "r");
                        companyAddress = companyAddress.Replace("drive", "dr");
                        if( checkIfAddressExistInPage(companyAddress) )
                        {
                            dataCompanyList.Rows[i].Cells[3].Style.BackColor = Color.Yellow;
                            matchedAddress = true;
                        }
                    }

                    /* If searching for a keyword */
                    if ( auditKeyword.Checked & !matchedKeyword)
                    {
                        List<string> keywords = textAuditWords.Text.Split(',').ToList<string>();
                        int totalKeywords = keywords.Count;
                        int foundKeywords = 0;
                        foreach (string keyword in keywords)
                        {
                            if (checkIfKeywordExistInPage(keyword))
                            {
                                if (dataCompanyList.Rows[i].Cells[14].Value == null || !dataCompanyList.Rows[i].Cells[14].Value.ToString().Contains(keyword))
                                {
                                    dataCompanyList.Rows[i].Cells[14].Value += keyword + ",";
                                    foundKeywords++;
                                    if (foundKeywords == totalKeywords)
                                    {
                                        dataCompanyList.Rows[i].Cells[14].Style.BackColor = Color.Yellow;
                                        matchedKeyword = true;
                                    }   
                                }
                            }
                        }
                    }
                }
            }
            auditCompanyInfoProgress.Value = checkEnd;
            statusBotStrip.Text = "Status: Done";
        }
        /*
        public void Wait(double delay, double interval)
        {
            // Causes the WebDriver to wait for at least a fixed delay
            var now = DateTime.Now;
            var wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(delay));
            wait.PollingInterval = TimeSpan.FromMilliseconds(interval);
            wait.Until(wd => (DateTime.Now - now) - TimeSpan.FromMilliseconds(delay) > TimeSpan.Zero);
        } */

        private void button10_Click_1(object sender, EventArgs e)
        {
            //!Make sure to add the path to where you extracting the chromedriver.exe:
           /*
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            ChromeOptions options = new ChromeOptions();
            //options.AddArgument("start-minimized");

            driver = new ChromeDriver(driverService, options);
            //IWebDriver driver = new ChromeDriver(Directory.GetCurrentDirectory()); //<-Add your path
            string title = String.Format("{0} - Google Chrome", driver.Title);
            var process = Process.GetProcesses()
                .FirstOrDefault(x => x.MainWindowTitle == title);

            if (process != null)
            {
                SetParent(process.MainWindowHandle, tabChromeBrowser.Handle);
                MakeExternalWindowBorderless(process.MainWindowHandle);
                ShowWindow(process.MainWindowHandle, 3);
                SetWindowPos(process.MainWindowHandle, IntPtr.Zero, 0, 0, 2000, 2000, 0);
            }
            driver.Navigate().GoToUrl("http://www.google.com"); */
        }

        private void buttonFilterCategories_Click(object sender, EventArgs e)
        {
            dataCompanyList.CurrentCell = null;
            foreach (DataGridViewRow Row in dataCompanyList.Rows)
            {
                try
                {
                    Row.Visible = true;
                }
                catch { }
                bool Hide = true;
                foreach (string category in checkListCategories.CheckedItems)
                {
                    if (Hide)
                    {
                        string categories = Row.Cells[15].Value == null ? "" : Row.Cells[15].Value.ToString();
                        foreach (string rowCategory in categories.Split(';'))
                        {
                            if (rowCategory == category)
                            {
                                Hide = false;
                                break;
                            }
                        }
                    }
                }
                if (Hide)
                {
                    try
                    {
                        Row.Visible = false;
                    }
                    catch { }
                }
            }
        }
    }
}
