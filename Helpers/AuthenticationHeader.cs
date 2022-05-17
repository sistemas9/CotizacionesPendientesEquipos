using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

namespace CotizacionesPendientesEquipos
{
  public class AuthenticationHeader
    {
        //private static string ODataEntityPath = "https://tes-ayt.sandbox.operations.dynamics.com/Data/";
        //private static Uri oDataUri = new Uri(ODataEntityPath, UriKind.Absolute);
        //private static Resources context = new Resources(oDataUri);
        public static String clientId;
        public static String clientSecretId;
        public static String urlBase;

        public async Task<String> getAuthenticationHeader()
        {
          var AppSettings = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(@"appsettings.json"));
          String ambiente = AppSettings["Config"];
          String company = AppSettings["Company"];

          String token = await GetToken(ambiente,company);
          return token;
        }

          public async Task<String> GetToken(String ambiente, String company)
          {
            String url = String.Empty;
            switch (company) {
            case "ATP":
              if (ambiente == "DESARROLLO")
              {
                url = "https://solutiontinaxdev.azurewebsites.net/SolutionToken/api/SolutionToken";
              }
              else
              {
                url = "https://solutiontinax-solutiontokeninaxpr.azurewebsites.net/SolutionToken/api/SolutionToken";
              }
              break;
            case "LIN":
              if (ambiente == "DESARROLLO")
              {
                url = "https://solutiontinaxdev.azurewebsites.net/SolutionTokenLIN/api/SolutionToken";
              }
              else
              {
                url = "https://solutiontinaxprod.azurewebsites.net/SolutionTokenLID/api/SolutionToken";
              }
              break;
            case "INN":
              if (ambiente == "DESARROLLO")
              {
                url = "https://solutiontinaxdev.azurewebsites.net/SolutionTokenINN/api/SolutionTokenINN";
              }
              else
              {
                url = "https://solutiontinaxprod.azurewebsites.net/SolutionTokenINN/api/SolutionTokenINN";
              }
              break;
      }
          /////////////////////////////////////generar token/////////////////////////////////////////////////////////////////////////////
          token authenticationHeader = new token();
          System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls;
          HttpClient token = new HttpClient();
          //HttpResponseMessage response = await token.GetAsync(url);
          HttpResponseMessage response = await GetTokenReponseAsync(url,token);
          response.EnsureSuccessStatusCode();
          String responseBody = await response.Content.ReadAsStringAsync();
          responseBody = responseBody.Substring(1, responseBody.Length - 2);
          authenticationHeader = JsonConvert.DeserializeObject<token>(responseBody);
          ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
          return authenticationHeader.Token;
    }
    public async Task<HttpResponseMessage> GetTokenReponseAsync(String url, HttpClient token)
    {
      return token.GetAsync(url).Result;
    }
  }

    public class token
    {
      public String Token { get; set; }
    }
}