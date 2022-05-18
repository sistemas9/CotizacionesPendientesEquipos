using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using iTextSharp.tool.xml.html;
using iTextSharp.tool.xml.parser;
using iTextSharp.tool.xml.pipeline.css;
using iTextSharp.tool.xml.pipeline.end;
using iTextSharp.tool.xml.pipeline.html;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CotizacionesPendientesEquipos
{
  class Program
  {
    public static EmployeesResult workers;
    static void Main(string[] args)
    {
      DateTime inicioGetList = DateTime.Now;
      Console.WriteLine("Inicio obtencion de cotizaciones: {0}",inicioGetList.ToString());
      SalesQuotationResult salesQuotationResult = GetListOfQuoations();
      workers = GetListOfEmployees();
      DateTime finGetList = DateTime.Now;
      Console.WriteLine("Fin de obtencion de cotizaciones: {0}", finGetList.ToString());
      Console.WriteLine("Total de tiempo en obtencion: {0}",((finGetList - inicioGetList).TotalSeconds).ToString());
      Console.WriteLine("------------------------------------------------");
      List<SalesQuotation> EquipoOrRefaccionSalesQuotation = new List<SalesQuotation>();
      DateTime inicioEqOrRef = DateTime.Now;
      Console.WriteLine("Inicio de equipo or refaccion: {0}",inicioEqOrRef.ToString());
      //foreach (SalesQuotation quotation in salesQuotationResult.value)
      Parallel.ForEach(salesQuotationResult.value, (quotation) =>
      {
        //foreach (SalesQuotationLines lines in quotation.SalesQuotationLines)
        Parallel.ForEach(quotation.SalesQuotationLines, (lines,state) =>
        {
          Boolean isEquipoOrRefaccion = IsEquipoOrRefaccion(lines.ItemNumber);
          if (isEquipoOrRefaccion)
          {
            EquipoOrRefaccionSalesQuotation.Add(quotation);
            state.Break();
          }
        });
      });
      DateTime finEqOrRef = DateTime.Now;
      Console.WriteLine("Fin equipo or refaccion : {0}",finEqOrRef.ToString());
      Console.WriteLine("Total de tiempo en equipo or refaccion: {0}", ((finEqOrRef - inicioEqOrRef).TotalMinutes).ToString());
      Console.WriteLine("Total de cotizaciones: {0}", EquipoOrRefaccionSalesQuotation.Count);
      EquipoOrRefaccionSalesQuotation = EquipoOrRefaccionSalesQuotation.Distinct().ToList();
      EquipoOrRefaccionSalesQuotation = EquipoOrRefaccionSalesQuotation.OrderBy(cot => cot.DefaultShippingSiteId).ThenBy(cot2 => cot2.SalesQuotationNumber).ToList();
      String template = GetTemplateHTML(EquipoOrRefaccionSalesQuotation);
      StringReader sr = new StringReader(template);
      using (MemoryStream ms2 = new MemoryStream())
      {
        Document document = new Document(PageSize.LETTER, 15, 15, 15, 15);
        PdfWriter writer2 = PdfWriter.GetInstance(document, ms2);
        document.Open();
        HtmlPipelineContext htmlContext = new HtmlPipelineContext(new CssAppliersImpl());
        htmlContext.SetTagFactory(Tags.GetHtmlTagProcessorFactory());
        ICSSResolver cssResolver = XMLWorkerHelper.GetInstance().GetDefaultCssResolver(true);
        IPipeline pipeline = new CssResolverPipeline(cssResolver, new HtmlPipeline(htmlContext, new PdfWriterPipeline(document, writer2)));
        XMLWorker worker = new XMLWorker(pipeline, true);
        XMLParser p = new XMLParser(true, worker, Encoding.UTF8);
        try
        {
          p.Parse(stringToStream(template));
        }
        catch (Exception e)
        {
          Console.WriteLine(e.Message);
        }
        document.Close();
        writer2.Close();
        Mail correo = new Mail();
        correo.SendMail(template, ms2);
      }
      Console.WriteLine("Fin...");
    }

    public static Stream stringToStream(String txt)
    {
      var stream = new MemoryStream();
      var w = new StreamWriter(stream);
      w.Write(txt);
      w.Flush();
      stream.Position = 0;
      return stream;
    }

    public static SalesQuotationResult GetListOfQuoations()
    {
      SalesQuotationResult SalesQuotations = new SalesQuotationResult();
      ConsultaEntity entity = new ConsultaEntity();
      String url = "https://ayt.operations.dynamics.com/Data/SalesQuotationHeadersV2?$filter=SalesQuotationStatus eq Microsoft.Dynamics.DataEntities.SalesQuotationStatus'Created' and ReceiptDateRequested gt 2022-05-10T12:00:00Z&$expand=SalesQuotationLines($select=LineCreationSequenceNumber,ItemNumber,LineDescription,SalesUnitSymbol,RequestedSalesQuantity,LineAmount)&$select=SalesQuotationNumber,ReceiptDateRequested,RequestingCustomerAccountNumber,SalesQuotationName,QuotationTakerPersonnelNumber,QuotationResponsiblePersonnelNumber";
      var result = entity.QueryEntity(url);
      SalesQuotations = JsonConvert.DeserializeObject<SalesQuotationResult>(result.Result.Content);
      return SalesQuotations;
    }

    public static EmployeesResult GetListOfEmployees()
    {
      ConsultaEntity entity = new ConsultaEntity();
      String url = "https://ayt.operations.dynamics.com/Data/Employees?$select=PersonnelNumber,Name";
      var result = entity.QueryEntity(url);
      EmployeesResult employees = JsonConvert.DeserializeObject<EmployeesResult>(result.Result.Content);
      return employees;
    }

    public static Boolean IsEquipoOrRefaccion(String ItemNumber)
    {
      ConsultaEntity entity = new ConsultaEntity();
      String url = "https://ayt.operations.dynamics.com/Data/ProductCategoryAssignments?$filter=ProductNumber eq '" + ItemNumber + "' and ProductCategoryHierarchyName eq 'CATEGORIA PROYECTO' and (ProductCategoryName eq 'EQUIPO' or ProductCategoryName eq 'REFACCION')";
      var result = entity.QueryEntity(url);
      ProductCategoryResult categoria = JsonConvert.DeserializeObject<ProductCategoryResult>(result.Result.Content);
      if (categoria.value.Count > 0)
      {
        return true;
      }
      return false;
    }

    public static String GetTemplateHTML(List<SalesQuotation> origen)
    {
      String html = "";
      foreach(SalesQuotation cotizacion in origen)
      {
        String SalesTaker = workers.value.Where(emp => emp.PersonnelNumber.Equals(cotizacion.QuotationTakerPersonnelNumber)).ToList()[0].Name;
        String SalesResponsible = workers.value.Where(emp => emp.PersonnelNumber.Equals(cotizacion.QuotationResponsiblePersonnelNumber)).ToList()[0].Name;
        html += "<table style=\"border-collapse:collapse; width: 90 %; font-size: 11px; margin: 20px !important; page-break-inside:avoid; border:solid 2px #CCC;\">";
        html += "  <thead>";
        html += "    <tr>";
        html += "      <td colspan=\"8\">&nbsp;</td>";
        html += "    </tr>";
        html += "    <tr style=\"font-weight:bolder;width: 90%;\">";
        html += "      <th>&nbsp;</th>";
        html += "      <th colspan=\"6\"><b>ID Cliente: </b>" + cotizacion.RequestingCustomerAccountNumber + "</th>";
        html += "      <th>&nbsp;</th>";
        html += "    </tr>";
        html += "    <tr style=\"font-weight:bolder;width: 90%;\">";
        html += "      <th>&nbsp;</th>";
        html += "      <th colspan=\"6\"><b>Nombre Cliente: </b>" + cotizacion.SalesQuotationName + "</th>";
        html += "      <th>&nbsp;</th>";
        html += "    </tr>";
        html += "    <tr style=\"font-weight:bolder;width: 90%;\">";
        html += "      <th colspan=\"8\">&nbsp;</th>";
        html += "    </tr>";
        html += "    <tr style=\"font-weight:bolder;width: 90%;\">";
        html += "      <th>&nbsp;</th>";
        html += "      <th colspan=\"6\"><b>Cotizacion: </b>" + cotizacion.SalesQuotationNumber + "</th>";
        html += "      <th>&nbsp;</th>";
        html += "    </tr>";
        html += "    <tr style=\"font-weight:bolder;width: 90%;\">";
        html += "      <th>&nbsp;</th>";
        html += "      <th colspan=\"6\"><b>Total: </b>{TotalCotizacion}</th>";
        html += "      <th>&nbsp;</th>";
        html += "    </tr>";
        html += "    <tr style=\"font-weight:bolder;width: 90%;\">";
        html += "      <th colspan=\"8\">&nbsp;</th>";
        html += "    </tr>";
        html += "    <tr style=\"font-weight:bolder;width: 90%;\">";
        html += "      <th>&nbsp;</th>";
        html += "      <th colspan=\"6\"><b>Sales Taker: </b>" + SalesTaker + "</th>";
        html += "      <th>&nbsp;</th>";
        html += "    </tr>";
        html += "    <tr style=\"font-weight:bolder;width: 90%;\">";
        html += "      <th>&nbsp;</th>";
        html += "      <th colspan=\"6\"><b>Sales Responsible: </b>" + SalesResponsible + "</th>";
        html += "      <th>&nbsp;</th>";
        html += "    </tr>";
        html += "    <tr>";
        html += "      <th colspan=\"8\">&nbsp;</th>";
        html += "    </tr>";
        html += "  </thead>";
        html += "  <tbody>";
        html += "    <tr style=\"width: 90%;\">";
        html += "      <td>&nbsp;</td>";
        html += "      <td style=\"border: solid 1px black; font-size: 11px; font-weight: bolder;\">Linea</td>";
        html += "      <td style=\"border: solid 1px black; font-size: 11px; font-weight: bolder;\">Item</td>";
        html += "      <td style=\"border: solid 1px black; font-size: 11px; font-weight: bolder;\">Descripcion</td>";
        html += "      <td style=\"border: solid 1px black; font-size: 11px; font-weight: bolder;\">Unidad</td>";
        html += "      <td style=\"border: solid 1px black; font-size: 11px; font-weight: bolder;\">Cantidad</td>";
        html += "      <td style=\"border: solid 1px black; font-size: 11px; font-weight: bolder;\">Total</td>";
        html += "      <td>&nbsp;</td>";
        html += "    </tr>";
        Double totalCoti = 0;
        foreach (SalesQuotationLines lines in cotizacion.SalesQuotationLines)
        {
          html += "    <tr style=\"width: 90%;\">";
          html += "      <td>&nbsp;</td>";
          html += "      <td style=\"border: solid 1px black; font-size: 10px;\">" + lines.LineCreationSequenceNumber + "</td>";
          html += "      <td style=\"border: solid 1px black; font-size: 10px;\">" + lines.ItemNumber + "</td>";
          html += "      <td style=\"border: solid 1px black; font-size: 10px;\">" + lines.LineDescription + "</td>";
          html += "      <td style=\"border: solid 1px black; font-size: 10px;\">" + lines.SalesUnitSymbol + "</td>";
          html += "      <td style=\"border: solid 1px black; font-size: 10px;\">" + lines.RequestedSalesQuantity.ToString()+ "</td>";
          html += "      <td style=\"border: solid 1px black; font-size: 10px;\">" + String.Format("{0:C}", lines.LineAmount) + "</td>";
          html += "      <td>&nbsp;</td>";
          html += "    </tr>";
          totalCoti += lines.LineAmount;
        }
        html += "    <tr>";
        html += "      <td colspan=\"8\">&nbsp;</td>";
        html += "    </tr>";
        html += "  </tbody>";
        html += "</table>";
        html = html.Replace("{TotalCotizacion}", String.Format("{0:C}", totalCoti));
        totalCoti = 0;
      }
      return html;
    }
  }

  public class SalesQuotationResult
  {
    public List<SalesQuotation> value { get; set; }
  }

  public class SalesQuotation
  {
    public String SalesQuotationNumber { get; set; }
    public String ReceiptDateRequested { get; set; }
    public String RequestingCustomerAccountNumber { get; set; }
    public String SalesQuotationName { get; set; }
    public String QuotationTakerPersonnelNumber { get; set; }
    public String QuotationResponsiblePersonnelNumber { get; set; }
    public String DefaultShippingSiteId { get; set; }
    public List<SalesQuotationLines> SalesQuotationLines { get; set; }
  }

  public class SalesQuotationLines
  {
    public Double LineCreationSequenceNumber { get; set; }
    public String ItemNumber { get; set; }
    public String LineDescription { get; set; }
    public String SalesUnitSymbol { get; set; }
    public Double RequestedSalesQuantity { get; set; }
    public Double LineAmount { get; set; }
  }

  public class ProductCategory
  {
    public String ProductCategoryName { get; set; }
    public String ProductCategoryHierarchyName { get; set; }
  }

  public class ProductCategoryResult
  {
    public List<ProductCategory> value { get; set; }
  }

  public class Employees
  {
    public String PersonnelNumber { get; set; }
    public String Name { get; set; }
  }

  public class EmployeesResult
  {
    public List<Employees> value { get; set; }
  }
}
