using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace CotizacionesPendientesEquipos
{
  class Mail
  {
    private SmtpClient client = new SmtpClient();
    private SmtpClient clientLIN = new SmtpClient();
    private SmtpClient clientINN = new SmtpClient();
    private NetworkCredential Credential = new NetworkCredential("notificacioninterna@avanceytec.com.mx", "8hn?X#~J?rvOI.+#=:E1S1}2C3xi6h)K");
    private NetworkCredential CredentialLIN = new NetworkCredential("facturacion.lideart@lideart.com.mx", "jhU_?qM.E?pNK%L8");
    private NetworkCredential CredentialINN = new NetworkCredential("facturacion.innovagames@innovagamesmexico.com.mx", "jhU_?qM.E?pNK%L8");
    private MailMessage mail = new MailMessage();

    public Mail()
    {
      /////////////////////////cliente atp de correo///////////
      client.Port = 587;
      client.DeliveryMethod = SmtpDeliveryMethod.Network;
      client.DeliveryFormat = SmtpDeliveryFormat.International;
      client.UseDefaultCredentials = false;
      client.Host = "smtp.office365.com";
      client.EnableSsl = true;
      client.Credentials = Credential;
      //////////////////////cliente lideart de correo///////////////////////////
      clientLIN.Port = 587;
      clientLIN.DeliveryMethod = SmtpDeliveryMethod.Network;
      clientLIN.DeliveryFormat = SmtpDeliveryFormat.International;
      clientLIN.UseDefaultCredentials = false;
      clientLIN.Host = "smtp.office365.com";
      clientLIN.EnableSsl = true;
      clientLIN.Credentials = CredentialLIN;
      //////////////////////cliente innova de correo///////////////////////////
      clientINN.Port = 587;
      clientINN.DeliveryMethod = SmtpDeliveryMethod.Network;
      clientINN.DeliveryFormat = SmtpDeliveryFormat.International;
      clientINN.UseDefaultCredentials = false;
      clientINN.Host = "smtp.office365.com";
      clientINN.EnableSsl = true;
      clientINN.Credentials = CredentialINN;
    }

    public Boolean SendMail(String template, MemoryStream ms2)
    {
      String body = String.Empty;
      String company = String.Empty;
      String cuentaAlias = String.Empty;
      String cuentaFactura = String.Empty;

      try
      {
        cuentaAlias = "Notificaciones Avance";
        mail.Subject = "Cotizaciones pendientes Equipos";
        mail.Body = "<p>Esta usted recibiendo un archivo adjunto con la lista de contizaciones pendientes con articulos que son de la categoria refaccion o equipo.<p>";
        mail.IsBodyHtml = true;
        //String emailCliente = "edacosta@avanceytec.com.mx";
        //String emailCliente = "sistemas04@avanceytec.com.mx";
        var AppSettings = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(@"appsettings.json"));
        var listaCorreos = AppSettings["ListaCorreo"];
        foreach (String emailCliente in listaCorreos)
        {
          if (IsValidEmail(emailCliente))
          {
            mail.To.Add(new MailAddress(emailCliente, emailCliente, Encoding.UTF8));
            //mail.To.Add(new MailAddress("sistemas04@avanceytec.com.mx", "sistemas04@avanceytec.com.mx", Encoding.UTF8));
          }
        }
        mail.Bcc.Add("larmendariz@avanceytec.com.mx");
        mail.Bcc.Add("edacosta@avanceytec.com.mx");

        System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);
        System.Net.Mail.Attachment attach = new System.Net.Mail.Attachment(new MemoryStream(ms2.ToArray()), ct);
        attach.ContentDisposition.FileName = "Cotizaciones Pendientes Equipos (" + DateTime.Now.ToString("yyyy-MM-dd") + ").pdf";
        
        try
        {
          mail.From = new MailAddress("notificacioninterna@avanceytec.com.mx", cuentaAlias);
          mail.Attachments.Add(attach);
          client.Send(mail);
        }
        catch (Exception mailEx)
        {
          Console.WriteLine(mailEx.Message);
          return false;
        }
        mail.Attachments.Clear();
        return true;
      }
      catch (Exception e)
      {
        Console.WriteLine(e.Message);
        return false;
      }
    }
    public Boolean IsValidEmail(string email)
    {
      if (email.Trim().EndsWith("."))
      {
        return false; // suggested by @TK-421
      }
      if (email.Contains(".@"))
      {
        return false; // suggested by jack(@Luis Armendariz)
      }
      try
      {
        var addr = new System.Net.Mail.MailAddress(email,email,Encoding.UTF8);
        Boolean isValid = addr.Host.Contains(".");
        if (isValid)
        {
          return addr.Address == email;
        }
        else
        {
          return false;
        }
      }
      catch
      {
        return false;
      }
    }
  }
}
