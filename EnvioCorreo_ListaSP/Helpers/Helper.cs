using EnvioCorreo_ListaSP.DTO;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace EnvioCorreo_ListaSP.Helpers
{
    public class Helper
    {
        public static void ConstruirCorreoError(Correo correo, string cuerpo)
        {
            try
            {
                correo.Cuerpo = $"<table border='0' align='left' cellpadding='0' cellspacing='0' style='width: 100%'>" +
             "<tr>" +
                 "<td align='left' valign='top'> Estimados Sres Empresa,</td>" +
             "</tr>" +
             "<tr>" +
                 $"<td align='left' valign='top'> {cuerpo}</td>" +
             "</tr>" +

                  "<tr style='padding:10px;' align='center'>" +
                      "<td align='left' valign='top'>" +
                          "<span> Saludos, </span> <br>" +
                          "<b> " + "Area de TI - PSAM <br>" +
                          "<b> " + " <br> </b></td></tr><br></table>";

                var correoEnvio = ConfigurationManager.AppSettings["CorreoErrorEnvio"].ToString().Split('|');
                var correosPara = (ConfigurationManager.AppSettings["CorreoErrorPara"].ToString().IndexOf(',') > -1) ? ConfigurationManager.AppSettings["CorreoErrorPara"].ToString().Split(',') : new string[] { ConfigurationManager.AppSettings["CorreoErrorPara"].ToString() };
                var correosCC = (ConfigurationManager.AppSettings["CorreoErrorCC"].ToString().IndexOf(',') > -1) ? ConfigurationManager.AppSettings["CorreoErrorCC"].ToString().Split(',') : new string[] { ConfigurationManager.AppSettings["CorreoErrorCC"].ToString() };
                var correosCO = (ConfigurationManager.AppSettings["CorreoErrorCO"].ToString().IndexOf(',') > -1) ? ConfigurationManager.AppSettings["CorreoErrorCO"].ToString().Split(',') : new string[] { ConfigurationManager.AppSettings["CorreoErrorCO"].ToString() };
                List<StructMail> correosErrorPara = new List<StructMail>();
                List<StructMail> correosErrorCC = new List<StructMail>();
                List<StructMail> correosErrorCO = new List<StructMail>();

                correo.CorreoEmisor.Mail = correoEnvio[0];
                correo.CorreoEmisor.Password = correoEnvio[1];
                correo.CorreoEmisor.NameMail = correoEnvio[2];
                correosPara = (string.IsNullOrEmpty(correosPara[0])) ? null : correosPara;
                correosCC = (string.IsNullOrEmpty(correosCC[0])) ? null : correosCC;
                correosCO = (string.IsNullOrEmpty(correosCO[0])) ? null : correosCO;
                if (correo.CorreosPara.Count == 0)
                {
                    if (correosPara != null)
                    {
                        foreach (var correoPara in correosPara)
                        {
                            correo.CorreosPara.Add(new StructMail() { Mail = correoPara, NameMail = "", Password = string.Empty });
                        }
                    }
                }

                if (correosCC != null)
                {
                    foreach (var correoCC in correosCC)
                    {

                        correo.CorreosCC.Add(new StructMail() { Mail = correoCC, NameMail = "", Password = string.Empty });
                    }
                }
                if (correosCO != null)
                {
                    foreach (var correoCO in correosCO)
                    {

                        correo.CorreosCCO.Add(new StructMail() { Mail = correoCO, NameMail = "", Password = string.Empty });
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }
        public static bool EnviarCorreoElectronico(Correo correo, bool esHTML, bool AcuseRecibo = true)
        {
            var mailMsg = new MailMessage();
            mailMsg.From = new MailAddress(correo.CorreoEmisor.Mail, correo.CorreoEmisor.NameMail);
            foreach (StructMail correopara in correo.CorreosPara)
                mailMsg.To.Add(new MailAddress(correopara.Mail, correopara.NameMail));
            foreach (StructMail correocc in correo.CorreosCC)
                mailMsg.CC.Add(new MailAddress(correocc.Mail, correocc.NameMail));
            foreach (StructMail correocco in correo.CorreosCCO)
                mailMsg.Bcc.Add(new MailAddress(correocco.Mail, correocco.NameMail));
            mailMsg.Subject = correo.Asunto;
            mailMsg.Body = correo.Cuerpo;
            if (esHTML)
            {
                AlternateView htmlView;
                htmlView = AlternateView.CreateAlternateViewFromString(correo.Cuerpo, Encoding.UTF8, "text/html");
                mailMsg.AlternateViews.Add(htmlView);
            }

            if (correo.Adjuntos.Count > 0)
            {
                foreach (var adjunto in correo.Adjuntos)
                {
                    Attachment att = new Attachment(new MemoryStream(adjunto.archivo), adjunto.nombreArchivo);
                    mailMsg.Attachments.Add(att);
                }
            }

            var SmtpClient = new SmtpClient();
            try
            {
                SmtpClient = new SmtpClient("smtp.office365.com", 587);//Cambiar por proveedor de appsettings.json
                //SmtpClient = new SmtpClient("smtp.gmail.com", 587);//Cambiar por proveedor de appsettings.json
                SmtpClient.Credentials = new NetworkCredential(correo.CorreoEmisor.Mail, correo.CorreoEmisor.Password);
                SmtpClient.EnableSsl = true; //correoEmisor.SMTPSSL;
            }
            catch (Exception ex)
            {

            }

            if (AcuseRecibo)
            {
                //SOLICITAR ACUSE DE RECIBO Y LECTURA
                mailMsg.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure | DeliveryNotificationOptions.OnSuccess | DeliveryNotificationOptions.Delay;
                mailMsg.Headers.Add("Disposition-Notification-To", correo.CorreoEmisor.Mail); //solicitar acuse de recibo al abrir mensaje
            }
            try
            {
                SmtpClient.Send(mailMsg);
            }
            catch (Exception ex)
            {
                try
                {
                    //reenviando en caso de error
                    mailMsg.DeliveryNotificationOptions = DeliveryNotificationOptions.None;
                    mailMsg.Headers.Remove("Disposition-Notification-To");
                    SmtpClient.Send(mailMsg);
                }
                catch (Exception exc)
                {

                }
            }
            return true;
        }
    }
}
