using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net.Mail;

namespace RenameFiles
{
    class Program
    {

        static void Main(string[] args)
        {
            string Source = @"r:\micros\opera\export\operabg";
            string targetPath = @"r:\micros\opera\export\operabg\tmp_files";
            string subdirectory;
            bool GRSE = false;
            bool RMO = false;
            bool RVB = false;
            string EmailGRSE = "";
            string EmailRMO = "";
            string EmailRVB = "";
            DateTime data = new DateTime();
            data = DateTime.Today.AddDays(-1);
            string data2 = data.ToString("ddMMyy");


            if (Directory.Exists(targetPath))
            {
                System.IO.Directory.Delete(targetPath, true);
                
            }
            System.IO.Directory.CreateDirectory(targetPath);
            
            try
            {
                foreach (string newPath in Directory.GetFiles(Source, "*.aud", SearchOption.AllDirectories))
                {
                    subdirectory = Path.GetDirectoryName(newPath);
                    string Filename;
                
                string dateFile;
                string HotelName;

                    Filename = Path.GetFileName(newPath);
                    dateFile = Filename.Substring(0, 6);
                    if (data2 == dateFile)
                    {
                        HotelName = Filename.Substring(6);
                        switch (HotelName)
                        {
                            case "grse.aud":
                                GRSE = true;
                                break;
                            case "rmo.aud":
                                RMO = true;
                                break;
                            case "hrb.aud":
                                RVB = true;
                                break;
                        }
                        if (GRSE == true)
                        {
                            
                            Console.WriteLine("encontrou GRSE");
                            File.Copy(newPath, newPath.Replace(subdirectory, targetPath), true);
                            EmailGRSE = targetPath + "\\" + Path.GetFileNameWithoutExtension(Filename) + ".zip";
                        }
                        if (RMO == true)
                        {
                            
                            Console.WriteLine("encontrou RMO");
                            File.Copy(newPath, newPath.Replace(subdirectory, targetPath), true);
                            EmailRMO = targetPath + "\\" + Path.GetFileNameWithoutExtension(Filename) + ".zip";
                        }
                        if (RVB == true)
                        {
                            
                            Console.WriteLine("encontrou RVB");
                            File.Copy(newPath, newPath.Replace(subdirectory, targetPath), true);
                            EmailRVB = targetPath + "\\" + Path.GetFileNameWithoutExtension(Filename) + ".zip";
                        }
                        GRSE = false;
                        RMO = false;
                        RVB = false;
                        
                    }
                }
                    
                }
                        catch (Exception e)
            {
                Console.WriteLine("{0} Erro na pesquisa do ficheiro", e);
            }

            try
            {
                
                foreach (string newPath in Directory.GetFiles(targetPath, "*.*"))
                {
                    File.Move(newPath, Path.ChangeExtension(newPath, ".zip"));
                }
                Console.WriteLine("mudança de extensão com sucesso");
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Erro na mudança de extensão", e);
            }

            // Emails censured for privacy reasons
            if(EmailGRSE != String.Empty)
            {
                Console.WriteLine("a enviar GRSE email...");
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            message.To.Add("x@x.com");
            message.CC.Add("x@x.com");
            Attachment anexo = new Attachment(EmailGRSE);
            message.Attachments.Add(anexo);
            message.Subject = "Relatórios de Audit do Grande Real Santa Eulália";
            message.From = new System.Net.Mail.MailAddress("x@x.com");
            message.Body = "Bom dia a todos, segue em anexo os relatórios de Audit " + data.ToString("dd-MM-yyyy") + " referentes ao Grande Real Santa Eulália";

            System.Net.Mail.SmtpClient smtp = new SmtpClient();
            smtp.Host = "x.x.com";
            smtp.EnableSsl = false;
            smtp.Credentials = new System.Net.NetworkCredential("x@x.com", "");
            
            smtp.Send(message);
            Console.WriteLine("Enviado para o destinatario");

            }
            if (EmailRMO != String.Empty)
            {
                Console.WriteLine("a enviar RMO email...");
                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
                message.To.Add("x@x.com");
                message.CC.Add("x@x.com");
                Attachment anexo = new Attachment(EmailRMO);
                message.Attachments.Add(anexo);
                message.Subject = "Relatórios de Audit do Hotel Real Marina";
                message.From = new System.Net.Mail.MailAddress("x@x.com");
                message.Body = "Bom dia a todos, segue em anexo os relatórios de Audit " + data.ToString("dd-MM-yyyy") + " referentes ao Hotel Real Marina";

                System.Net.Mail.SmtpClient smtp = new SmtpClient();
                smtp.Host = "mail.hoteisreal.com";
                smtp.EnableSsl = false;
                smtp.Credentials = new System.Net.NetworkCredential("x@x.com", "");

                smtp.Send(message);
                Console.WriteLine("Enviado para o destinatario");
            }
            if (EmailRVB != String.Empty)
            {
                Console.WriteLine("a enviar RBV email...");
                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
                message.To.Add("x@x.com");
                message.CC.Add("x@x.com");
                Attachment anexo = new Attachment(EmailRVB);
                message.Attachments.Add(anexo);
                message.Subject = "Relatórios de Audit do Hotel Real Bellavista";
                message.From = new System.Net.Mail.MailAddress("x@x.com");
                message.Body = "Bom dia a todos, segue em anexo os relatórios de Audit "+data.ToString("dd-MM-yyyy")+" referentes ao Hotel Real Bellavista";

                System.Net.Mail.SmtpClient smtp = new SmtpClient();
                smtp.Host = "x.x.com";
                smtp.EnableSsl = false;
                smtp.Credentials = new System.Net.NetworkCredential("x@x.com", "");

                smtp.Send(message);
                Console.WriteLine("Enviado para o destinatario");
            }

        }
    }
}
