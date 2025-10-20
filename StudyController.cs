using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Web;
using System.Web.Http;
using System.Data;
using WebApplication1.Models;
using System.Net.NetworkInformation;
using System.IO;
using System.Net;
using Spire.Xls;
using System.Net.Mail;



namespace WebApplication1.Controllers

{
    public class StudyController : ApiController
    {


        [Route("webapi/study/getip")]
        [System.Web.Http.HttpGet]
        public string GetIp()
        {

            string ipAddress = string.Empty;


            if (HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"] != null)
            {
                ipAddress = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"].ToString();
            }
            else if (!string.IsNullOrEmpty(HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"]))
            {
                ipAddress = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
            }


            return ipAddress;
        }

        [Route("webapi/study/getmac")]
        [System.Web.Http.HttpGet]
        public string GetMac()
        {

            string clientMac = "";
            try
            {
                NetworkInterface[] interfaces = NetworkInterface.GetAllNetworkInterfaces();
                foreach (NetworkInterface nic in interfaces)
                {
                    if (nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                    {
                        PhysicalAddress physicalAddr = nic.GetPhysicalAddress();
                        byte[] bytes = physicalAddr.GetAddressBytes();
                        for (int i = 0; i < bytes.Length; i++)
                        {
                            clientMac += bytes[i].ToString("X2");
                            if (i != bytes.Length - 1)
                            {
                                clientMac += ":";
                            }
                        }
                        if (!string.IsNullOrEmpty(clientMac))
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                clientMac = "Error: " + ex.Message;
            }
            return clientMac;
        }

        [HttpGet]
        [Route("webapi/study/GET_IP")]
        public HttpResponseMessage GetClientIp()
        {

            var clientIp = GetClientIPAddress();
            var clientMac = GetClientMACAddress();


            List<Study> list_Study = new List<Study>();
            Study ip = new Study
            {
                IP = "IP: " + clientIp,
                IP_MAC = "MAC: " + clientMac
            };
            list_Study.Add(ip);

            return list_Study != null
                ? Request.CreateResponse(HttpStatusCode.OK, list_Study)
                : Request.CreateResponse(HttpStatusCode.NotFound);
        }



        public string GetClientIPAddress()
        {
            string ipAddress = string.Empty;

          
            if (HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"] != null)
            {
                ipAddress = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"].ToString();
            }
            else if (!string.IsNullOrEmpty(HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"]))
            {
                ipAddress = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
            }


            return ipAddress;
        }

        

        public string GetClientMACAddress()
        {
            string clientMac = "";
            try
            {
                NetworkInterface[] interfaces = NetworkInterface.GetAllNetworkInterfaces();
                foreach (NetworkInterface nic in interfaces)
                {
                    if (nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                    {
                        PhysicalAddress physicalAddr = nic.GetPhysicalAddress();
                        byte[] bytes = physicalAddr.GetAddressBytes();
                        for (int i = 0; i < bytes.Length; i++)
                        {
                            clientMac += bytes[i].ToString("X2");
                            if (i != bytes.Length - 1)
                            {
                                clientMac += ":";
                            }
                        }
                        if (!string.IsNullOrEmpty(clientMac))
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                clientMac = "Error: " + ex.Message;
            }
            return clientMac;
        }


        //=======using HttpPost để thực hiện function xóa tất cả các file trong folder + xóa folder đó //

        [HttpPost]
        [Route("webapi/study/DeleteFileOrFolder")]
        public IHttpActionResult DeleteFileOrFolder([FromBody] Study item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.FilePath))
                return BadRequest("Invalid request: file path is required.");

            string fullPath = item.FilePath; 

            try
            {
       
                if (System.IO.File.Exists(fullPath))
                {
                    System.IO.File.Delete(fullPath);
                    return Ok(new { message = "Deleted file" });
                }
               
                else if (System.IO.Directory.Exists(fullPath))
                {
                    
                    System.IO.Directory.Delete(fullPath, recursive: true);
                    return Ok(new { message = "Deleted folder " });
                }
                else
                {
                    return Content(HttpStatusCode.NotFound, new { message = " not found" });
                }
            }
            
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }

        //- using HttpPost để thực hiện function sau khi update file sẽ tạo folder mới và save file vào folder vừa tạo với file name = date + tên file ban đầu.
        [HttpPost]
        [Route("webapi/study/UploadFile")]
        public IHttpActionResult UploadFile()
        {
            try
            {
                var httpRequest = HttpContext.Current.Request;

                if (httpRequest.Files.Count == 0)
                    return BadRequest("No file uploaded.");

  
                string rootPath = @"\\cvn-veng\DCResource\Study_test\test";
                if (!Directory.Exists(rootPath))
                    Directory.CreateDirectory(rootPath);

                
                foreach (string fileKey in httpRequest.Files)
                {
                    var postedFile = httpRequest.Files[fileKey];
                    string originalFileName = Path.GetFileName(postedFile.FileName);

                   
                    string folderName = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string newFolderPath = Path.Combine(rootPath, folderName);
                    Directory.CreateDirectory(newFolderPath);

                 
                    string newFileName = $"{DateTime.Now:yyyyMMdd}_{originalFileName}";
                    string newFilePath = Path.Combine(newFolderPath, newFileName);

               
                    postedFile.SaveAs(newFilePath);

                    return Ok(new
                    {
                        message = " File uploaded successfully",
                        folder = newFolderPath,
                        fileName = newFileName,
                        filePath = newFilePath
                    });
                }

                return BadRequest("No valid file found in request.");
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }

        //====using HttpPost để thực hiện function export 1 file excel có sẵn thành file pdf và lưu vào đúng đường dẫn chứa file excel đó.
        [HttpPost]
        [Route("webapi/study/ExportExcelToPdf")]
        public IHttpActionResult ExportExcelToPdf([FromBody] Study request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.FilePath))
                return BadRequest(" request: not file.");

            string excelPath = request.FilePath;

            if (!System.IO.File.Exists(excelPath))
                return Content(HttpStatusCode.NotFound, new { message = "not found" });

            try
            {
           
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(excelPath);

                
                string folder = Path.GetDirectoryName(excelPath);
        

 
                string pdfPath = Path.Combine(folder + ".pdf");

                
                workbook.SaveToFile(pdfPath, Spire.Xls.FileFormat.PDF);

                return Ok(new
                {
                    message = "Export successful",
                    pdfPath = pdfPath
                });
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }

        //===- using HttpPost để thực hiện function gửi mail có các tham số truyền đầu vào (from + subject + To + CC + body) hiện tại mình dùng SMTP có host và port như này nhé 
        [HttpPost]
        [Route("webapi/study/sendmail")]
        public IHttpActionResult SendMail([FromBody] Study emailBody)
        {
            try
            {
                if (emailBody == null)
                    return BadRequest("Email body is missing.");

                var mailMsg = new System.Net.Mail.MailMessage();
                mailMsg.From = new System.Net.Mail.MailAddress(emailBody.From);
                mailMsg.To.Add(emailBody.To);

                if (!string.IsNullOrEmpty(emailBody.Cc))
                    mailMsg.CC.Add(emailBody.Cc);

                mailMsg.Subject = emailBody.Subject;
                mailMsg.Body = emailBody.Body;
                mailMsg.IsBodyHtml = true;
                mailMsg.Priority = System.Net.Mail.MailPriority.Normal;

                using (var client = new System.Net.Mail.SmtpClient("nonauth-smtp.global.canon.co.jp", 25))
                {
                    client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    client.Credentials = new System.Net.NetworkCredential(emailBody.From, "cvn-sys");
                    client.EnableSsl =  false;

                    client.Send(mailMsg);
                }

                return Ok(new { success = true, message = "Mail sent successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest("Error sending mail: " + ex.Message);
            }
        }

    }
}