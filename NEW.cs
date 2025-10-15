using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using WebApplication1.Models;


namespace WebApplication3.Controllers

{
    public class StudyController : ApiController
    {



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

            // Kiểm tra các header phổ biến
            if (HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"] != null)
            {
                ipAddress = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"].ToString();
            }
            else if (!string.IsNullOrEmpty(HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"]))
            {
                ipAddress = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];


            }

            // Nếu có nhiều IP, lấy IP đầu tiên
            if (ipAddress.Contains(","))
            {
                ipAddress = ipAddress.Split(',')[0];
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
        //{
        //    "FilePath": "C:\\Users\\Admin\\Desktop\\test.txt"
        //}
        [HttpPost]
        [Route("webapi/study/DeleteFileOrFolder")]
        public IHttpActionResult DeleteFileOrFolder([FromBody] Study item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.FilePath))
                return BadRequest("Invalid request: file path is required.");

            string fullPath = item.FilePath; // Đường dẫn tuyệt đối

            try
            {
                // Nếu là file
                if (System.IO.File.Exists(fullPath))
                {
                    System.IO.File.Delete(fullPath);
                    return Ok(new { message = "Deleted file successfully" });
                }
                // Nếu là folder
                else if (System.IO.Directory.Exists(fullPath))
                {
                    // Xóa folder và toàn bộ nội dung bên trong
                    System.IO.Directory.Delete(fullPath, recursive: true);
                    return Ok(new { message = "Deleted folder successfully" });
                }
                else
                {
                    return Content(HttpStatusCode.NotFound, new { message = "File or folder not found" });
                }
            }
            catch (UnauthorizedAccessException)
            {
                return Content(HttpStatusCode.Forbidden, new { message = "No permission to delete" });
            }
            catch (IOException ex)
            {
                // Trường hợp folder bị khóa, đang được sử dụng
                return Content(HttpStatusCode.Conflict, new { message = "File or folder is in use", error = ex.Message });
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
                return BadRequest("Invalid request: file path is required.");

            string excelPath = request.FilePath;

            if (!System.IO.File.Exists(excelPath))
                return Content(HttpStatusCode.NotFound, new { message = "Excel file not found" });

            try
            {
                // Load Excel file
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(excelPath);

                // Lấy thư mục chứa file Excel
                string folder = Path.GetDirectoryName(excelPath);
                string filenameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);

                // Tạo đường dẫn cho file PDF
                string pdfPath = Path.Combine(folder, filenameWithoutExt + ".pdf");

                // Xuất ra PDF
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

                // ✅ Thư mục gốc
                string rootPath = @"C:\Users\Admin\Desktop\study\test";
                if (!Directory.Exists(rootPath))
                    Directory.CreateDirectory(rootPath);

                // Duyệt từng file gửi lên
                foreach (string fileKey in httpRequest.Files)
                {
                    var postedFile = httpRequest.Files[fileKey];
                    string originalFileName = Path.GetFileName(postedFile.FileName);

                    // ✅ Tạo folder mới theo ngày giờ
                    string folderName = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string newFolderPath = Path.Combine(rootPath, folderName);
                    Directory.CreateDirectory(newFolderPath);

                    // ✅ Đặt tên file mới
                    string newFileName = $"{DateTime.Now:yyyyMMdd}_{originalFileName}";
                    string newFilePath = Path.Combine(newFolderPath, newFileName);

                    // ✅ Lưu file
                    postedFile.SaveAs(newFilePath);

                    return Ok(new
                    {
                        message = "✅ File uploaded successfully",
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





        //===========================================================================================================================
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
                    client.EnableSsl = false;

                    client.Send(mailMsg);
                }

                return Ok(new { success = true, message = "Mail sent successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest("Error sending mail: " + ex.Message);
            }
        }

         private static void SendEmail(string partno, string ckPoint, string ckPosition)
        {
            try
            {
                string senderEmail = "Asset-sys@local.canon-vn.com.vn";  // Địa chỉ gửi
                string receiverEmail = "anh.nguyen707@mail.canon";       // Địa chỉ nhận
                string smtpServer = "nonauth-smtp.global.canon.co.jp";        // Máy chủ SMTP
                int smtpPort = 25;
                string smtpUser = "Asset-sys@local.canon-vn.com.vn";  // Tài khoản gửi
                string smtpPassword = "YourPasswordHere";  // Mật khẩu (bảo mật)

                

                string subject = "Thông báo lỗi";
                string body = $"Điểm đo đầu tiên không tìm thấy của partno! \n {partno} \nĐiểm: {ckPoint}\n - {ckPosition}";

                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(senderEmail);
                mail.To.Add(receiverEmail);
                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = false;

                SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort);
                smtpClient.Credentials = new NetworkCredential(smtpUser, smtpPassword);
                smtpClient.EnableSsl = true; // Kích hoạt SSL nếu cần

                smtpClient.Send(mail);
                Console.WriteLine("Email đã gửi thành công!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi gửi email: {ex.Message}");
            }
        }
    }

}
