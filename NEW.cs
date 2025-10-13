using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Web;
using System.Web.Http;
using System.Data;
using WebApplication3.Models;
using System.Net.NetworkInformation;
using System.IO;

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
                IP_MAC = "MAC: "+ clientMac
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
            catch (Exception ex){
                clientMac = "Error: " + ex.Message;
            }
            return clientMac;
        }


        //=======using HttpPost để thực hiện function xóa tất cả các file trong folder + xóa folder đó //

        [HttpPost]
        [Route("webapi/study/DeleteFile")]
        public IHttpActionResult DeleteFile([FromBody] Study file)
        {
            if (file == null || string.IsNullOrWhiteSpace(file.FilePath))
                return BadRequest("Invalid request: file path is required.");

            string fullPath = file.FilePath; // chấp nhận đường dẫn tuyệt đối

            try
            {
                if (System.IO.File.Exists(fullPath))
                {
                    System.IO.File.Delete(fullPath);
                    return Ok(new { message = "Deleted file" });
                }
                else
                {
                    return Content(HttpStatusCode.NotFound, new { message = "File not found" });
                }
            }
            catch (UnauthorizedAccessException)
            {
                return Content(HttpStatusCode.Forbidden, new { message = "No permission to delete file" });
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }

    }
}
