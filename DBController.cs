using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Results;
using System.Data.SqlClient;
using System.Data;
using WebApplication1.App_Start;
using System.Data.OleDb;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
//using System.Data.SqlClient;

namespace WebApplication1.Controllers
{
    public class DBController : ApiController
    {
        

        public string Hello(string name)
        {
            return "Hello" + name;
        }

        public class Body
        {
            public string query { get; set; }
        }
        [Route("webapi/study/vaselect")]
        [HttpPost]
        public Object Select([FromBody] Body body)
        {
            try
            {
               // string sql = "SELECT * FROM [UserInformation_Test].[dbo].[ENG_memberList]";
                var dt = Common.ExcuteDataTable_User(body.query);
                return dt;
            }
            catch (Exception ex)
            {
               
                var response = new
                {
                    body = body,
                    errMessage = ex
                };
                return response;
            }
        }

        //UPDATE, INSERT, DELETE QUERY
        [Route("webapi/study/vaexecute")]
        [HttpPost]
        public Object Execute([FromBody] Body body)
        {

            try
            {
                Common.Excute_SQL( body.query);
                var response = new
                {
                    errMessage = "Success!"
                };
                return response;
            }
            catch (Exception ex)
            {
                var response = new
                {
                    body = body,
                    errMessage = ex
                };
                return response;
            }
        }

        [HttpGet]
        [Route("webapi/study/SelectSQL")]
        public DataTable SelectSQL()
        {
            DataTable dt = new DataTable();
            string sql = "SELECT * FROM [UserInformation_Test].[dbo].[ENG_memberList]";

            try
            {
                using (OleDbConnection conn = Common.GetConnectionUser()) 
                {
                    conn.Open();
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
                    {
                        da.Fill(dt);
                    }
                }
            }
            catch (System.Exception)
            {
               
                throw;
            }

            return dt;
        }

        public string Database = System.Configuration.ConfigurationManager.AppSettings["Database"];
        [HttpGet]
        [Route("webapi/study/SelectSQL2")]
        public DataTable SQLtest2(int top)
        {

           
            DataTable td = new DataTable();
            string sql = $"SELECT TOP {top} * FROM {Database}";
            try {
                using (OleDbConnection conn = Common.GetConnectionUser()) 
                {
                    conn.Open(); 
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
                    {
                        da.Fill(td);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return td;
        }


        [HttpGet]
        [Route("webapi/study/PostSelectSQL")]
        public IHttpActionResult SQLtest([FromUri]  string name, [FromUri]  string code)
        {
            
            
            try {
                string sql = $"Insert into {Database} (Province,Code) values(?, ?)";
                using (OleDbConnection conn = Common.GetConnectionUser()) 
                {
                    conn.Open(); 
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("?", name);
                        cmd.Parameters.AddWithValue("?", code);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                            return Ok("ok");
                        else
                            return BadRequest("NG");
                    }
                    
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        [HttpGet]
        [Route("webapi/study/UpdateSelectSQL")]
        public IHttpActionResult Updatetest([FromUri]  string name, [FromUri]  string code, [FromUri]  int id)
      // public IHttpActionResult Updatetest([FromUri] string sql) //insert /update/delete/....
        {


            try
            {
                string sql = $"Update {Database} Set Province ='"+ name + "', code='"+ code + "' where id ='"+ id +"' ";
                using (OleDbConnection conn = Common.GetConnectionUser()) 
                {
                    conn.Open(); 
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("name", name);
                        cmd.Parameters.AddWithValue("code", code);
                        cmd.Parameters.AddWithValue("id", id);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                            return Ok("ok");
                        else
                            return BadRequest("NG");
                    }

                }
            }
            catch (Exception)
            {
                throw;
            }
        }


        [HttpGet]
        [Route("webapi/study/DeleteSelectSQL")]
        public IHttpActionResult Deletetest([FromUri]  int id)
        {


            try
            {
                string sql = $"delete from {Database}  where id =? ";
                using (OleDbConnection conn = Common.GetConnectionUser()) 
                {
                    conn.Open();
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                     
                        cmd.Parameters.AddWithValue("?", id);
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                            return Ok("ok");
                        else
                            return BadRequest("NG");
                    }

                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public class test
        {
            public string name { get; set; }
            public string code { get; set; }
            public int id { get; set; }
        }
        [HttpGet]
        [Route("webapi/study/test")]
        public IHttpActionResult Test(test item)
        {
            try {
                string sql = "Update" + Database + "Set Province = (N'" + item.name + "'), code = (N'" + item.code + "') where id = (N'" + item.id + "')";

                using (OleDbConnection conn = Common.GetConnectionUser())
                {
                    conn.Open();
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.CommandText = sql;
                        cmd.Prepare();
                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                            return Ok("ok");
                        else
                            return BadRequest("NG");
                    }

                }

            }
            catch (Exception)
            {
                throw;
            }
        }



    }
}
