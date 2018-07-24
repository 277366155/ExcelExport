using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web.Mvc;
using ExcelExport.Common;
using ExcelExport.EF;
using NPOI.HSSF.UserModel;

namespace ExcelExport.Controllers
{
    public class TestController : Controller
    {
        // GET: Test
        public ActionResult Index()
        {
            return View();
        }

      
        public ActionResult Export(int num)
        {
            using (var entity = new BaseContext())
            {
                 entity.Users.Export("测试文件",num);//,10,m=>m.OrderBy(a=>a.Age).ThenByDescending(a=>a.Name));              
            }

            return null;
        }

        #region  生成指定数量的数据 
        [HttpPost]
        public ActionResult Init(int num)
        {
            return Json(InitData(num));
        }
        private int InitData(int num)
        {
            // int result = 0;
            var rd = new Random(GetRandomSeed());
            var str = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
            using (var entity = new BaseContext())
            {
                entity.Database.ExecuteSqlCommand("delete from User");
                for (int i = 0; i < num; i++)
                {
                    var user = new User();
                    user.Name = GetRandomStr(6, str);
                    user.Age = rd.Next(12, 80);
                    user.Sex = rd.Next(0, 2);

                    entity.Users.Add(user);
                }
                return entity.SaveChanges();
            }
        }
        /// <summary>
        /// 获取随机字符串
        /// </summary>
        /// <param name="lenth">结果字符串长度</param>
        /// <param name="str">种子字符</param>
        /// <returns></returns>
        private string GetRandomStr(int lenth, string str)
        {
            if (string.IsNullOrWhiteSpace(str) || lenth <= 0)
            {
                return null;
            }

            var result = new StringBuilder();

            var rd = new Random(GetRandomSeed());

            for (var i = 0; i < lenth; i++)
            {
                result.Append(str.ToCharArray()[rd.Next(0, str.Length)]);
            }

            return result.ToString();
        }
        private int GetRandomSeed()
        {
            byte[] bytes = new byte[4];
            System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();
            rng.GetBytes(bytes);
            return BitConverter.ToInt32(bytes, 0);
        }
        #endregion
    }
}