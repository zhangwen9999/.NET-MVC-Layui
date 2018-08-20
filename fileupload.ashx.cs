using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Net;
using System.Web;
using System.Collections;
using System.Globalization;
using System.Web.Script.Serialization;

namespace QiShi_BAS_MVC.Api
{
    /// <summary>
    /// fileupload 的摘要说明
    /// </summary>
    public class fileupload : IHttpHandler
    {
        JavaScriptSerializer ser = new JavaScriptSerializer();
        public void ProcessRequest(HttpContext context)
        {
            try
            {
                //文件保存目录路径
                String savePath = "../upload/excel/";
                //定义允许上传的文件扩展名
                Hashtable extTable = new Hashtable();
                extTable.Add("excel", "xls,xlsx");
                //最大文件大小
                int maxSize = 10000000;
                HttpPostedFile imgFile = context.Request.Files[0];
                if (imgFile == null)
                {
                    showError(context, "请选择需要上传的文件。");
                    return;
                }
                if (imgFile.InputStream == null || imgFile.InputStream.Length > maxSize)
                {
                    showError(context, "上传文件大小超过限制。");
                    return;
                }
                String dirPath = context.Server.MapPath(savePath);
                String fileName = imgFile.FileName;
                String fileExt = Path.GetExtension(fileName).ToLower();
                if (String.IsNullOrEmpty(fileExt) || Array.IndexOf(((String)extTable["excel"]).Split(','), fileExt.Substring(1).ToLower()) == -1)
                { 
                      showError(context,"上传文件扩展名是不允许的扩展名。\n只允许" + ((String)extTable["excel"]) + "格式。");
                    return;
                }
                //创建文件夹
                if (!Directory.Exists(dirPath))
                {
                    Directory.CreateDirectory(dirPath);
                }
                DateTime now = DateTime.Now;
                String ymd = now.ToString("yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                dirPath += ymd + "/";
                savePath += ymd + "/";
                if (!Directory.Exists(dirPath))
                {
                    Directory.CreateDirectory(dirPath);
                }

                String newFileName = now.ToString("HHmmss_ffff", DateTimeFormatInfo.InvariantInfo);
                String filePath = dirPath + newFileName + fileExt;
                string newpath = savePath + newFileName + fileExt;
                imgFile.SaveAs(filePath);
                Hashtable hash = new Hashtable(); 
                hash["Code"] = 1;
                hash["Data"] = newpath;
                hash["Message"] = "服务器本地上传成功"; 
                context.Response.Write(ser.Serialize(hash)); 
            }
            catch (Exception ex)
            { showError(context, ex.Message); }
        }
        private void showError(HttpContext context, string message)
        {
            Hashtable hash = new Hashtable();
            hash["Code"] = -1;
            hash["Message"] = message; 
            context.Response.Write(ser.Serialize(hash)); 
        }
        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}