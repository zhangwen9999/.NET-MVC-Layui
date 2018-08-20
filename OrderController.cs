using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json; 
using QiShi_BAS_DAL;
using QiShi_BAS_VO;

namespace QiShi_BAS_MVC.Controllers
{
    public class OrderController : BaseController
    {
        Order_ComplaintDAL cdata = new Order_ComplaintDAL(); 
        //
        // GET: /Order/
        public ActionResult Index()
        {
            return View();
        }

        #region 投诉
        public ActionResult ComplaintInfo(string st, string et, string cid, string cpro, string cnum, string cancel, int page = 0, int limit = 0)
        {
            if (page != 0)
            {
                ResultJsonInfo r = new ResultJsonInfo();
                DataTable data = new DataTable();
                try
                {
                    string where = " 1=1 ";
                    if (!string.IsNullOrWhiteSpace(st)) where = where + " and ComplainDate>='" + st + "'";
                    if (!string.IsNullOrWhiteSpace(et)) where = where + " and ComplainDate<='" + et + "'";
                    if (!string.IsNullOrWhiteSpace(cid)) where = where + " and ComplaintID like'%" + cid + "%'";
                    if (!string.IsNullOrWhiteSpace(cpro)) where = where + " and ComplainProduct like'%" + cpro + "%'";
                    if (!string.IsNullOrWhiteSpace(cnum)) where = where + " and (ComplaintNum like'%" + cnum + "%' or ComplainPhone like'%" + cnum + "%')";
                    if (!string.IsNullOrWhiteSpace(cancel)) where = where + " and Cancel=" + cancel;
                    r.count = cdata.CountOrder_Complaint(where);
                    if (r.count > 0) data = cdata.TableOrder_Complaint(where, page, limit);
                    else
                    {
                        r.msg = "暂无数据";
                    }
                }
                catch (Exception ex)
                {
                    r.code = -1;
                    r.msg = ex.Message;
                }
                r.data = data;
                return Content(JsonConvert.SerializeObject(r));
            }
            else
            {
                ViewData["Title"] = "投诉清单管理";
                return View();
            }
        }
        public ActionResult Complaintdel(string ids)
        {
            var mes = "删除失败！";
            try
            {
                if (cdata.UpdateStatusOrder_Complaint(" ID in(" + ids + ")", "Status=0")) mes = "OK";
            }
            catch (Exception ex) { }
            return Content(mes);
        }
        public ActionResult ComplaintImport(string filepath)
        { 
            try
            {
                string result = "";
                bool r = FileUtil.ExcelImport(1, 1, filepath,"", ref result);
                return Content(FileUtil.ToJson(new { result = r, data = result }));
            }
            catch (Exception ex)
            {
                return Content(FileUtil.ToJson(new { result = false, data = ex.Message }));
            }
        }
        public ActionResult ComplaintSave(Order_Complaint m)
        {
            var mes = "保存失败！";
            try
            {
                bool r = false;
                m.Status = 1;
                m.Time = DateTime.Now;
                Order_ComplaintDAL data1=new Order_ComplaintDAL();
                if (string.IsNullOrWhiteSpace(m.ID))
                {

                    r = data1.InsertOrder_Complaint(m);
                }
                else
                {
                    r = data1.UpdateOrder_Complaint(m);
                }
                if (r) { mes = "OK"; }
            }
            catch (Exception ex) { }
            return Content(mes);
        }  
        public ActionResult ComplaintToTable(string st, string et, string cid, string cpro, string cnum, string cancel)
        {
            try
            {
                string result = "";
                string where = " 1=1 ";
                if (!string.IsNullOrWhiteSpace(st)) where = where + " and ComplainDate>='" + st + "'";
                if (!string.IsNullOrWhiteSpace(et)) where = where + " and ComplainDate<='" + et + "'";
                if (!string.IsNullOrWhiteSpace(cid)) where = where + " and ComplaintID like'%" + cid + "%'";
                if (!string.IsNullOrWhiteSpace(cpro)) where = where + " and ComplainProduct like'%" + cpro + "%'";
                if (!string.IsNullOrWhiteSpace(cnum)) where = where + " and (ComplaintNum like'%" + cnum + "%' or ComplainPhone like'%" + cnum + "%')";
                if (!string.IsNullOrWhiteSpace(cancel)) where = where + " and Cancel=" + cancel;
                bool r = FileUtil.ExcelToTable(1,"", where, ref result);
                return Content(FileUtil.ToJson(new { result = r, data = result }));
            }
            catch (Exception ex)
            {
                return Content(FileUtil.ToJson(new { result = false, data = ex.Message }));
            }
        }
        #endregion

     
    }
}