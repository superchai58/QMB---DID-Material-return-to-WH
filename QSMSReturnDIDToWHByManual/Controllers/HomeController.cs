using Connect.BLL;
using ExcelDataReader;
using QSMSReturnDIDToWHByManual.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Media;
using System.Web;
using System.Web.Mvc;
using static System.Net.Mime.MediaTypeNames;

namespace QSMSReturnDIDToWHByManual
{
    public class HomeController : Controller
    {
        DataTable dt = new DataTable();
        ConnectDB oCon = new ConnectDB();
        SqlCommand cmd = new SqlCommand();
        string resultChkDID = "";
        string result = "";

        // GET: Home

        //[HttpPost]
        public ActionResult Index()
        {
            Session.Clear();
            return View();
        }

        public ActionResult Register()
        {
            Session.Clear();
            return View();
        }

        [HttpPost]
        public ActionResult Register(QSMS_DIDCheckTempInfo model)   //Manual
        {
            result = "";
            resultChkDID = "";
            try
            {
                string did = model.DID.Trim();
                resultChkDID = QSMS_DID_Check(did);

                if (resultChkDID == "OK")
                {
                    dt = new DataTable();
                    dt.Clear();
                    dt.Columns.Add("DID");
                    object[] o = { did };
                    dt.Rows.Add(o);

                    result = QSMS_DID_ToWHStatus(dt);
                    ViewBag.message = result;
                }         
                else //Error message
                {
                    playNG();
                    ViewBag.message = resultChkDID;
                }
            }
            catch (Exception ex)
            {
                playNG();
                ViewBag.message = ex.ToString().Trim();
            }

            //--ConnectDB (End)--
            return View();
        }

        public string QSMS_DID_Check(string did)
        {
            string resultChk = "";

            dt = new DataTable();
            oCon = new ConnectDB();

            //--ConnectDB (Begin)--
            cmd = new SqlCommand();
            cmd.CommandText = "Select TOP 1 DID From QSMS_DIDCheckTemp with(nolock) Where DID = @DID";
            cmd.Parameters.Add(new SqlParameter("@DID", did));
            cmd.CommandTimeout = 180;
            dt = oCon.Query(cmd);

            if (dt.Rows.Count == 0)  //Update status for resend
            {                
                resultChk = "OK";
            }
            else //Error message
            {                
                resultChk = "DID: " + did + " Already manual resend to QWMS !!!";
            }

            return resultChk;
        }

        public string QSMS_DID_ToWHStatus(DataTable dt)
        {
            DataTable dtResult = new DataTable();
            string result = "";
            int flage = 0;
            foreach (DataRow row in dt.Rows)
            {
                try
                {
                    oCon = new ConnectDB();
                    cmd = new SqlCommand();
                    cmd.CommandText = "EXEC RollbackStatus_QSMS_DID_ToWH '" + row["DID"].ToString().Trim() + "'";
                    //cmd.Parameters.Add(new SqlParameter("@DID", row["DID"].ToString().Trim()));
                    cmd.CommandTimeout = 180;
                    dtResult = oCon.Query(cmd);

                    if (dtResult.Rows.Count > 0)
                    {
                        if (dtResult.Rows[0]["msg"].ToString().Trim() == "Fail")
                        {
                            result = "DID: " + row["DID"].ToString().Trim() + " format fail.";
                            flage = 1;
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    flage = 1;
                    result = ex.ToString().Trim();                    
                }
            }

            if (flage == 0)
            {
                playOK();
                result = "These DID already send to QWMS.";
            }
            else
            {
                playNG();
            }

            return result;
        }

        public ActionResult ImportExcel()
        {
            DataTable dt = new DataTable();

            try
            {         
                dt = (DataTable)Session["tmpdata"];
            }
            catch (Exception ex)
            {

            }

            return View(dt);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult ImportExcel(HttpPostedFileBase upload)
        {

            if (ModelState.IsValid)
            {

                if (upload != null && upload.ContentLength > 0)
                {
                    // ExcelDataReader works with the binary Excel file, so it needs a FileStream
                    // to get started. This is how we avoid dependencies on ACE or Interop:
                    Stream stream = upload.InputStream;

                    IExcelDataReader reader = null;


                    if (upload.FileName.EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (upload.FileName.EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        ModelState.AddModelError("File", "This file format is not supported");                       
                        playNG();
                        return View();
                    }
                    //Check Row, Column in excel
                    int fieldcount = reader.FieldCount;
                    int rowcount = reader.RowCount;

                    DataTable dt = new DataTable();
                    DataRow row;
                    DataTable dt_ = new DataTable();
                    try
                    {
                        dt_ = reader.AsDataSet().Tables[0];
                        for (int i = 0; i < dt_.Columns.Count; i++)
                        {
                            if (i == 0)
                            {
                                if (dt_.Rows[0][i].ToString().Trim() != "DID")
                                {                                    
                                    playNG();
                                    ViewBag.message = "Please Setting first column is 'DID' in excel file";
                                    return View();
                                }
                                else
                                {
                                    dt.Columns.Add(dt_.Rows[0][i].ToString());
                                }
                            }
                            else
                            {
                                dt.Columns.Add(dt_.Rows[0][i].ToString());
                            }
                        }
                        int rowcounter = 0;
                        for (int row_ = 1; row_ < dt_.Rows.Count; row_++)
                        {
                            row = dt.NewRow();

                            for (int col = 0; col < dt_.Columns.Count; col++)
                            {
                                row[col] = dt_.Rows[row_][col].ToString();
                                rowcounter++;
                            }
                            dt.Rows.Add(row);
                        }

                    }
                    catch (Exception ex)
                    {
                        ModelState.AddModelError("File", "Unable to Upload file!");
                        return View();
                    }

                    //--Resend these DID to QWMS (Begin)--
                    DataTable dtTmp = new DataTable();
                    dtTmp.Columns.Add("DID");
                    resultChkDID = "";
                    result = "";

                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow rowItem in dt.Rows)
                        {
                            resultChkDID = QSMS_DID_Check(rowItem["DID"].ToString().Trim());
                            if (resultChkDID == "OK")
                            {
                                dtTmp.Rows.Add(rowItem["DID"].ToString().Trim());
                            }
                            else
                            {
                                ViewBag.message = resultChkDID;
                                return View();
                            }
                        }
                    }

                    if (dtTmp.Rows.Count > 0)
                    {
                        result = QSMS_DID_ToWHStatus(dtTmp);

                        if(result.Substring(result.Length - 12) == "format fail.")
                        {
                            ViewBag.message = result;
                            return View();
                        }
                    }
                    //--Resend these DID to QWMS (End)--

                    //Send dt to grid
                    DataSet resultDs = new DataSet();
                    resultDs.Tables.Add(dt);
                    reader.Close();
                    reader.Dispose();
                    DataTable tmp = resultDs.Tables[0];
                    Session["tmpdata"] = tmp;  //store datatable into session
                    Session["time"] = 1;
                    //return RedirectToAction("ImportExcel");
                }
                else
                {
                    ModelState.AddModelError("File", "Please Upload Your file");
                }
            }
            return RedirectToAction("ImportExcel");
        }

        public void playOK()
        {
            string exePath = Server.MapPath(Url.Content("~/Sound/OK.wav"));
            SoundPlayer simpleSound = new SoundPlayer(exePath);
            simpleSound.Play();
            //return;
        }

        public void playNG()
        {
            string exePath = Server.MapPath(Url.Content("~/Sound/OO.wav"));
            SoundPlayer simpleSound = new SoundPlayer(exePath);
            simpleSound.Play();
            //return;
        }
    }
}