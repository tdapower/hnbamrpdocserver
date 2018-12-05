using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Net.Http.Headers;
using System.IO;
using System.Data.OracleClient;
using System.Data;
using System.Configuration;

namespace MRPSDocumentServer.Controllers
{
    public class MedicalLetterController : ApiController
    {
        static string ConnectionString = ConfigurationManager.ConnectionStrings["OracleConString"].ToString();


        [HttpGet]
        [ActionName("GetMedicalLetterCreditDocument")]
        public HttpResponseMessage GetMedicalLetterCreditDocument(int id)
        {
            HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK, "value");
            TableLogOnInfo crTableLogOnInfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();
            ReportDocument crystalReport = new ReportDocument();
            crystalReport.Load(System.Web.Hosting.HostingEnvironment.MapPath("~/Documents/MedicalLetterCredit/MedicalLetterCredit.rpt"));
            crystalReport.SetDatabaseLogon("hnba_crc", "HNBACRC", System.Configuration.ConfigurationManager.AppSettings["REPORT_DB_SERVER_NAME"].ToString(), "");
            crystalReport.SetParameterValue("MedicalLetterId", id);
            Stream oStream = null;
            byte[] byteArray = null;
            oStream = crystalReport.ExportToStream(ExportFormatType.PortableDocFormat);
            byteArray = new byte[oStream.Length];
            oStream.Read(byteArray, 0, Convert.ToInt32(oStream.Length - 1));
            response.Content = new ByteArrayContent(byteArray);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

            SaveDocumentToDB(id, GetCrystalPDF(crystalReport));
            return response;

        }


        [HttpGet]
        [ActionName("GetMedicalLetterNonCreditDocument")]
        public HttpResponseMessage GetMedicalLetterNonCreditDocument(int id)
        {
            HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK, "value");
            TableLogOnInfo crTableLogOnInfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();
            ReportDocument crystalReport = new ReportDocument();
            crystalReport.Load(System.Web.Hosting.HostingEnvironment.MapPath("~/Documents/MedicalLetterNonCredit/MedicalLetterNonCredit.rpt"));
            crystalReport.SetDatabaseLogon("hnba_crc", "HNBACRC", System.Configuration.ConfigurationManager.AppSettings["REPORT_DB_SERVER_NAME"].ToString(), "");
            crystalReport.SetParameterValue("MedicalLetterId", id);
            Stream oStream = null;
            byte[] byteArray = null;
            oStream = crystalReport.ExportToStream(ExportFormatType.PortableDocFormat);
            byteArray = new byte[oStream.Length];
            oStream.Read(byteArray, 0, Convert.ToInt32(oStream.Length - 1));
            response.Content = new ByteArrayContent(byteArray);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            
            SaveDocumentToDB(id, GetCrystalPDF(crystalReport));
            return response;

        }

        [HttpGet]
        [ActionName("GetFurtherMedicalLetterDocument")]
        public HttpResponseMessage GetFurtherMedicalLetterDocument(int id)
        {
            HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK, "value");
            TableLogOnInfo crTableLogOnInfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();
            ReportDocument crystalReport = new ReportDocument();
            crystalReport.Load(System.Web.Hosting.HostingEnvironment.MapPath("~/Documents/FurtherMedicalLetter/FurtherMedicalLetter.rpt"));
            crystalReport.SetDatabaseLogon("hnba_crc", "HNBACRC", System.Configuration.ConfigurationManager.AppSettings["REPORT_DB_SERVER_NAME"].ToString(), "");
            crystalReport.SetParameterValue("MedicalLetterId", id);
            Stream oStream = null;
            byte[] byteArray = null;
            oStream = crystalReport.ExportToStream(ExportFormatType.PortableDocFormat);
            byteArray = new byte[oStream.Length];
            oStream.Read(byteArray, 0, Convert.ToInt32(oStream.Length - 1));
            response.Content = new ByteArrayContent(byteArray);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");


            SaveDocumentToDB(id, GetCrystalPDF(crystalReport));
            return response;

        }



        public void SaveDocumentToDB(int medicalLetterId, byte[] doc)
        {
            try
            {
                OracleConnection conProcess = new OracleConnection(ConnectionString);
                conProcess.Open();
                OracleCommand spProcess = null;
                int AppCode = 0;
                string strQuery = "";
                
                OracleParameter blobParameterDoc = new OracleParameter();
                
                strQuery = "UPDATE  MRPSMedicalLetter  SET Document=:doc ";
                strQuery += " WHERE SeqId=:MedicalLetterId ";

                blobParameterDoc.ParameterName = "doc";
                blobParameterDoc.Direction = ParameterDirection.Input;
                blobParameterDoc.Value = doc;
                

                spProcess = new OracleCommand(strQuery, conProcess);
                spProcess.Parameters.Add(blobParameterDoc);
                spProcess.Parameters.Add("MedicalLetterId", OracleType.Number).Value = medicalLetterId;

                spProcess.ExecuteNonQuery();
                conProcess.Close();
                conProcess.Dispose();
            }
            catch (Exception ex)
            {

            }
        }


        private byte[] GetCrystalPDF(ReportDocument rpt)
        {
            byte[] buffer = null;
            try
            {
                using (var sr = new StreamReader(rpt.ExportToStream(ExportFormatType.PortableDocFormat)))
                {
                    buffer = new byte[sr.BaseStream.Length];
                    sr.BaseStream.Read(buffer, 0, buffer.Length);
                    rpt.Close();
                    rpt.Dispose();
                    rpt = null;
                }
            }
            catch (Exception ex)
            {
                // blah
            }
            return buffer;
        }

    }
}
