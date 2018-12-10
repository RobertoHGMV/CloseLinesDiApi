using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace CloseLinesDiApi
{
    public class AdonService
    {
        public Company Company { get; private set; }

        public string Server { get; set; }
        public string CompanyDB { get; set; }
        public string DbUserName { get; set; }
        public string DbPassword { get; set; }
        public BoDataServerTypes DbServerType { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }

        public AdonService()
        {
            Company = new Company();
        }

        public void ConnectSbo()
        {
            Company.Server = Server;
            Company.CompanyDB = CompanyDB;
            Company.DbUserName = DbUserName;
            Company.DbPassword = DbPassword;
            Company.DbServerType = BoDataServerTypes.dst_MSSQL2014;
            Company.UserName = UserName;
            Company.Password = Password;

            Company.UseTrusted = false;
            Company.language = BoSuppLangs.ln_Portuguese_Br;
            Company.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;

            if (Company.Connect() != 0)
                throw new Exception("Erro ao conectar no sap:" +
                           $"[{Company.GetLastErrorCode()}] - [{Company.GetLastErrorDescription()}]");
        }

        public void CloseLinesOfQuotation(int docEntry)
        {
            var businessObject = Company.GetBusinessObject(BoObjectTypes.oQuotations) as Documents;

            if (!businessObject.GetByKey(docEntry))
                throw new Exception($"Não foi possível localizar a cotação N°[{docEntry}] no SAP.");

            CloseLines(businessObject);
        }

        private void CloseLines(Documents businessObject)
        {
            for (var i = 0; i < businessObject.Lines.Count; i++)
            {
                businessObject.Lines.SetCurrentLine(i);

                if (businessObject.Lines.LineStatus == BoStatus.bost_Open)
                {
                    businessObject.Lines.LineStatus = BoStatus.bost_Close;

                    if (businessObject.Update() != 0)
                        throw new Exception($"Erro ao atualizar documento no SAP.\n[{Company.GetLastErrorCode()}]-[{Company.GetLastErrorDescription()}]");
                }
            }
        }
    }
}
