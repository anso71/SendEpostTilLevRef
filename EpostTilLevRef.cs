using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Agresso.ServerExtension;
using Agresso.Interface.CommonExtension;
using System.Net;
using System.Net.Mail;

namespace HLS_SendEpostTilLevFeilRef
{
    [Report("EIN01", "*", "Sende epost til  leverandør for feil Deres ref")]
    public class EpostTilLevRef : IProjectServer
    {
        private IReport _ein01;
        private String _client;
        private String _emailaddress;
        public void Initialize(IReport ireport)
        {
            _ein01 = ireport;
            _client = _ein01.API.GetParameter("client");
            _emailaddress = _ein01.API.GetParameter("emailAdress");


            _ein01.OnStop += _report_OnStop;
        }
        private void _report_OnStop(object sender, ReportEventArgs e)
        {
            IStatement sqlwrongReffromAparid = CurrentContext.Database.CreateStatement();
            DataTable WrongApri = new DataTable("WrongRef");
            IDictionary<int, string> Sendt = new Dictionary<int, string>();

            sqlwrongReffromAparid.Append("SELECT a.apar_id, b.apar_name, a.arrival_date, a.bank_account, a.batch_id, a.comments, a.contract_id, a.cur_amount, a.client, a.comp_reg_no,");
            sqlwrongReffromAparid.Append("a.currency, a.due_date, a.ext_inv_ref, a.ext_ord_ref, a.invoice_id, a.kid,a.last_update, a.order_id, a.reg_amount, a.status, a.trans_date, a.tax_cur_amt,");
            sqlwrongReffromAparid.Append("a.queue_id, a.accounting_cost, a.contact_id, a.order_reference, a.contract_reference,a.metering_point_id, a.prepaid_amount, a.rounding_amount FROM a47en53invoiceheader");
            sqlwrongReffromAparid.Append(" a JOIN acrclient c ON a.client = c.client AND c.client IN(@client)");
            sqlwrongReffromAparid.Append(" LEFT OUTER JOIN asuheader b ON a.apar_id = b.apar_id AND b.client = c.pay_client AND");
            sqlwrongReffromAparid.Append(" a.client = c.client AND b.status = 'N' WHERE a.arrival_date >= @start_Date");
            sqlwrongReffromAparid.Append(" AND a.arrival_date <= @end_date");
            sqlwrongReffromAparid.Append(" AND a.status = 'N'  AND a.invoice_id IN(SELECT DISTINCT invoice_id ");
            sqlwrongReffromAparid.Append("FROM a47en53log WHERE status = 'N' AND lower(log_type) = 'warning' AND client IN(@client)) ");
            sqlwrongReffromAparid["client"] = _client;
            sqlwrongReffromAparid["start_Date"] = Convert.ToDateTime("1900-01-01");
            sqlwrongReffromAparid["end_date"] = Convert.ToDateTime("2100-01-01");
            _ein01.API.WriteLog(sqlwrongReffromAparid.ToString());
            _ein01.API.WriteLog(_client);
            _ein01.API.WriteLog(Convert.ToDateTime("1900-01-01").ToString());
            _ein01.API.WriteLog(Convert.ToDateTime("2100-01-01").ToString());

            CurrentContext.Database.Read(sqlwrongReffromAparid, WrongApri);
            StringBuilder ReportText = new StringBuilder();
            StringBuilder ReportNotText = new StringBuilder();

            ReportText.Append("Epost sendt til følgende kunder: \r\n");
            ReportNotText.Append("Epost manglet på følgende kunder: \r\n");
            foreach (DataRow row in WrongApri.Rows)
            {
                IStatement sqlFindEmailfromAparid = CurrentContext.Database.CreateStatement();
                sqlFindEmailfromAparid.Append("select e_mail from agladdress where dim_value = @apar_id and client = @client and attribute_id = 'A5'");
                sqlFindEmailfromAparid["apar_id"] = row["apar_id"];
                sqlFindEmailfromAparid["client"] = _client;
                string email_address = "";
                CurrentContext.Database.ReadValue(sqlFindEmailfromAparid, ref email_address);
                if (!string.IsNullOrEmpty(email_address))
                {
                    if (Int32.TryParse(row["apar_id"].ToString(), out int kundenr))
                    {
                        if (!Sendt.ContainsKey(kundenr))
                        {
                            StringBuilder EpostText = new StringBuilder();
                            EpostText.Append("Hei\r\n  Du får denne mailen da Deres Ref er feil på ordre: ");
                            EpostText.Append(row["invoice_id"]);
                            EpostText.Append("\r\nMed forfall den: ");
                            EpostText.Append(row["trans_date"]);
                            EpostText.Append("\r\n \"Deres ref\" var satt til: ");
                            EpostText.Append(row["a.ext_ord_ref"]);
                            EpostText.Append("\r\n  Vi bruker ressurnr:  6 siffer");
                            EpostText.Append("\r\n Vi håper Dere kan sette inn korrekt \"Deres Ref\" på neste regning\r\n");
                            EpostText.Append("Med vennlig hilsen \r\n");
                            IStatement sqldescription = CurrentContext.Database.CreateStatement();
                            sqldescription.Append("select description from agldescription where dim_value= @client and client = @client and attribute_id='A3'");
                            sqldescription["client"] = _client;
                            string description = "";
                            CurrentContext.Database.ReadValue(sqldescription, ref description);
                            EpostText.Append("Fakturaavdelingen \r\n");
                            EpostText.Append(description);
                            Sendt.Add(kundenr, row["apar_id"].ToString());
                            _ein01.API.WriteLog(EpostText.ToString()); // for test
                            try
                            {
                                if (_ein01.API.SendMail(EpostText.ToString(), "", "Feil i Deres Ref", "Feil i Deres Ref", "andre.sollie@stange.kommune.no", ""))
                                {
                                    ReportText.Append(email_address);
                                    ReportText.Append(" Kundnr: ");
                                    ReportText.Append(row["apar_id"]);
                                    ReportText.Append(" gjelder ordernr: ");
                                    ReportText.Append(row["invoice_id"]);
                                    ReportText.Append("\n");
                                }
                                else
                                {
                                    _ein01.API.WriteLog("Feil med utsendelse av mail {0}", email_address);
                                }
                            }
                            catch (IOException ei)
                            {
                                _ein01.API.WriteLog("Epost ikke sendt grunnet exception :{0} {1}", email_address, ei.Source);
                            }
                        }
                        else
                        {
                            _ein01.API.WriteLog("{0} har blitt sendt epost til", kundenr);
                        }
                    }
                }
                else
                {
                    ReportNotText.Append("Kundenr: ");
                    ReportNotText.Append(row["apar_id"]);
                    ReportNotText.Append("\r\n");
                }

            }
            ReportText.Append("\r\n\r\n");
            ReportText.Append(ReportNotText.ToString());
            _ein01.API.WriteLog(ReportText.ToString());
            try
            {
                if (_ein01.API.SendMail(ReportText.ToString(), "", "Report på hva som er sendt", "Report på hva som er sendt", "andre.sollie@stange.kommune.no", ""))
                {
                    _ein01.API.WriteLog("Report sendt");
                }
                else
                {
                    _ein01.API.WriteLog("Report ikke sendt");
                }
            }
            catch (IOException ei1)
            {
                _ein01.API.WriteLog("Report ikke sendt grunnet exeception : {0}", ei1.Source);
            }

        }
    }
}
