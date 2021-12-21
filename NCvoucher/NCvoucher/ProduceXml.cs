using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NCvoucher
{
    class ProduceXml
    {
        public static string GetXml(JArray arrList, JObject detData)
        {
            DateTime nowdate = DateTime.Now;
            DateTime.TryParse( detData["prepareddate"].ToString(),out nowdate);
            XmlDocument docheader = new XmlDocument();
            XmlElement header = docheader.CreateElement("voucher_head");
            header.AppendChild(docheader.CreateElement("company")).InnerText = detData["company"].ToString();
            header.AppendChild(docheader.CreateElement("voucher_type")).InnerText = detData["voucher_type"].ToString();
            header.AppendChild(docheader.CreateElement("fiscal_year")).InnerText = nowdate.ToString("yyyy");
            header.AppendChild(docheader.CreateElement("accounting_period")).InnerText = nowdate.ToString("MM");
            header.AppendChild(docheader.CreateElement("voucher_id")).InnerText = "1";
            header.AppendChild(docheader.CreateElement("attachment_number")).InnerText = "-1";
            header.AppendChild(docheader.CreateElement("prepareddate")).InnerText = detData["prepareddate"].ToString();
            header.AppendChild(docheader.CreateElement("enter")).InnerText = detData["enter"].ToString();
            header.AppendChild(docheader.CreateElement("cashier")).InnerText = "";
            header.AppendChild(docheader.CreateElement("signature")).InnerText = "";
            header.AppendChild(docheader.CreateElement("checker")).InnerText = "";
            header.AppendChild(docheader.CreateElement("posting_date")).InnerText = "";
            header.AppendChild(docheader.CreateElement("posting_person")).InnerText = "";
            header.AppendChild(docheader.CreateElement("voucher_making_system")).InnerText = "外部系统交换平台";
            header.AppendChild(docheader.CreateElement("memo1")).InnerText = "";
            header.AppendChild(docheader.CreateElement("memo2")).InnerText = "";
            header.AppendChild(docheader.CreateElement("reserve1")).InnerText = "";
            header.AppendChild(docheader.CreateElement("reserve2")).InnerText = "GL0000000000019";
            header.AppendChild(docheader.CreateElement("revokeflag")).InnerText = "";
            
            List<entry> list = new List<entry>();
            for (int j = 0; j < arrList.Count; j++)
            {
                entry _entry = new entry();
                _entry.Zy = arrList[j]["摘要"].ToString();
                _entry.Code = arrList[j]["科目编码"].ToString();
                //加一个借方和贷方的判断条件
                if (arrList[j]["贷方"].ToString()=="0")
                {
                    _entry.Debit = int.Parse(arrList[j]["借方"].ToString());
                    
                    auxiliary _auxCredit = new auxiliary();
                    _auxCredit.Name = "费用属性辅助核算";
                    _auxCredit.Val = detData["Val1"].ToString();
                    _entry.AuxiliaryList.Add(_auxCredit);

                    auxiliary _auxCredit1 = new auxiliary();
                    _auxCredit1.Name = "业务员";
                    _auxCredit1.Val = detData["Val2"].ToString();
                    _entry.AuxiliaryList.Add(_auxCredit1);

                    //cashflowcase cash = new cashflowcase();
                    //cash.Cash_item = "01";
                    //cash.Natural_debit_currency = arrList[j]["借方"].ToString();
                    //cash.Natural_credit_currency = arrList[j]["贷方"].ToString();
                    //_entry.CashflowcaseList.Add(cash);

                    //coderemarkcase code = new coderemarkcase();
                    //code.Name = "部门档案";
                    //code.Val = "001";
                    //_entry.CoderemarkcaseList.Add(code);
                }
                else {
                    _entry.Credit = int.Parse(arrList[j]["贷方"].ToString());
                    
                    auxiliary _auxCredit = new auxiliary();
                    _auxCredit.Name = "银行类别";
                    _auxCredit.Val = detData["Val3"].ToString();
                    _entry.AuxiliaryList.Add(_auxCredit);

                    auxiliary _auxCredit1 = new auxiliary();
                    _auxCredit1.Name = "银行档案";
                    _auxCredit1.Val = detData["Val4"].ToString();
                    _entry.AuxiliaryList.Add(_auxCredit1);

                    auxiliary _auxCredit2 = new auxiliary();
                    _auxCredit2.Name = "银行账户";
                    _auxCredit2.Val = detData["Val5"].ToString();
                    _entry.AuxiliaryList.Add(_auxCredit2);

                    //cashflowcase cash = new cashflowcase();
                    //cash.Cash_item = "01";
                    //cash.Natural_debit_currency = arrList[j]["借方"].ToString();
                    //cash.Natural_credit_currency = arrList[j]["贷方"].ToString();
                    //_entry.CashflowcaseList.Add(cash);
                    cashflowcase cash = new cashflowcase();
                    cash.Cashflow = "支付的其他与经营活动有关的现金";
                    cash.Money = int.Parse(arrList[j]["贷方"].ToString());
                    _entry.CashflowcaseList.Add(cash);

                    //coderemarkcase code = new coderemarkcase();
                    //code.Name = "部门档案";
                    //code.Val = "002";
                    //_entry.CoderemarkcaseList.Add(code);
                }
                
                list.Add(_entry);
            }

            //entry _entry = new entry();//借方
            //_entry.Code = "660110";
            //_entry.Zy = "付讲师费用";
            //_entry.Debit = 100;

            //auxiliary _auxCredit = new auxiliary();
            //_auxCredit.Name = "dept_id";
            //_auxCredit.Val = "001";
            //_entry.AuxiliaryList.Add(_auxCredit);

            //auxiliary _auxCredit1 = new auxiliary();
            //_auxCredit1.Name = "personnel_id";
            //_auxCredit1.Val = "";
            //_entry.AuxiliaryList.Add(_auxCredit1);

            //list.Add(_entry);

            //entry _entry2 = new entry();//贷方
            //_entry2.Code = "100202";
            //_entry2.Zy = "付讲师费用";
            //_entry2.Credit = 100;

            //cashflowcase cash = new cashflowcase();
            //cash.Cash_item = "01";
            //cash.Natural_debit_currency = "48.06";
            //cash.Natural_credit_currency = "0.00";
            //_entry2.CashflowcaseList.Add(cash);


            //coderemarkcase code = new coderemarkcase();
            //code.Name = "部门档案";
            //code.Val = "001";
            //coderemarkcase code1 = new coderemarkcase();
            //code1.Name = "客户辅助核算";
            //code1.Val = "";
            //_entry2.CoderemarkcaseList.Add(code);
            //_entry2.CoderemarkcaseList.Add(code1);

            //auxiliary _aux = new auxiliary();
            //_aux.Name = "dept_id";
            //_aux.Val = "002";

            //auxiliary _aux1 = new auxiliary();
            //_aux1.Name = "personnel_id";
            //_aux1.Val = "";

            //auxiliary _aux2 = new auxiliary();
            //_aux2.Name = "cust_id";
            //_aux2.Val = "";

            //_entry2.AuxiliaryList.Add(_aux);
            //_entry2.AuxiliaryList.Add(_aux1);
            //_entry2.AuxiliaryList.Add(_aux2);

            //list.Add(_entry2);

            XmlElement body = docheader.CreateElement("voucher_body");
            int i = 1;
            foreach (entry en in list)
            {
                //凭证明细            
                XmlNode entry = docheader.CreateElement("entry");
                entry.AppendChild(docheader.CreateElement("entry_id")).InnerText = i.ToString();
                entry.AppendChild(docheader.CreateElement("account_code")).InnerText = en.Code;
                entry.AppendChild(docheader.CreateElement("abstract")).InnerText = en.Zy;
                entry.AppendChild(docheader.CreateElement("settlement")).InnerText = "";
                entry.AppendChild(docheader.CreateElement("document_id")).InnerText = "";
                entry.AppendChild(docheader.CreateElement("document_date")).InnerText = "";
                entry.AppendChild(docheader.CreateElement("currency")).InnerText = "人民币";
                entry.AppendChild(docheader.CreateElement("unit_price")).InnerText = "";
                entry.AppendChild(docheader.CreateElement("exchange_rate1")).InnerText = "";
                entry.AppendChild(docheader.CreateElement("exchange_rate2")).InnerText = "0";
                entry.AppendChild(docheader.CreateElement("debit_quantity")).InnerText = "0";
                entry.AppendChild(docheader.CreateElement("primary_debit_amount")).InnerText = en.Debit.ToString(); //原币借方发生额
                entry.AppendChild(docheader.CreateElement("secondary_debit_amount")).InnerText = "";//辅币借方发生额
                entry.AppendChild(docheader.CreateElement("natural_debit_currency")).InnerText = en.Debit.ToString();//本币借方发生额
                entry.AppendChild(docheader.CreateElement("credit_quantity")).InnerText = "0";
                entry.AppendChild(docheader.CreateElement("primary_credit_amount")).InnerText = en.Credit.ToString();//原币贷方发生额
                entry.AppendChild(docheader.CreateElement("secondary_credit_amount")).InnerText = "";//辅币
                entry.AppendChild(docheader.CreateElement("natural_credit_currency")).InnerText = en.Credit.ToString();//本币贷方发生额
                entry.AppendChild(docheader.CreateElement("bill_type")).InnerText = "";
                entry.AppendChild(docheader.CreateElement("bill_id")).InnerText = "";
                entry.AppendChild(docheader.CreateElement("bill_date")).InnerText = "";

                //XmlNode auxiliary = docheader.CreateElement("auxiliary_accounting");
                //auxiliary.InnerText = "";
                //foreach (auxiliary au in en.AuxiliaryList)
                //{
                //    XmlElement item = docheader.CreateElement("item");
                //    item.SetAttribute("name", au.Name);
                //    item.InnerText = au.Val;
                //    auxiliary.AppendChild(item);
                //}
                //entry.AppendChild(auxiliary);

                //XmlNode detail = docheader.CreateElement("detail");
                //detail.InnerText = "";
                
                //XmlNode cashflowstatement = docheader.CreateElement("cash_flow_statement");
                //cashflowstatement.InnerText = "";
                //XmlElement cashflow = docheader.CreateElement("cash_flow");
                //foreach (cashflowcase ca in en.CashflowcaseList)
                //{
                //    cashflow.SetAttribute("cash_item", ca.Cash_item);
                //    cashflow.SetAttribute("natural_debit_currency", ca.Natural_credit_currency);
                //    cashflow.SetAttribute("natural_credit_currency", ca.Natural_credit_currency);
                //    cashflowstatement.AppendChild(cashflow);
                //}
                //detail.AppendChild(cashflowstatement);
                    
                //XmlNode coderemarkstatement = docheader.CreateElement("code_remark_statement");
                //coderemarkstatement.InnerText = "";
                //XmlElement coderemark = docheader.CreateElement("code_remark");
                //coderemark.SetAttribute("i_id", "1");
                //foreach (coderemarkcase cr in en.CoderemarkcaseList)
                //{
                //    XmlElement item1 = docheader.CreateElement("item");
                //    item1.SetAttribute("name", cr.Name);
                //    item1.InnerText = cr.Val;
                //    coderemark.AppendChild(item1);
                //}
                //if (en.CoderemarkcaseList.Count > 0)
                //    coderemarkstatement.AppendChild(coderemark);
                //detail.AppendChild(coderemarkstatement);
                
                //entry.AppendChild(detail);

                XmlNode auxiliary = docheader.CreateElement("auxiliary_accounting");
                foreach (auxiliary au in en.AuxiliaryList)
                {
                    XmlElement item = docheader.CreateElement("item");
                    item.SetAttribute("name", au.Name);
                    item.InnerText = au.Val;
                    auxiliary.AppendChild(item);
                }
                if (en.AuxiliaryList.Count > 0)
                    entry.AppendChild(auxiliary);

                XmlNode userdata = docheader.CreateElement("otheruserdata");

                foreach (cashflowcase item in en.CashflowcaseList)
                {
                    XmlNode cashflowcase = docheader.CreateElement("cashflowcase");
                    cashflowcase.AppendChild(docheader.CreateElement("money")).InnerText = item.Money.ToString();
                    cashflowcase.AppendChild(docheader.CreateElement("moneyass")).InnerText = item.Money.ToString();
                    cashflowcase.AppendChild(docheader.CreateElement("moneymain")).InnerText = item.Money.ToString();
                    cashflowcase.AppendChild(docheader.CreateElement("pk_cashflow")).InnerText = item.Cashflow;
                    userdata.AppendChild(cashflowcase);
                }

                if (en.CashflowcaseList.Count > 0)
                    entry.AppendChild(userdata);
                body.AppendChild(entry);
                i++;
            }
            StringBuilder xml = new StringBuilder();
            xml.Append("<?xml version=\"1.0\" encoding='UTF-8'?>");
            //xml.Append("<ufinterface roottag=\"voucher\" billtype=\"gl\" docid=\"\" receiver=\"u8\" sender=\"001\" proc=\"Query\" isexchange=\"Y\" codeexchanged=\"N\" renewproofno=\"n\" exportneedexch=\"N\" timestamp=\"0x000000000007303B\">");
            xml.Append("<ufinterface roottag=\"voucher\" billtype=\"gl\" replace=\"Y\" receiver=\"ce@ce-0001\" sender=\"001\" isexchange=\"Y\" filename=\"凭证样本数据文件.xml\" proc=\"add\" operation=\"req\">");
            XmlElement voucher = docheader.CreateElement("voucher");
            voucher.SetAttribute("id", "GL0000000000019");
            voucher.AppendChild(header);
            voucher.AppendChild(body);
            xml.Append(voucher.OuterXml);
            xml.Append("</ufinterface>");
            //header.SetAttribute("cname", "测试");
            //CallBack(xml.ToString());

            //生成XML
            return xml.ToString();
        }
    }
}
