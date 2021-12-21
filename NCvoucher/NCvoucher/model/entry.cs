using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NCvoucher
{
    class entry
    {
        private string code;

        public string Code
        {
            get { return code; }
            set { code = value; }
        }
        private string zy;

        public string Zy
        {
            get { return zy; }
            set { zy = value; }
        }
        private int debit;

        public int Debit
        {
            get { return debit; }
            set { debit = value; }
        }
        private int credit;

        public int Credit
        {
            get { return credit; }
            set { credit = value; }
        }
        private string cashflow;

        public string Cashflow
        {
            get { return cashflow; }
            set { cashflow = value; }
        }
        private int money;

        public int Money
        {
            get { return money; }
            set { money = value; }
        }
        private List<auxiliary> auxiliaryList = new List<auxiliary>();

        internal List<auxiliary> AuxiliaryList
        {
            get { return auxiliaryList; }
            set { auxiliaryList = value; }
        }
        private List<cashflowcase> cashflowcaseList = new List<cashflowcase>();

        internal List<cashflowcase> CashflowcaseList
        {
            get { return cashflowcaseList; }
            set { cashflowcaseList = value; }
        }
    }
}
