using System;
using System.Collections.Generic;
using System.Text;

namespace OPM.OPMEnginee
{
    class Packagelist
    {
        private string _year;
        private string _po_number;
        private string _dp_number;
        private string _province;
        private int _number;
        private string _type;
        private List<string> _serial =  new List<string>();

        public Packagelist()
        {
            _type = "Hang_Chinh";
        }
        ~Packagelist()
        {

        }
        
        public string DP
        {
            set { _dp_number = value; }
            get { return _dp_number; }
        }
        public string Year
        {
            set { _year = value; }
            get { return _year; }
        }
        public string PO_number
        {
            set { _po_number = value; }
            get { return _po_number; }
        }
        public string Province
        {
            set { _province = value; }
            get { return _province; }
        }
        public int Numberdevice
        {
            set { _number = value; }
            get { return _number; }
        }
        public string Type
        {
            set { _type = value; }
            get { return _type; }
        }
        public void SetSerial(string strItem)
        {
            _serial.Add(strItem);
        }
        public string GetItem(int index)
        {
            return _serial[index];
        }
        
        public List<string> GetListSerial()
        {
            List<string> xlCloneSerial = new List<string>(_serial);
            return xlCloneSerial;
        }
    }

}
