using System;
using System.Collections.Generic;
using System.Text;

namespace OPM.OPMEnginee
{
    class _Device
    {
        private string _serial;
        private string _id_storage;
        private string _MAC;
        private string _serial_gpon;
        private string _device_name;
        private string _note;

        public _Device()
        {

        }
         ~_Device()
        {

        }
        public string serial
        {
            get;
            set;
        }
        public string id_storage
        {
            get;
            set;
        }
        public string MAC
        {
            get;
            set;
        }
        public string serial_gpon
        {
            get;
            set;
        }
        public string device_name
        {
            get;
            set;
        }
    }
}
