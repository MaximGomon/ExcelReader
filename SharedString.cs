using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ExcelReader
{
    public class SharedString
    {
        [XmlElement("t")]
        public string _text;

        public string Text
        {
            get
            {
                if (_text == null)
                {
                    return String.Empty;
                }
                else
                {
                    return _text;
                }
            }
        }

    }

    [Serializable()]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class SharedStringRoot
    {
        [XmlAttribute]
        public string UniqueCount;

        [XmlAttribute]
        public string Count;

        [XmlElement("si")]
        public SharedString[] AllStrings;
    }
}
