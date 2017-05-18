using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EstimatesAssembly {
    public class initBookmark {

        public Dictionary<string, string> fillBookmark() {
            Dictionary<string, string> mapBookmark = new Dictionary<string, string>();
            mapBookmark.Add("свидетельство", MainFormAsm.iniSet.TbCertificate);
            mapBookmark.Add("наименование_заказчика", MainFormAsm.iniSet.TbCustomer);
            mapBookmark.Add("наименование_стройки", MainFormAsm.iniSet.TbNameBuilding);
            mapBookmark.Add("наименование_объекта", MainFormAsm.iniSet.TbNameObject);
            mapBookmark.Add("номер_раздела", MainFormAsm.iniSet.TbSectionNumber);
            mapBookmark.Add("шифр", MainFormAsm.iniSet.TbCodeObject);
            mapBookmark.Add("том", MainFormAsm.iniSet.NumVolumeNumber);
            mapBookmark.Add("всего_томов", MainFormAsm.iniSet.TbVolCount);
            mapBookmark.Add("подпись_руководителя", "");
            mapBookmark.Add("должность_руководителя", MainFormAsm.iniSet.TbChiefPsition);
            mapBookmark.Add("фио_гип", MainFormAsm.iniSet.TbGipFio);
            mapBookmark.Add("фио_руководителя", MainFormAsm.iniSet.TbChiefFio);
            mapBookmark.Add("подпись_гип", "");
            mapBookmark.Add("год", MainFormAsm.iniSet.TbYearTitle);
            return mapBookmark;
        }
    }
}
