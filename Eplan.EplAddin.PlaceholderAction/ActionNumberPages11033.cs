// © Mishkin_Ivan@mail.ru 
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.HEServices;
using System;
using System.Linq;
using System.Collections.Generic;
using Eplan.EplApi.DataModel;

namespace Eplan.EplAddin.PlaceholderAction {
    /// <summary>
    /// Нумерация страниц тома (томов). Выделить можно любую страницу проекта или несколько проектов сразу.
    /// </summary>
    class NumberPagesInProperty11033 : IEplAction {
        public bool OnRegister(ref string Name, ref int Ordinal) {
            // Имя для регистрации в системе
            Name = "NumberPagesIn11033";
            Ordinal = 22;
            return true; ;
        }
        public bool Execute(ActionCallingContext oActionCallingContext) {
            int countedPages = 0;
            SelectionSet set = new SelectionSet();
            var cursor = set.Selection.OfType<Page>().GetEnumerator();
            if (cursor.MoveNext())
            {
                using (Page firstPage = cursor.Current)
                {
                    countedPages = firstPage.Properties.PAGE_ADDITIONALSHEETNUMBER.ToInt(); 
                }
                while (cursor.MoveNext())
                {
                    using (Page currentPage = cursor.Current)
                    {
                        countedPages += 1;
                        currentPage.LockObject();
                        currentPage.Properties.PAGE_ADDITIONALSHEETNUMBER.Set(countedPages);
                        //currentPage.Dispose();
                    }
                }
            }
           
            return countedPages > 0;
        }


      
        public void GetActionProperties(ref ActionProperties actionProperties) {
            ;
        }

    }
}
