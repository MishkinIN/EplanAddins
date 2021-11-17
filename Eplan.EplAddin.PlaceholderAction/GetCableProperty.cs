// GNU General Public License v2.0 https://opensource.org/licenses/GPL-2.0
// Copyright (c) 2021 Mishkin_Ivan@mail.ru
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using System;
using System.Text.RegularExpressions;

namespace Eplan.EplAddin.PlaceholderAction {
    /// <summary>
    /// Данная обработка фильтрует поле "Трасса маршрутизации" в отчете по кабелю для формирования таблицы соединений
    /// </summary>
    class GetCableProperty : IEplAction {
        // Фильтр по умолчанию:
        // выводим в отчет только участки трассы кабеля вида К1, К13, и т.д. Всё остальное не надо.
        const string defaultCableRegex = "K[0-9]*";
        // Свойство проекта для хранения выражения фильтра
        const string regexPath = "EPLAN.Project.UserSupplementaryField1";
        
        private Regex cableRegex = null;
        private string strCableRegex = null;
        public GetCableProperty() {
            cableRegex = new Regex(defaultCableRegex);
            strCableRegex = defaultCableRegex;
        }
        public bool OnRegister(ref string Name, ref int Ordinal) {
            // Записать Name - Eplan не подскажет...
            Name = "GetCablePath";
            Ordinal = 24;
            return true;
        }
        // Поехали!!!
        public bool Execute(ActionCallingContext oActionCallingContext) {
            string objectNames ="";
            oActionCallingContext.GetParameter("objects", ref objectNames);
            // Получим объект текущей строки. Конечно, объектов может быт несколько, но не в этом отчете...
            StorableObject cable = StorableObject.FromStringIdentifier(objectNames);
            if (cable!=null) {
                Project oProject = cable.Project;
                PropertyValue oOnlineProperty = oProject.Properties[regexPath];
                String strProjectCableRegex = oOnlineProperty.ToString(ISOCode.Language.L___);
                if (strProjectCableRegex!=null && strProjectCableRegex!=strCableRegex && !String.IsNullOrWhiteSpace(strProjectCableRegex)) {
                    strCableRegex = strProjectCableRegex;
                    try {
                        Regex cableRegex1 = new Regex(strProjectCableRegex);
                        cableRegex = cableRegex1;
                    }
                    catch (Exception) {
                        ;
                    }
                }
            }
            // Получим свойство 20237 "Топология: Трасса маршрутизации"
            string sCABLING_PATH = cable.Properties[20237];

            MatchCollection matches = cableRegex.Matches(sCABLING_PATH);
            String filteredPath = "";
          
            var enumerator = matches.GetEnumerator();
            if (enumerator.MoveNext()) {
                filteredPath = ((Match)enumerator.Current).Value;
                while (enumerator.MoveNext()) {
                    filteredPath += "; " + (Match)enumerator.Current;
                }
            }

            string[] strings = new string[1];
            strings[0] = filteredPath;
            oActionCallingContext.SetStrings(strings);
            return true;
        }
        // Obsolete - игнорируем.
        public void GetActionProperties(ref ActionProperties actionProperties) {
           return;
        }

    }
}
