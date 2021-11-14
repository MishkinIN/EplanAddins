// © Mishkin_Ivan@mail.ru 
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;

namespace Eplan.EplAddin.PlaceholderAction {
    /// <summary>
    /// Нумерация страниц тома (томов). Выделить можно любую страницу проекта или несколько проектов сразу.
    /// </summary>
    class ActionNumberPages : IEplAction {
        public bool OnRegister(ref string Name, ref int Ordinal) {
            // Имя для регистрации в системе
            Name = "NimberPagesInProperty11031";
            Ordinal = 21;
            return true; ;
        }
        public bool Execute(ActionCallingContext oActionCallingContext) {
            SelectionSet set = new SelectionSet();
            var projects = set.SelectedProjects;
            if (projects == null || projects.Length == 0)
                return false;
            int countedPages = 0;
            foreach (var project in projects) {
                try {
                    project.LockObject();
                    countedPages += NumberPages(project);
                }
                catch (Exception) {
                }
                finally {
                    project.Dispose();
                }
            }
            return countedPages > 0;
        }

        private static int NumberPages(Eplan.EplApi.DataModel.Project project) {
            int countedPages = 0;
            // Если в проекте несколько томов, каждый нумеруем отдельно.
            Dictionary<string, int> counters = new Dictionary<string, int>();
            var allPages = project.Pages;
            foreach (var page in allPages) {

                try {
                    // В данном скрипте предполагается, что признаком тома в структуре страницы
                    // является <1160> "Определённая пользователем структура" (DESIGNATION_USERDEFINED)
                    string udfKey = page.Properties.DESIGNATION_USERDEFINED;
                    int counter = RestoreCounter(counters, udfKey);
                    MultiLangString pageNumber = new MultiLangString();
                    pageNumber.AddString(ISOCode.Language.L___, (++counter).ToString());
                    page.LockObject();
                    // Запоминаем номер страницы в <11031> Дополнительное поле страницы.
                    page.Properties.PAGE_ADDITIONALPAGE = pageNumber;
                    pageNumber.Dispose();
                    ++countedPages;
                    StoreCounter(counters, udfKey, counter);
                }
                catch (Exception) {

                    throw;
                }
                finally {
                    page.Dispose();
                }
            }
            return countedPages;
        }

        private static void StoreCounter(Dictionary<string, int> counters, string udfKey, int value) {
            counters[udfKey] = value;
        }

        private static int RestoreCounter(Dictionary<string, int> counters, string udfKey) {
            if (counters.ContainsKey(udfKey)) {
                return counters[udfKey];
            }
            else
                return 0;
        }

        public void GetActionProperties(ref ActionProperties actionProperties) {
            ;
        }

    }
}
