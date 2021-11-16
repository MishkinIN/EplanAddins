// GNU General Public License v2.0 https://opensource.org/licenses/GPL-2.0
// Copyright (c) 2021 Mishkin_Ivan@mail.ru
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Eplan.EplAddin.PlaceholderAction {
    class ActionFillPropertiesFromFormat : IEplAction {
        public bool OnRegister(ref string Name, ref int Ordinal) {
            Name = "CopyProperties20202";
            Ordinal = 20;
            return true;
        }
        private int[] TBBNumbers = { 10, 11, 26, 27 };
        public bool Execute(ActionCallingContext oActionCallingContext) {
            SelectionSet set = new SelectionSet();
            foreach (Eplan.EplApi.DataModel.Function item in set.Selection.OfType<Function>()) {
                PropertyValue blockValues = item.Properties.FUNC_BLOCK_VALUE;
                foreach (int blValueNumber in TBBNumbers) {//blockValues.Indexes
                    try {
                        string strValue = blockValues[blValueNumber];
                        item.Properties.FUNC_SUPPLEMENTARYFIELD[blValueNumber] = strValue;
                    }
                    catch (Exception ex) {
                        //throw new ApplicationException(String.Format("Ошибка обработки {0}: {1}", item.Name, ex.Message));
                    }
                }
            }

            return true;
        }

        public void GetActionProperties(ref ActionProperties actionProperties) {
        }
    }
}
