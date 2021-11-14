
using System;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Gui;


namespace Eplan.EplAddin.PlaceholderAction {
    /// <summary>
    ///   That is an example for a EPLAN addin.  
    ///   Exactly a class must implement the interface Eplan.EplApi.ApplicationFramework.IEplAddIn.  
    ///   An Assembly is identified through this criterion as EPLAN addin!  
    /// </summary>  
    public class AddInModule : IEplAddIn {
        private uint[] nCommandId = new uint[5];
        private Menu ourMenu;
        
        /// <summary>
        /// The function is called once during registration add-in.
        /// </summary>
        /// <param name="bLoadOnStart"> true: In the next P8 session, add-in will be loaded during initialization</param>
        /// <returns></returns>
        public bool OnRegister(ref System.Boolean bLoadOnStart) {
            bLoadOnStart = true;
            return true;
        }
        /// <summary>
        /// The function is called during unregistration the add-in.
        /// </summary>
        /// <returns></returns>
        public bool OnUnregister() {
            if (ourMenu != null) {
                foreach (var id in nCommandId)
                {
                    ourMenu.RemoveMenuItem(id);
                }
                ourMenu.Dispose();
                ourMenu = null;
            }
            return true;
        }

        /// <summary>
        /// The function is called during P8 initialization or registration the add-in.  
        /// </summary>
        /// <returns></returns>
        public bool OnInit() {

            return true;

        }
        /// <summary>
        /// The function is called during P8 initialization or registration the add-in, when GUI was already initialized and add-in can modify it. 
        /// </summary>
        /// <returns></returns>
        public bool OnInitGui() {
            ourMenu = ourMenu ?? new Menu();
             nCommandId[4] = ourMenu.AddMainMenu("Обработка кабеля", Menu.MainMenuName.eMainMenuUtilities, "Для кабелей присвоить изделие металлорукава", "ActionSetCableParts", "Присвоение металлорукавов для кабелей", 1);
            nCommandId[0] = ourMenu.AddMainMenu("Обработка", Menu.MainMenuName.eMainMenuUtilities, "Заполнить свойства для  формы ТВВ", "CopyProperties20202",
                "Перенос свойств блока в свойства для формы ТВВ", 2);
            nCommandId[1] = ourMenu.AddMenuItem("Пронумеровать страницы томов", "NumberPagesInProperty11031", "Нумерация страниц", nCommandId[0], Int16.MaxValue, false, false);
            nCommandId[2] = ourMenu.AddMenuItem("Пронумеровать доп. поле номера листа", "NumberPagesIn11033", "Нумерация доп. номеров листа", nCommandId[0], Int16.MaxValue, false, false);
             nCommandId[3] = ourMenu.AddMenuItem("Для кабелей присвоить изделие металлорукава", "ActionSetCableParts", "Присвоение металлорукавов для кабелей", nCommandId[0], Int16.MaxValue, false, false);
           return true;

        }
        /// <summary>
        /// This function is called during closing P8 or unregistration the add-in. 
        /// </summary>
        /// <returns></returns>
        public bool OnExit() {
            if (ourMenu != null) {
                foreach (var id in nCommandId)
                {
                    ourMenu.RemoveMenuItem(id);
                }
                ourMenu.Dispose();
                ourMenu = null;
            }
            return true;
        }

    }
}
