// GNU General Public License v2.0 https://opensource.org/licenses/GPL-2.0
// Copyright (c) 2021 Mishkin_Ivan@mail.ru
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Eplan.EplAddin.PlaceholderAction
{
    /// <summary>
    /// Данная обработка выполняет автоматическое присвоение изделия "Металлорукав" для функции кабеля на основании значения свойств "[20237] Маршрут" для кабеля и "[20345] Тип семента маршрутизации" для каждого сегмента, перечисленного в маршруте.
    /// Использование:
    /// 1. Подготавливаем схему соединений, назначаем кабели, кабелю присваиваем изделие кабеля в позиции 1. 
    /// 2. В изделии кабеля должен быть указан наружный диаметр в мм, с указанием единицы измерения через пробел
    /// 3. Размещаем изделия на топологии, подготавливаем трассу маршрутизации.
    /// 4. Для участков маршрута, прокладываемых в металлорукваве, назначаем тип сегмента "Металлорукав" (записываем в свойство [20345])
    /// 5. На чертеже или в навигаторе выбираем функции и запускаем обработку через меню.
    /// Обработка производит перебор выделенных функций. Для функций типа "Кабель" производятся следующие действия:
    /// — Извлекается свойство "[20237] Маршрут". Для каждого сегмента в маршруте анализируется тип сегмента. Если в маршруте присутствует сегмент типа "Металлорукав", для кабеля выполняется следующее:
    /// — Извлекается диаметр кабеля. 
    /// — Находится изделие металлорукава, записанное на функции. Если не найдено, добавляется изделие металлорукава. Для металлорукава:
    /// — устанавливается диаметр металлорукава в зависимости от диаметра кабеля
    /// — устанавливается длина металлорукава как сумма длин сегментов с типом "Металлорукав" в маршруте
    /// </summary>
    class ActionSetCableParts : IEplAction
    {
        public static readonly string ActionName = typeof(ActionSetCableParts).Name;
        /// <summary>
        /// Диаметр кабеля по умолчанию. Принимается, если для изделия кабеля диаметр не указан
        /// </summary>
        public double DefaulfCableOuterDiameter = 12;
        /// <summary>
        /// Известные типы участков маршрута
        /// </summary>
        public enum CABLINGSEGMENT_DUCT_TYPES
        {
            Неизвестен=0,
            Лоток = 1,
            Металлорукав = 2,
            Труба = 3,
            Траншея = 4,
        }
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = ActionName;
            Ordinal = 23;
            return true;
        }
        /// <summary>
        /// Выполнение:
        /// Для кабеля просматриваем маршрут.
        /// Для участков маршрута с типом "Металлорукав" присваиваем наиболее подходящее по диаметру изделие из перечня знакомых изделий металлорукава (Здесь - МРПИ).
        /// </summary>
        public bool Execute(ActionCallingContext oActionCallingContext)
        {

            SelectionSet set = new SelectionSet();
            StorableObject[] storableObjects = set.SelectionRecursive.Length > 0 // выделение выполнено в навигаторе
                ? set.SelectionRecursive // Выбрать, включая вложенные уровни навигатора
                : set.Selection; // либо выделение выполнено на листе

            if (storableObjects.Length > 0)
            {
                EplApi.DataModel.Project project = null;
                IEnumerable<Function> functionsQuery = null;
                foreach (Eplan.EplApi.DataModel.EObjects.Cable funcCable in storableObjects.OfType<Eplan.EplApi.DataModel.EObjects.Cable>())
                {
                    if (funcCable.FunctionDefinition.IsMainFunction)
                    {
                        project = project ?? funcCable.Project;
                        functionsQuery = project.Pages.SelectMany(p => p.Functions);
                        string cablePath = funcCable.Properties.CABLING_PATH.ToString();
                        if (!String.IsNullOrWhiteSpace(cablePath))
                        {
                            string strCableDiameter = String.Empty;
                            double cableDiameter;
                            Article cableArticle = funcCable.Articles.Length > 0 ? funcCable.Articles[0] : null;
                            strCableDiameter = cableArticle?.Properties.ARTICLE_OUTERDIAMETER.ToString();
                            bool isKnownCableDiameter = !String.IsNullOrWhiteSpace(strCableDiameter);
                            cableDiameter = isKnownCableDiameter ? GetCableDiameter(strCableDiameter) : 0.0;
                            ArticleReference metalHoseArticle = null;
                            bool isFoundMetalHoseInPath = false;
                            double metalHoseLength = 0;

                            // Разбиваем строку "Сегменты топологии" на имена сегментов
                            string[] topologySegments = cablePath.Split(';');
                            foreach (string nameSegment in topologySegments)
                            {
                                // По имени сегмента извлекаем функцию
                                EplApi.DataModel.Topology.Segment topologySegment = functionsQuery
                                    .OfType<EplApi.DataModel.Topology.Segment>()
                                    .FirstOrDefault(f => f.Properties.FUNC_IDENTNAME == nameSegment);
                                if (null != topologySegment) //Нашли очередной сегмент топологии
                                {
                                    // Обрабатываем свойство сегмента "[20345] Тип семента маршрутизации"
                                    var cabblingSegment_DustType = topologySegment.Properties.CABLINGSEGMENT_DUCT_TYPE;
                                    if (cabblingSegment_DustType.IsEmpty)
                                    {
                                        continue;
                                    }
                                    // Свойство [20345] - мультиязыковая строка. Значение может быть занесено как интернациональное или ru-Ru
                                    string dustType = cabblingSegment_DustType.ToMultiLangString()
                                        .GetString(EplApi.Base.ISOCode.Language.L___);
                                    if (String.IsNullOrEmpty(dustType))
                                    {
                                        dustType = cabblingSegment_DustType
                                        .ToMultiLangString()
                                        .GetString(EplApi.Base.ISOCode.Language.L_ru_RU);
                                    }
                                    if (isKnownCableDiameter)
                                    {
                                        // Для сегмента с типом "Металлорукав" к кабелю добавляем изделие металлорукава.
                                        if (isMetalHose(dustType))
                                        {
                                            if (!isFoundMetalHoseInPath) // Если это первый найденный сегмент металлорукава
                                            {
                                                // находим первое изделие металлорукава и перезаписываем на металлорукав нужного диаметра
                                                metalHoseArticle = GetOrAddMetalHoseArticleRef(funcCable, cableDiameter);
                                                // Длина будет вычислена заново
                                                metalHoseArticle.Properties.ARTICLE_PARTIAL_LENGTH_IN_PROJECT_UNIT = 0;
                                                isFoundMetalHoseInPath = true;
                                                metalHoseLength = 0;
                                            }
                                            if (!topologySegment.Properties.FUNC_CABLING_LENGTH.IsEmpty)
                                            {
                                                double segmentLength = topologySegment.Properties.FUNC_CABLING_LENGTH;
                                                metalHoseLength += segmentLength;
                                                // К металлорукаву добавляем длину сегмента
                                                metalHoseArticle.Properties.ARTICLE_PARTIAL_LENGTH_VALUE = metalHoseLength;
                                            }
                                        }

                                    }

                                }
                            }
                            if (metalHoseArticle != null)
                            {
                                // 
                                metalHoseArticle.StoreToObject();
                            }
                        }
                    }
                }
            }

            return true;
        }

        private double GetCableDiameter(string strCableDiameter)
        {
            double cableDiameter = DefaulfCableOuterDiameter;
            Double.TryParse(strCableDiameter.Split(' ')[0], out cableDiameter);
            return cableDiameter;
        }

        private bool isMetalHose(string cabblingSegment_DustType)
        {
            return CABLINGSEGMENT_DUCT_TYPES.Металлорукав.ToString() == cabblingSegment_DustType;
        }

        private ArticleReference GetOrAddMetalHoseArticleRef(Function func, double cableDiameter)
        {
            ArticleReference ar = func.ArticleReferences.FirstOrDefault(a =>
            {
                var aPartNumber = a.Properties.ARTICLEREF_PARTNO.ToString();
                return !String.IsNullOrEmpty(aPartNumber) && aPartNumber.StartsWith("МРПИ");
            });
            if (null == ar)
            {
                string strArticleNr = GetMetalHoseName(cableDiameter);
                func.SmartLock();
                ar = func.AddArticleReference(strArticleNr, strVariantNR: "1", nCount: 1, bClean: true);
            }
            else
            {
                ar.PartNr = GetMetalHoseName(cableDiameter);
            }
            ar.SmartLock();
            ar.StoreToObject();
            return ar;
        }
        /// <summary>
        /// Возвращает имя изделия металлорукава как "МРПИnn", где nn = подходящий диаметр металлорукава
        /// </summary>
        /// <param name="cableDiameter">Диаметр кабеля, затягиваемого в металлорукав</param>
        /// <returns></returns>
        private string GetMetalHoseName(double cableDiameter)
        {
                //МРПИ15
                //МРПИ20
                //МРПИ25
                //МРПИ32
                //МРПИ38
            return "МРПИ" + GetMetalHoseDy(cableDiameter);
            int GetMetalHoseDy(double d)
            {

                if (d < 10.0)
                {
                    return 15;
                }
                if (d < 13.0)
                {
                    return 20;
                }
                if (d < 17.0)
                {
                    return 25;
                }
                if (d < 22.0)
                {
                    return 32;
                }
                return 38;
            }
        }
        public void GetActionProperties(ref ActionProperties actionProperties)
        {
        }
    }
}
