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
    /// Для кабеля просматриваем маршрут.
    /// Для участков "Металлорукав" присваиваем наиболее подходящее по диаметру изделие из перечня изделий, присвоенных участку.
    /// Если встретился переход "Лоток-Металлорукав", присваиваем переходник подходящего диаметра
    /// </summary>
    class ActionSetCableParts : IEplAction
    {
        /// <summary>
        /// Диаметр кабеля по умолчанию. Принимается, если для изделия кабеля диаметр не указан
        /// </summary>
        public double DefaulfCableOuterDiameter = 12;
        public enum CABLINGSEGMENT_DUCT_TYPES
        {
            Лоток,
            Металлорукав,
            Труба,
            Траншея,
        }
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "ActionSetCableParts";
            Ordinal = 23;
            return true;
        }
        public bool Execute(ActionCallingContext oActionCallingContext)
        {
            SelectionSet set = new SelectionSet();
            StorableObject[] storableObjects = set.Selection;
            if (storableObjects.Length > 0)
            {
                EplApi.DataModel.Project project = null;
                IEnumerable<Function> functionsQuery = null;
                foreach (Eplan.EplApi.DataModel.EObjects.Cable funcCable in storableObjects.OfType<Eplan.EplApi.DataModel.EObjects.Cable>())
                {
                    var funcIsMain = funcCable.FunctionDefinition.IsMainFunction;
                    //var funcIsCable = func.FunctionDefinition.Category == Function.Enums.Category.Cable;
                    if (funcIsMain /*& funcIsCable*/)
                    {

                        project = project ?? funcCable.Project;
                        functionsQuery = project.Pages.SelectMany(p => p.Functions);
                        PropertyValue pvCablePath = funcCable.Properties.CABLING_PATH;
                        string cablePath = pvCablePath.ToString();
                        if (!String.IsNullOrWhiteSpace(cablePath))
                        {
                            string strCableDiameter = String.Empty;
                            double cableDiameter;
                            Article cableArticle = funcCable.Articles.Length > 0 ? funcCable.Articles[0] : null;
                            strCableDiameter = cableArticle?.Properties.ARTICLE_OUTERDIAMETER;
                            ArticleReference metalHoseArticle = null;
                            bool isKnownCableDiameter = !String.IsNullOrWhiteSpace(strCableDiameter);
                            bool isFoundMetalHoseInPath = false;


                            string[] topologySegments = cablePath.Split(';');
                            double metalHoseLength = 0;
                            foreach (string nameSegment in topologySegments)
                            {
                                var topologySegment = functionsQuery
                                    .OfType<EplApi.DataModel.Topology.Segment>()
                                    .FirstOrDefault(f => f.Properties.FUNC_IDENTNAME == nameSegment);
                                if (null != topologySegment)
                                {
                                    //Нашли очередной сегмент топологии
                                    // var cabblingSegment_DustType = topologySegment.Properties.CABLINGSEGMENT_DUCT_TYPE;

                                    var cabblingSegment_DustType = topologySegment.Properties.CABLINGSEGMENT_DUCT_TYPE;
                                    if (cabblingSegment_DustType.IsEmpty)
                                    {
                                        continue;
                                    }
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
                                        cableDiameter = GetCableDiameter(strCableDiameter);
                                        // Для сегмента с типом "Металлорукав" к кабелю добавляем изделие металлорукава.
                                        if (isMetalHose(dustType))
                                        {
                                            if (!isFoundMetalHoseInPath) // Если это первый найденный сегмент металлорукава
                                            {
                                                // находим первое изделие металлорукава
                                                // и перезаписываем на металлорукав нужного диаметра
                                                metalHoseArticle = GetOrAddMetalHoseArticleRef(funcCable, cableDiameter);
                                                metalHoseArticle.Properties.ARTICLE_PARTIAL_LENGTH_IN_PROJECT_UNIT = 0;
                                                isFoundMetalHoseInPath = true;
                                                metalHoseLength = 0;
                                            }
                                            double segmentLength = topologySegment.Properties.FUNC_CABLING_LENGTH;
                                            metalHoseLength += segmentLength;
                                            // К металлорукаву добавляем длину сегмента
                                            metalHoseArticle.Properties.ARTICLE_PARTIAL_LENGTH_VALUE = metalHoseLength;
                                        }

                                    }

                                }
                            }
                            if (metalHoseArticle != null)
                            {
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
                //if (!func.IsLocked)
                //{
                func.SmartLock();
                ar = func.AddArticleReference(strArticleNr, strVariantNR: "1", nCount: 1, bClean: true);
                //}
                //else
                //{
                //    throw new LockingExceptionFailedLockAttempt($"Для внесения изменений требуется экслюзивный доступ к функции {func.Properties.FUNC_IDENTNAME}", func.ToStringIdentifier());
                //}
            }
            else
            {
                ar.PartNr = GetMetalHoseName(cableDiameter);
            }
            ar.SmartLock();
            ar.StoreToObject();
            return ar;
        }

        private string GetMetalHoseName(double cableDiameter)
        {
            return "МРПИ" + GetMetalHoseDy(cableDiameter);
            int GetMetalHoseDy(double d)
            {
                //МРПИ15
                //            МРПИ20
                //            МРПИ25
                //МРПИ32
                //МРПИ38

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
