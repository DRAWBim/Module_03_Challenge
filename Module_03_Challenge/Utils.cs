using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Parameter = Autodesk.Revit.DB.Parameter;

namespace Module_03_Challenge
{
    internal static class Utils
    {
        internal static string GetParameterValueAsString(Element elem, string paramName)
        {
            IList<Parameter> paramList = elem.GetParameters(paramName);
            Parameter myParam = paramList.First();

            return myParam.AsString();
        }

        // Set parameter value
        internal static void SetParameterValueAsString(Element elem, string paramName, string paramValue)
        {
            IList<Parameter> paramList = elem.GetParameters(paramName);
            Parameter Param = paramList.First();

            Param.Set(paramValue);
        }
        internal static void SetParameterValueAsDouble(Element elem, string paramName, double paramValue)
        {
            IList<Parameter> paramList = elem.GetParameters(paramName);
            Parameter Param = paramList.First();

            Param.Set(paramValue);
        }

        // Get family symbol
        internal static FamilySymbol GetFamilySymbolByName(Document doc, string famName, string fsName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));
            foreach (FamilySymbol familySymbol in collector)
            {
                if (familySymbol.Name == fsName && familySymbol.FamilyName == famName)
                    return familySymbol;
            }

            return null;
        }
    }
}
