using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "5.5",
        Description = "Returns data from a pivot table")]
    internal class GetPivotData : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            int ix = 0;
            ExcelPivotTable pt=null;
            string dataFieldName = "";
            var paramList= new List<string>();
            foreach (var arg in arguments)
            {
                if(ix==0)
                {
                    dataFieldName = arg.Value.ToString();
                }
                else if (ix==1)
                {
                    var address = ArgToAddress(arguments, 1, context);
                    pt = context.ExcelDataProvider.GetPivotTableFromAddress(context.Scopes.Current.Address.Worksheet, address);
                }
                else
                {
                    paramList.Add(arg.Value.ToString());
                }
            }

            object v;
            if (pt == null || string.IsNullOrEmpty(dataFieldName))
            {
                v=ExcelErrorValue.Create(eErrorType.Ref);
            }
            else
            {
                v = pt.GetPivotData(dataFieldName, paramList.ToArray());
            }

            var crf = new CompileResultFactory();
            return crf.Create(v);
        }


        private object CalculateData(ExcelPivotTable pt, List<string[]> fields)
        {
            return 0;
        }
    }
}
