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
            var address = ArgToAddress(arguments, 1, context);
            var pt = context.ExcelDataProvider.GetPivotTableFromAddress(context.Scopes.Current.Address.Worksheet, address);
            object v;
            if(pt==null)
            {
                v = ExcelErrorValue.Create(eErrorType.Ref);            
            }
            else
            {
                v = GetData(pt, arguments);
            }
            var crf = new CompileResultFactory();
            return crf.Create(v);
        }

        private object GetData(ExcelPivotTable pt, IEnumerable<FunctionArgument> arguments)
        {
            var l = arguments.ToList();
            if (l.Count % 2 == 1) ExcelErrorValue.Create(eErrorType.Ref);

            var dataField = pt.Fields[l[0].ToString()];
            if(dataField==null || dataField.IsDataField==false)
            {
                return ExcelErrorValue.Create(eErrorType.Ref);
            }
            var fields = new List<string[]>();
            for (int i=2;i<l.Count;i+=2)
            {                
                var fieldName = l[i].ToString();
                var item = l[i + 1].ToString();

                var f = pt.Fields[fieldName];
                if(f.IsColumnField==true || f.IsRowField==true)
                {
                    fields.Add(new string[] { fieldName, item });
                }
                else
                {
                    return ExcelErrorValue.Create(eErrorType.Ref);
                }                
            }

            return CalculateData(pt, fields);
        }

        private object CalculateData(ExcelPivotTable pt, List<string[]> fields)
        {
            return 0;
        }
    }
}
