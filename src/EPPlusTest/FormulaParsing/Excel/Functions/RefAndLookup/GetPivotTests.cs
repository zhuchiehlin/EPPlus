/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  10/15/2020         EPPlus Software AB       EPPlus 5.5
 *******************************************************************************/
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class GetPivotDataTests : TestBase
    {
        ExcelPackage _pck;
        [TestInitialize]
        public void Initialize()
        {
            _pck = OpenPackage("GetPivotData.xlsx", true);
            var ws = _pck.Workbook.Worksheets.Add("Data");
            LoadItemData(ws);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _pck.Dispose();
        }


        [TestMethod]
        public void GetPivotSholdReturn()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Pivot-WithRecord");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable1");
            //pt.ColumnFields.Add(pt.Fields[1]);
            var rf = pt.RowFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = true;

            ws.Cells["J1"].Formula = "GetPivotData(\"Stock\",A1,\"Category\",\"Groceries\")";
            ws.Calculate();
        }
    }
}

