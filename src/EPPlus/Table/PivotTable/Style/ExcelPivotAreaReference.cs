﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotAreaReference
    {
        public int Field { get; set; }
        public bool Selected { get; set; }
        public List<object> Values { get; } = new List<object>();
    }
}