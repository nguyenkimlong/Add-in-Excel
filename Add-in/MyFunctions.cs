using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using NetOffice;
using NetOffice.ExcelApi.Enums;
using Excel = NetOffice.ExcelApi;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace Add_in
{
    public class ModelClass
    {
        public string No { get; set; }
        public string ColumnName { get; set; }
        public string DataType { get; set; }
        public string DefaultValue { get; set; }
        public string Nulls { get; set; }
        public string PK { get; set; }
        public string UK { get; set; }
        public string FK { get; set; }
        public string ReferenceTable { get; set; }
        public string ReferenceColumns { get; set; }
        public string Description { get; set; }

    }
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }
    }

    [ComVisible(false)]
    public class MyAddIn : IExcelAddIn
    {

        //public static dynamic _Application = null;

        public static Excel.Application _Application = null;


        public void AutoClose()
        {

        }


        public void AutoOpen()
        {
            try
            {
                if (MyAddIn._Application == null)
                {

                    ExcelDna.Integration.ExcelIntegration.RegisterUnhandledExceptionHandler(ErrorHandler);
                    //_Application = ExcelDnaUtil.Application;
                    _Application = new Excel.Application(null, ExcelDnaUtil.Application);
                    _Application.WorkbookOpenEvent += _Application_WorkbookOpenEvent;

                    _Application.WorkbookActivateEvent += _Application_WorkbookActivateEvent;
                }
            }
            catch (Exception e)
            {

            }
        }

        void _Application_WorkbookActivateEvent(Excel.Workbook Wb)
        {

        }

        void _Application_WorkbookOpenEvent(Excel.Workbook Wb)
        {
            //_Application.Calculation = XlCalculation.xlCalculationAutomatic;
        }

        private object ErrorHandler(object exceptionObject)
        {
            ExcelReference caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

            // Calling reftext here requires all functions to be marked IsMacroType=true, which is undesirable.
            // A better plan would be to build the reference text oneself, using the RowFirst / ColumnFirst info
            // Not sure where to find the SheetName then....
            string callingName = (string)XlCall.Excel(XlCall.xlfReftext, caller, true);


            // return #VALUE into the cell anyway.
            return ExcelError.ExcelErrorValue;
        }



    }

    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        internal static IRibbonUI ribbon = null;
        public override string GetCustomUI(string uiName)
        {
            if (MyAddIn._Application == null)
            {

            }

            return
             @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='LoadImage' onLoad='OnRibbonLoad'>
            <ribbon>
                <tabs>
                <tab id='CustomTab' label='AccNet UX'>
                    <group id='SampleGroup' label='Thiết lập'>
                     <button id='btnluu' label='Cấu hình' image='S' size='large' onAction='OnShowSetting' tag='Config' />
                    </group>
                    <group id='SampleGroup1' label='File'>
                     <button id='btnxuatfile' label='Xuất File .cs' image='M' size='large' onAction='OnPrintClass' tag='File' />
                    </group>

                    <group id='SampleGroup2' label='File Trans'>
                     <button id='btntrans_vn' label='Xuất File Translation_vn' image='M' size='large' onAction='OnPrintTransVn' tag='File' />
                     <button id='btntrans_en' label='Xuất File Translation_en' image='M' size='large' onAction='OnPrintTransEn' tag='File' />
                    </group>    
                       <group id='SampleGroup3' label='File Setting'>
                        <button id='btnsetting' label='Xuất File Setting' image='M' size='large' onAction='OnPrintSetting' tag='File' />
                      </group> 
                </tab>
                </tabs>
            </ribbon>
            </customUI>";

        }

        public void OnRibbonLoad(IRibbonUI objRibbon)
        {
            ribbon = objRibbon;
        }
        public void OnShowSetting(IRibbonControl control)
        {
            ShowManage.ShowCTPSetting();

        }
        public void OnPrintTransVn(IRibbonControl control)
        {
            EventLogFile("OnPrintTransVn");
        }

        public void OnPrintTransEn(IRibbonControl control)
        {
            EventLogFile("OnPrintTransEn");
        }
        public void OnPrintSetting(IRibbonControl control)
        {
            EventLogFile("OnPrintSetting");
        }

        public void EventLogFile(string typeLog)
        {
            //if (Setting.FromCol == null || Setting.ToCol==null) return;
            MyAddIn._Application = new Excel.Application(null, ExcelDnaUtil.Application);

            List<ModelClass> model = new List<ModelClass>();

            Dictionary<string, string> valuedefault = new Dictionary<string, string>();


            dynamic ActiveSheet = MyAddIn._Application.ActiveSheet;

            dynamic nameTitle = MyAddIn._Application.Cells[1, 1].Value;

            object data = MyAddIn._Application.Range(Setting.FromCol + ":" + Setting.ToCol).Value;
            var objLength = ((dynamic)data).Length / 11;
            for (int i = 1; i <= objLength; ++i)
            {
                //add data
                model.Add(new ModelClass
                {
                    No = Convert.ToString(((dynamic)data)[i, 1]),
                    ColumnName = ((dynamic)data)[i, 2],
                    DataType = ((dynamic)data)[i, 3],
                    DefaultValue = Convert.ToString(((dynamic)data)[i, 4]),
                    Nulls = Convert.ToString(((dynamic)data)[i, 5]),
                    PK = Convert.ToString(((dynamic)data)[i, 6]),
                    UK = Convert.ToString(((dynamic)data)[i, 7]),
                    FK = Convert.ToString(((dynamic)data)[i, 8]),
                    ReferenceTable = Convert.ToString(((dynamic)data)[i, 9]),
                    ReferenceColumns = Convert.ToString(((dynamic)data)[i, 10]),
                    Description = Convert.ToString(((dynamic)data)[i, 11]),

                });
                //check default value
                if (Convert.ToString(((dynamic)data)[i, 4]) != null)
                {
                    valuedefault.Add(((dynamic)data)[i, 2], Convert.ToString(((dynamic)data)[i, 4]));
                }
            }

            string name = ActiveSheet.Name;
            if (name.IndexOf("ctg") >= 0)
            {
                name = ActiveSheet.Name.ToString().Substring(name.IndexOf("ctg") + 3);
            }
            switch (typeLog)
            {
                case "OnPrintClass":
                    {
                        LogWriteExport(model, ActiveSheet.Name, valuedefault, nameTitle);
                        break;
                    }
                case "OnPrintTransVn":
                    {
                        WriteTransVN(model, name, valuedefault, nameTitle);
                        break;
                    }
                case "OnPrintTransEn":
                    {
                        WriteTransEN(model, name, valuedefault, nameTitle);
                        break;
                    }
                default:
                    WriteSetting(model, ActiveSheet.Name, valuedefault, nameTitle);
                    break;
            }
        }

        public void OnPrintClass(IRibbonControl control)
        {
            try
            {
                EventLogFile("OnPrintClass");
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        void WriteSetting(List<ModelClass> logdata, string name, Dictionary<string, string> valueDefault, string nameTitle)
        {
            try
            {
                string nameFile = (name.IndexOf("ctg") >= 0) ? Char.ToLowerInvariant(name[name.IndexOf("ctg") + 3]) + name.Substring(name.IndexOf("ctg") + 4) : name;

                var path = "D:\\" + "\\" + nameFile + "Setting.js";
                if (File.Exists(path))
                {
                    File.WriteAllText(path, String.Empty);
                }
                using (StreamWriter w = File.AppendText(path))
                {
                    string nameFirst = "";
                    if (name.IndexOf("ctg") >= 0)
                    {
                        nameFirst = name.ToString().Substring(name.IndexOf("ctg") + 3 );
                        nameFirst = Char.ToLowerInvariant(nameFirst[0]) + nameFirst.Substring(1);
                    }
                    string valueString = "\n";
                    valueString += $"var {nameFirst}Setting = null ; \n";
                    valueString += $"function {nameFirst}InitSetting() \n";
                    valueString += "{\n";
                    valueString += $"{nameFirst}Setting = \n";
                    valueString += "{ \n";
                    valueString += "view : { \n";
                    valueString += "module: \'other\', \n" +
                                 $"formName: '{nameFirst}', \n" +
                                $"gridName: 'grv{nameFirst}', \n" +
                                $"entityName: 'FN_{nameFirst}', \n " +
                                "}, \n";
                    valueString += $"grid{Char.ToUpper(nameFirst[0]) + nameFirst.Substring(1)}: {"{"} \n";
                    valueString += "options: { \n";
                    valueString += "\t height: 474 \n";
                    valueString += "}, \n";
                    valueString += "columns: [ \n";
                    List<string> fieldRequire = new List<string>();

                    foreach (var item in logdata)
                    {
                        string refField = Char.ToLowerInvariant(item.ColumnName[0]) + item.ColumnName.Substring(1, 1).ToLower() + item.ColumnName.Substring(2);

                        string fieldNull = item.Nulls == "x" ? "false" : "true";
                        if (item.FK == "x")
                        {
                            string refTable = Char.ToLowerInvariant(item.ReferenceTable[0]) + item.ReferenceTable.Substring(1);

                            valueString += "{" + $" field: '{ refField}' , name: {nameFirst}Translation.{item.ColumnName.ToUpper()},width: 300, controlType: 'Combobox' ,editable: true,IsRequired : {fieldNull} , ";
                            valueString += "combobox: { \n";
                            valueString += "condition: \"it.Active.ToString() == \\\"Active\\\"\", \n";
                            valueString += $"tableName:\"{item.ReferenceTable}\", \n";
                            valueString += $"baseOn:\"{refTable}Data.{(refField).Substring(0, item.ColumnName.Length - 2) }Name\" , \n";
                            valueString += "dataValueField : \"id\" , \n";
                            valueString += $"dataTextField : \"{(refField).Substring(0, item.ColumnName.Length - 2)}Name\" , \n";
                            valueString += "columns :[ \n";
                            valueString += "{" + $"field: \"{refField}\", width:100 " + "}, \n";
                            valueString += "{" + $"field: \"{(refField).Substring(0, item.ColumnName.Length - 2)}Name\", width:100 " + "} ,\n";
                            valueString += "] \n";
                            valueString += "} \n";


                            valueString += "}, \n";
                        }
                        else if (item.DataType == "enum")
                        {
                            valueString += "{" + $" field: '{refField}' , name: {nameFirst}Translation.{item.ColumnName.ToUpper()},width: 300, controlType: 'CheckBox' ,editable: true,IsRequired : {fieldNull} ";
                            valueString += "}, \n";
                        }
                        else if (item.DataType == "datetime")
                        {
                            valueString += "{" + $" field: '{refField}' , name: {nameFirst}Translation.{item.ColumnName.ToUpper()},width: 150, controlType: 'DateTime' ,editable: true,IsRequired : {fieldNull} ";
                            valueString += "}, \n";
                        }
                        else if (item.DataType == "integer")
                        {
                            valueString += "{" + $" field: '{refField}' , name: {nameFirst}Translation.{item.ColumnName.ToUpper()},width: 200, controlType: 'Numeric' ,editable: true,IsRequired : {fieldNull} ";
                            valueString += "}, \n";
                        }
                        else
                        {
                            valueString += "{" + $" field: '{refField}' , name: {nameFirst}Translation.{item.ColumnName.ToUpper()},width: 300, controlType: 'TextBox' ,editable: true,IsRequired : {fieldNull} ";
                            valueString += "}, \n";
                        }

                        if (item.Nulls != "x")
                        {
                            fieldRequire.Add(refField);
                        }

                    }
                    valueString += "], \n";
                    valueString += "}, \n";
                    valueString += "valuelist: {}, \n";
                    valueString += "options: {}, \n";
                    valueString += "required: [ \n";
                    valueString += "\"" + String.Join("\",\"", fieldRequire) + "\" \n";
                    valueString += "] ,\n";
                    valueString += " validate: { \n";
                    foreach (var item in fieldRequire)
                    {
                        valueString += $"{item} : " + "{" + "required : true }, \n";
                    }
                    valueString += "} \n";
                    valueString += "} \n";
                    valueString += "} \n";
                    w.Write(valueString);
                }

            }
            catch (Exception)
            {

                throw;
            }
        }

        void WriteTransVN(List<ModelClass> logdata, string name, Dictionary<string, string> valueDefault, string nameTitle)
        {
            try
            {
                string nameFile = Char.ToLowerInvariant(name[0]) + name.Substring(1);
                var path = "D:\\" + "\\" + nameFile + "Translation_vn.js";
                List<string> fieldRequire = new List<string>();
                if (File.Exists(path))
                {
                    File.WriteAllText(path, String.Empty);
                }
                using (StreamWriter w = File.AppendText(path))
                {

                    string valueString = "\n";
                    valueString += $"var {nameFile}Translation = " + "{";
                    foreach (var item in logdata)
                    {
                        valueString += "\n";
                        valueString += $"\t {item.ColumnName.ToUpper()} : \"{item.Description}\" ,";
                        valueString += "\n";
                        if (item.Nulls != "x")
                        {
                            valueString += $"\t ERR_{item.ColumnName.ToUpper()}_REQUIRED : \"{item.Description} không được bỏ trống\" ,";
                        }
                    }
                    valueString += $"\n TITLE: \"{nameTitle}\",\n\t UNACTIVE: \"Ngưng hoạt động\" ,\n\t BTN_QUERY: \"Truy vấn\", \n\t BTN_SEARCH: \"Tra cứu\",  \n\t TOTAL: \"Tổng\", \n\t INPUT_SEARCH: \"Tìm kiếm\", \n\t BTN_FILTER: \"Lọc\", \n\t SUCCESS: \"Cập nhật thành công\" ";
                    valueString += "\n }";
                    w.Write(valueString);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        void WriteTransEN(List<ModelClass> logdata, string name, Dictionary<string, string> valueDefault, string nameTitle)
        {
            try
            {
                string nameFile = Char.ToLowerInvariant(name[0]) + name.Substring(1);
                var path = "D:\\" + "\\" + nameFile + "Translation_en.js";
                if (File.Exists(path))
                {
                    File.WriteAllText(path, String.Empty);
                }
                using (StreamWriter w = File.AppendText(path))
                {

                    string valueString = "\n";
                    valueString += $"var {nameFile}Translation = " + "{";
                    foreach (var item in logdata)
                    {
                        valueString += "\n";
                        valueString += $"\t {item.ColumnName.ToUpper()} : \"{item.ColumnName}\" ,";
                        valueString += "\n";
                        if (item.Nulls != "x")
                        {
                            valueString += $"\t ERR_{item.ColumnName.ToUpper()}_REQUIRED : \"ERR {item.ColumnName.ToUpper()} REQUIRED\" ,";
                        }
                    }
                    valueString += "\n TITLE: \"Category\",\n\t UNACTIVE: \"unActive\" ,\n\t BTN_QUERY: \"query\", \n\t BTN_SEARCH: \"search\",  \n\t TOTAL: \"total\", \n\t INPUT_SEARCH: \"search\", \n\t BTN_FILTER: \"FILTER\", \n\t SUCCESS:\"Add Success\" ";
                    valueString += "\n }";
                    w.Write(valueString);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }




        void LogWriteExport(List<ModelClass> logdata, string name, Dictionary<string, string> valueDefault, string nameTitle)
        {
            try
            {
                var path = "D:\\" + "\\" + name + ".cs";
                if (File.Exists(path))
                {
                    File.WriteAllText(path, String.Empty);
                }
                using (StreamWriter w = File.AppendText(path))
                {

                    string valueString = "";
                    valueString += $"using System;";
                    valueString += $"\n";
                    valueString += $"using System.Collections.Generic;";
                    valueString += $"\n";
                    valueString += $"using System.Linq;";
                    valueString += $"\n";
                    valueString += $"using System.Text;";
                    valueString += $"\n";
                    valueString += $"using LV.Core.Common;";
                    valueString += $"\n";
                    valueString += $"using MongoDB.Bson.Serialization.Attributes;";
                    valueString += $"\n";
                    valueString += $"namespace LV.Entities";
                    valueString += "{ \n";
                    valueString += "[BsonIgnoreExtraElements]";
                    valueString += "\n";
                    valueString += $"public class {name} : BaseEntity";
                    valueString += "\n";
                    valueString += "{ \n";
                    valueString += $"//{nameTitle}";
                    valueString += "\n";
                    valueString += $"public {name}() \n";
                    valueString += "{ \n";
                    foreach (var item in valueDefault)
                    {
                        string value = item.Value;
                        value = item.Value == "getDate()" ? "DateTime.Now" : item.Value;
                        valueString += $"this.{item.Key} = {value}; \n";
                    }
                    valueString += "} \n";

                    foreach (var item in logdata)
                    {
                        var datatype = item.DataType;
                        if (item.DataType.IndexOf("varchar") >= 0)
                        {
                            datatype = "string";
                        }
                        if (item.DataType.IndexOf("date") >= 0)
                        {
                            datatype = "DateTime";
                        }
                        if (item.DataType.IndexOf("Decimal") >= 0)
                        {
                            datatype = "decimal";
                        }
                        item.Description = item.Description == null ? item.Description = "\n" : item.Description.Replace("\n", "\n//");
                        valueString += $"//{item.Description} ";
                        if (item.ReferenceTable != null)
                        {
                            valueString += "\n";
                            valueString += $"// \"{name}|{item.ColumnName}|Id\" ";
                            valueString += "\n";
                            //valueString += $"\n public {item.ReferenceTable} {item.ColumnName}" + "{ get; set; } \n";
                            valueString += $"\n public string {item.ColumnName}" + "{ get; set; } \n";
                            valueString += "\n ";
                            valueString += $"[Related(\"{item.ReferenceTable}\", \"{item.ColumnName}\",\"_id\")]";
                            valueString += "\n";
                            valueString += $"\n public {item.ReferenceTable} {item.ReferenceTable}Data" + "{ get; set; } \n";
                        }
                        else
                        {
                            valueString += $"\n public {datatype} {item.ColumnName}" + "{ get; set; } \n";
                        }
                    }
                    valueString += "} \n";
                    valueString += "} \n";
                    w.Write(valueString);
                }
            }
            catch (Exception ex)
            {
            }
        }

    }

}
