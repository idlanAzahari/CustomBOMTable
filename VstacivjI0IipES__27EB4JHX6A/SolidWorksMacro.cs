using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using Excel = Microsoft.Office.Interop.Excel;



namespace Macro1
{
    public partial class SolidWorksMacro
    {
        
        
        public SldWorks swApp;
        public void Main()
        {
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            Component2 swComp = default(Component2);
            ConfigurationManager swConfMgr = default(ConfigurationManager);
            Configuration swConf = default(Configuration);
            bool checkBomPath = checkBOMTemplate();

            if (checkBomPath)
            {
                int nErrors = 0;
                int nWarnings = 0;
                swModel = (ModelDoc2)swApp.ActiveDoc;
                swModel = (ModelDoc2)swApp.OpenDoc6(@"D:\idlan.azahari\Documents\BOM_table_task\2023-07-14-REV0 - SRU-ASR-001 - BOGGIE - TO IME\A - SRU-ASR-001 - BOGGIE - ALL - TO IME - V1.SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref nErrors, ref nWarnings);
                swModelDocExt = (ModelDocExtension)swModel.Extension;
                swConfMgr = (ConfigurationManager)swModel.ConfigurationManager;
                swConf = (Configuration)swConfMgr.ActiveConfiguration;
                swComp = (Component2)swConf.GetRootComponent();

                //Ni function nak test
                //List<string> firstLayerNames = getFirstLayerAssemblyNames(swModel);
                //createBomTable(swModel);
                List<string> partNumber = getCustomPropertyPartNum1(swModel);
                //saveToExcel(swModel);

                //create firstlayer column
                //createColumn(swModel);
            }
            else
                MessageBox.Show("Please make sure BOM template is in Solidworks template folder");
            
        }

        //ni settle
        public void createBomTable(ModelDoc2 swModel)
        {
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            BomTableAnnotation swBOMAnnotation = default(BomTableAnnotation);

            swModelDocExt = (ModelDocExtension)swModel.Extension;

            String TemplateName = @"C:\Program Files\SOLIDWORKS Corp 2023\SOLIDWORKS\lang\english\BOM_Excel_template.sldbomtbt";
            int BomType = (int)swBomType_e.swBomType_PartsOnly;
            string configuration = "Default";
            int nbrType = (int)swNumberingType_e.swNumberingType_Detailed;

            swBOMAnnotation = (BomTableAnnotation)swModelDocExt.InsertBomTable3(TemplateName, 0, 0, BomType, configuration, false, nbrType, true);
            swModel.ForceRebuild3(false);
            swModel.ViewZoomtofit2();
        }

        //ni settle
        public void saveToExcel(ModelDoc2 swModel)
        {
            ModelDocExtension swModelExt = default(ModelDocExtension);
            SelectionMgr swSM = default(SelectionMgr);
            TableAnnotation swTable = default(TableAnnotation);

            swModelExt = (ModelDocExtension)swModel.Extension;
            swSM = (SelectionMgr)swModel.SelectionManager;
            swModelExt.SelectByID2("DetailItem1@Annotations", "ANNOTATIONTABLES", 0, 0, 0, false, 0, null, 0);

            swTable = (TableAnnotation)swSM.GetSelectedObject6(1, 0);

            BomTableAnnotation swBomTable = default(BomTableAnnotation);
            swBomTable = (BomTableAnnotation)swTable;

            swBomTable.SaveAsExcel(@"D:\idlan.azahari\Documents\BOM_table_task\BomTableTest3.xls", false, false);



        }

        //ni settle
        public bool checkBOMTemplate()
        {
            string BOMPath = @"C:\Program Files\SOLIDWORKS Corp 2023\SOLIDWORKS\lang\english\BOM_Excel_template.sldbomtbt";
            string BOMPath1 = @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\BOM_Excel_template.sldbomtbt";

            if (File.Exists(BOMPath) || File.Exists(BOMPath1))
                return true;
            else
                return false;
        }

        //public List<string> getPartNumberList()
        //{
        //    List<string> partNumber = new List<string>();

        //}

        //get first layer names
        public List<string> getFirstLayerAssemblyNames(ModelDoc2 swModel)
        {
            var childPartNum = new List<string>();
            object[] childComp;
            Component2 swChildComp = default(Component2);
            Component2 swComp = default(Component2);
            Configuration swConf = default(Configuration);

            swConf = (Configuration)swModel.GetActiveConfiguration();
            swComp = (Component2)swConf.GetRootComponent3(true);

            childComp = (object[])swComp.GetChildren();

            for (int i = 0; i < childComp.Length; i++)
            {
                swChildComp = (Component2)childComp[i];
                Debug.Print("Component name: " + swChildComp.Name2);
                string partNum = getProperty(swModel, swComp);
                childPartNum.Add(swChildComp.Name2);
            }

            return childPartNum;
        }

        //
        public List<string> getCustomPropertyPartNum(ModelDoc2 swModel)
        {
            Component2 swComp = default(Component2);
            AssemblyDoc swAssy = default(AssemblyDoc);
            Configuration swConf = default(Configuration);

            List<string> partNumList = new List<string>();
            swAssy = (AssemblyDoc)swModel;
            object[] vComps = (Object[])swAssy.GetComponents(true);

            for (int i = 0; i< vComps.Length; i++)
            {
                swComp = (Component2)vComps[i];
                swModel = (ModelDoc2)swComp.GetModelDoc2();
                swConf = (Configuration)swModel.GetActiveConfiguration();
                swComp = (Component2)swConf.GetRootComponent3(true);

                string PartNum = getProperty(swModel, swComp);
                partNumList.Add(PartNum);
            }
            return partNumList;
        }

        public string getProperty(ModelDoc2 swModel, Component2 swComp)
        {
            CustomPropertyManager cpm = default(CustomPropertyManager);
            cpm = (CustomPropertyManager)swModel.Extension.get_CustomPropertyManager("");
            string val;
            string valout;

            cpm.Get4("swPartNum", false, out val, out valout);

            Debug.Print("Value: " + val);
            Debug.Print("||||||||||");
            return val;
        }
        //belum agi
        public List<string> getCustomPropertyPartNum1(ModelDoc2 swModel)
        {
            Component2 swComp = default(Component2);
            AssemblyDoc swAssy = default(AssemblyDoc);
            Configuration swConf = default(Configuration);

            List<string> partNumList = new List<string>();
            swAssy = (AssemblyDoc)swModel;
            object[] vComps = (Object[])swAssy.GetComponents(true);

            for (int i = 0; i < vComps.Length; i++)
            {
                swComp = (Component2)vComps[i];
                swModel = (ModelDoc2)swComp.GetModelDoc2();
                swConf = (Configuration)swModel.GetActiveConfiguration();
                swComp = (Component2)swConf.GetRootComponent3(true);

                string PartNum = getProperty(swModel, swComp);
                partNumList.Add(PartNum);
                
            }
            return partNumList;
        }




        //Create firstLayer columns
        public void createColumn(ModelDoc2 swModel)
        {
            string pathName = @"D:\idlan.azahari\Documents\BOM_table_task\BomTableTest3.xls";
            List<string> firstLayer = getFirstLayerAssemblyNames(swModel);


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(pathName);

            Excel.Worksheet ws = wb.Sheets[1];

            int currentRow = 1;
            int columnNum = 7;

            foreach (var val in firstLayer)
            {
                ws.Cells[currentRow, columnNum].Value = val;
                columnNum++;
            }


            wb.Save();


        }

        public List<object> partNumFromExcel()
        {
            List<object> partNum = new List<object>();
            string filepath = @"D:\idlan.azahari\Documents\BOM_table_task\BomTableTest3.xls";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(filepath);
            Excel.Worksheet ws = wb.Sheets[1];
            ;

            Excel.Range usedRange = ws.UsedRange;

            int columnRetrieve = 2;

            for (int rowIndex = 2; rowIndex <= usedRange.Rows.Count; rowIndex++)
            {
                Excel.Range cell = (Excel.Range)usedRange.Cells[rowIndex, columnRetrieve];
                partNum.Add(cell.Value2);
            }

            return partNum;

        }

        //ni untuk comparison
        public List<object> partNumFromExcel_new()
        {
            List<object> partNum = partNumFromExcel();

            for (int i = 0; i < partNum.Count; i++)
            {
                if (partNum[i] != null)
                {
                    string val1 = partNum[i].ToString();
                    val1 = val1.Replace("\n", ""); // Correct the newline character
                    partNum[i] = val1;
                }
            }

            return partNum;

        }

        

        public void createRowsForSubAssemCol()
        {
            //ni list after susun
            List<int> val = new List<int>()
            {
                1,2,3,4,5,6
            };

            string filepath = @"D:\idlan.azahari\Documents\BOM_table_task\BomTableTest3.xls";

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(filepath);
            Excel.Worksheet ws = wb.Sheets[1];


            int row = 2; //start at row 2

            foreach (int vals in val)
            {
                ws.Cells[row, "g"].Value = vals;
                row++;
            }

            wb.Save();

        }
        //public void printGroup(ModelDoc2 swModel)
        //{

        //    List<string> a = getCustomPropertyPartNum(swModel);
        //    List<string> b = new List<int> { 3, 6, 10 };

        //    List<List<string>> c = SplitListByValues(a, b);

        //    Console.WriteLine("Output List:");
        //    foreach (List<string> sublist in c)
        //    {
        //        Console.WriteLine(string.Join(",", sublist));

        //    }
        //}

        public  List<List<int>> SplitListByValues(List<int> a, List<int> b)
        {
            List<List<int>> result = new List<List<int>>();
            List<int> sublist = new List<int>();

            foreach (int num in a)
            {
                sublist.Add(num);
                if (b.Contains(num))
                {
                    result.Add(sublist);
                    sublist = new List<int>();
                }
            }

            if (sublist.Any())
            {
                result.Add(sublist);
            }

            return result;
        }



        //public string compareItems(List<string> excelColumn, List<string> childComp)
        //{

        //}
    }
        
}

