using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Plugins;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Clash;
using System.Data;
using System.Data.SqlClient;
using Application = Autodesk.Navisworks.Api.Application;
using System.Reflection;
using System.IO;
using System.Windows;

namespace ExportClashesDB2021
{
    [Plugin("ExportClashes", "IYNO", DisplayName = "Экспорт пересечений")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class ExportClashesDB : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            DocumentClash getOldClashTests = Autodesk.Navisworks.Api.Application.MainDocument.GetClash();
            var filepath = Autodesk.Navisworks.Api.Application.MainDocument.FileName;
            var oldtests = getOldClashTests.TestsData.Tests;
            string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string settings = Path.Combine(assemblyFolder, "ExportClashesDBSettings.txt");
            var connectionString = File.ReadAllLines(settings)[0];
            if (oldtests.Count != 0 && File.Exists(settings))
            {
                foreach (ClashTest test in oldtests)
                {
                    getOldClashTests.TestsData.TestsRunTest(test as ClashTest);
                }
                var tests = Autodesk.Navisworks.Api.Application.MainDocument.GetClash().TestsData.Tests;
                try
                {
                    using (var connection = new System.Data.SqlClient.SqlConnection(connectionString))
                    {

                        if (!IsExistTable("ClashObjects", connection))
                            CreateTableInDataBaseClashObjects("ClashObjects", connection);

                        if (!IsExistTable("ClashTests", connection))
                            CreateTableInDataBaseClashTests("ClashTests", connection);
                        
                        if (!IsExistTable("ClashResults", connection))
                            CreateTableInDataBaseClashResults("ClashResults", connection);
                        
                        //To get data from DB table Elements
                        string query = "select ItemGuid from ClashObjects";
                        connection.Open();
                        SqlCommand cmd = new SqlCommand(query, connection);
                        DataTable sourceData = new DataTable("ClashObjects");
                        sourceData.Columns.Add(new DataColumn("ItemGuid", typeof(string)));
                        using (SqlDataAdapter sqlA = new SqlDataAdapter(cmd))
                        {
                            sqlA.Fill(sourceData);
                        }
                        var listObjectsGuid = new List<string>();
                        // Создаем DataTable для таблицы ClashTests
                        foreach (var test in tests)
                        {
                            DataTable dtClashTests = CreateDataTableClashTest();

                            Guid clashTestID = FillDataTabeClashTest(test, dtClashTests);
                            dtClashTests.DataTableBulkInsert(connection.ConnectionString, "ClashTests");

                            DataTable dtClashResults = CreateDataTableClashResults();

                            var clashResults = (test as ClashTest).Children;
                            foreach (var clashResult in clashResults)
                            {
                                ClashResult rt = clashResult as ClashResult;
                                if (rt.Guid != null && rt.Item1 != null && rt.Item2 != null && rt.Item1.InstanceGuid.ToString() != "00000000-0000-0000-0000-000000000000" && rt.Item2.InstanceGuid.ToString() != "00000000-0000-0000-0000-000000000000")
                                {
                                    clashTestID = FillDataTableClashResults(clashTestID, dtClashResults, rt);
                                    listObjectsGuid.Add(rt.Item1.InstanceGuid.ToString());
                                    listObjectsGuid.Add(rt.Item2.InstanceGuid.ToString());
                                }
                            }
                            dtClashResults.DataTableBulkInsert(connection.ConnectionString, "ClashResults");
                        }

                        var distinctListObjectGuid = listObjectsGuid.Distinct().ToList();
                        var sourcelist = sourceData.AsEnumerable().ToList();
                        var excludedListObjectGuid = distinctListObjectGuid.Where(x => sourcelist.All(y => y.ItemArray[0].ToString() != x.ToString()));

                        DataTable dtClashObjects = CreateDataTableClashObjects();

                        Search search_2 = new Search();
                        search_2.Selection.SelectAll();

                        List<SearchCondition[]> listcondit = new List<SearchCondition[]>();
                        foreach (var itemGuid in excludedListObjectGuid)
                        {
                            listcondit.Add(new SearchCondition[] { (SearchCondition.HasPropertyByName("LcOaNode", "LcOaNodeGuid").DisplayStringWildcard(itemGuid.ToString())) });
                        }
                        foreach (var sccon in listcondit)
                        {
                            search_2.SearchConditions.AddGroup(sccon);
                        }
                        var modelItems = search_2.FindAll(Autodesk.Navisworks.Api.Application.ActiveDocument, false);
                        if (modelItems != null)
                            foreach (var modelItem in modelItems)
                            {
                                FillDataTableClashObjects(dtClashObjects, modelItem);
                            }
                        dtClashObjects.DataTableBulkInsert(connection.ConnectionString, "ClashObjects");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
            else if (oldtests.Count != 0)
            {
                MessageBox.Show("Нет настроек для нахождения пересечений.");
                return 0;
            }
            else
            {
                MessageBox.Show("Файл с настройками подключения к базе данных не найден.");
                return 0;
            }
            return 0;
        }

        private void CreateTableInDataBaseClashResults(string tableName, SqlConnection conn)
        {
            conn.Open();
            SqlCommand cmd = new SqlCommand("create table " + tableName + " (Name nvarchar(256), ClashGUID nvarchar(80), X nvarchar(100), Y nvarchar(100), Z nvarchar(100), ClashGridIntersection nvarchar(256), ClashLevel nvarchar(256), Item1Guid nvarchar(256), Item2Guid nvarchar(256), ClashResultStatus nvarchar(256), ClashDistance nvarchar(256), ClashTestID nvarchar(256));", conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        private void CreateTableInDataBaseClashTests(string tableName, SqlConnection conn)
        {
            conn.Open();
            SqlCommand cmd = new SqlCommand("create table " + tableName + " (Name nvarchar(256), Type nvarchar(256), Status nvarchar(40), [Tolerance] nvarchar(100), [MergeComposite] nvarchar(100), [SelectionSetLeft] nvarchar(256), [SelectionSetRight] nvarchar(256), [CreateDate] nvarchar(50), [ClashTestID] nvarchar(80));", conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        private void CreateTableInDataBaseClashObjects(string tableName, SqlConnection conn)
        {
            conn.Open();
            SqlCommand cmd = new SqlCommand("create table " + tableName + " (ItemGuid nvarchar(80), ID nvarchar(50), WorksetName nvarchar(256), Category nvarchar(80), SourceFile nvarchar(256), FamilyName nvarchar(256), TypeName nvarchar(256));", conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        private bool IsExistTable(string tableName, SqlConnection conn)
        {
            bool exists;

            try
            { 
                var cmd = new SqlCommand(
                  "select case when exists((select * from information_schema.tables where table_name = '" + tableName + "')) then 1 else 0 end", conn);

                exists = (int)cmd.ExecuteScalar() == 1;
            }
            catch
            {
                try
                {
                    exists = true;
                    var cmdOthers = new SqlCommand("select 1 from " + tableName + " where 1 = 0");
                    cmdOthers.ExecuteNonQuery();
                }
                catch
                {
                    exists = false;
                }
            }
            return exists;
        }

        private static void FillDataTableClashObjects(DataTable dtClashObjects, ModelItem modelItem)
        {
            var newrowClashObject = dtClashObjects.NewRow();
            newrowClashObject["ItemGuid"] = modelItem.InstanceGuid;
            var elementId = modelItem.PropertyCategories.FindPropertyByName("LcRevitData_Element", "LcRevitPropertyElementId")?.Value;
            newrowClashObject["ID"] = elementId?.DataType == VariantDataType.Int32 ? elementId.ToInt32().ToString() : elementId.ToDisplayString();
            newrowClashObject["WorksetName"] = modelItem.PropertyCategories?.FindPropertyByDisplayName("Объект", "Рабочий набор")?.Value?.ToDisplayString();
            newrowClashObject["Category"] = modelItem.PropertyCategories?.FindPropertyByName("LcRevitData_Element", "LcRevitPropertyElementCategory")?.Value?.ToDisplayString();
            newrowClashObject["SourceFile"] = modelItem.PropertyCategories?.FindPropertyByName("LcOaNode", "LcOaNodeSourceFile")?.Value?.ToDisplayString();
            newrowClashObject["FamilyName"] = modelItem.PropertyCategories?.FindPropertyByName("LcRevitData_Element", "LcRevitPropertyElementFamily")?.Value?.ToDisplayString();
            newrowClashObject["TypeName"] = modelItem?.PropertyCategories?.FindPropertyByName("LcRevitData_Type", "LcRevitPropertyElementName")?.Value?.ToDisplayString();
            dtClashObjects.Rows.Add(newrowClashObject);
        }

        private static DataTable CreateDataTableClashObjects()
        {
            DataTable dtClashObjects = new DataTable("ClashObjects");
            dtClashObjects.Columns.Add(new DataColumn("ItemGuid", typeof(string)));
            dtClashObjects.Columns.Add(new DataColumn("ID", typeof(string)));
            dtClashObjects.Columns.Add(new DataColumn("WorksetName", typeof(string)));
            dtClashObjects.Columns.Add(new DataColumn("Category", typeof(string)));
            dtClashObjects.Columns.Add(new DataColumn("SourceFile", typeof(string)));
            dtClashObjects.Columns.Add(new DataColumn("FamilyName", typeof(string)));
            dtClashObjects.Columns.Add(new DataColumn("TypeName", typeof(string)));
            return dtClashObjects;
        }

        private Guid FillDataTableClashResults(Guid clashTestID, DataTable dtClashResults, ClashResult rt)
        {
            var clashpoint = rt.Center;
            var listIntersect = GetGridIntersection(clashpoint);
            var newrowClashResults = dtClashResults.NewRow();
            newrowClashResults["Name"] = rt.DisplayName;
            newrowClashResults["ClashGUID"] = rt.Guid.ToString();
            newrowClashResults["X"] = clashpoint.X.ToString();
            newrowClashResults["Y"] = clashpoint.Y.ToString();
            newrowClashResults["Z"] = clashpoint.Z.ToString();
            newrowClashResults["ClashGridIntersection"] = listIntersect.FirstOrDefault();
            newrowClashResults["ClashLevel"] = listIntersect.LastOrDefault();
            newrowClashResults["Item1Guid"] = rt.Item1.InstanceGuid.ToString();
            newrowClashResults["Item2Guid"] = rt.Item2.InstanceGuid.ToString();
            newrowClashResults["ClashResultStatus"] = rt.Status.ToString();
            newrowClashResults["ClashDistance"] = rt.Distance.ToString();
            newrowClashResults["ClashTestID"] = clashTestID.ToString();
            dtClashResults.Rows.Add(newrowClashResults);
            return clashTestID;
        }

        private static DataTable CreateDataTableClashResults()
        {
            DataTable dtClashResults = new DataTable("ClashResults");
            dtClashResults.Columns.Add(new DataColumn("Name", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("ClashGUID", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("X", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("Y", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("Z", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("ClashGridIntersection", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("ClashLevel", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("Item1Guid", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("Item2Guid", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("ClashResultStatus", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("ClashDistance", typeof(string)));
            dtClashResults.Columns.Add(new DataColumn("ClashTestID", typeof(string)));
            return dtClashResults;
        }

        private static Guid FillDataTabeClashTest(SavedItem test, DataTable dtClashTests)
        {
            var newrow = dtClashTests.NewRow();
            newrow["Name"] = (test as ClashTest).DisplayName.ToString();
            newrow["Type"] = (test as ClashTest).TestType.ToString();
            newrow["Status"] = (test as ClashTest).Status.ToString();
            newrow["Tolerance"] = (test as ClashTest).Tolerance.ToString();
            newrow["MergeComposite"] = (test as ClashTest).MergeComposites.ToString();
            newrow["SelectionSetLeft"] = (test as ClashTest).SelectionA.SelfIntersect.ToString();
            newrow["SelectionSetRight"] = (test as ClashTest).SelectionB.SelfIntersect.ToString();
            newrow["CreateDate"] = DateTime.Now.ToString();
            var clashTestID = Guid.NewGuid();
            newrow["ClashTestID"] = clashTestID.ToString();
            dtClashTests.Rows.Add(newrow);
            return clashTestID;
        }

        private static DataTable CreateDataTableClashTest()
        {
            DataTable dtClashTests = new DataTable("ClashTests");
            dtClashTests.Columns.Add(new DataColumn("Name", typeof(string)));
            dtClashTests.Columns.Add(new DataColumn("Type", typeof(string)));
            dtClashTests.Columns.Add(new DataColumn("Status", typeof(string)));
            dtClashTests.Columns.Add(new DataColumn("Tolerance", typeof(string)));
            dtClashTests.Columns.Add(new DataColumn("MergeComposite", typeof(string)));
            dtClashTests.Columns.Add(new DataColumn("SelectionSetLeft", typeof(string)));
            dtClashTests.Columns.Add(new DataColumn("SelectionSetRight", typeof(string)));
            dtClashTests.Columns.Add(new DataColumn("CreateDate", typeof(string)));
            dtClashTests.Columns.Add(new DataColumn("ClashTestID", typeof(string)));
            return dtClashTests;
        }

        public List<string> GetGridIntersection(Point3D point)
        {
            GridIntersection oGridIntersection = null;
            string nameIntesect;
            string level;
            var listIntersection = new List<String>();
            GridSystem oGS = Autodesk.Navisworks.Api.Application.ActiveDocument.Grids.ActiveSystem;
            if (oGS != null)
                oGridIntersection = oGS.ClosestIntersection(point);

            if (oGridIntersection != null)
            {
                nameIntesect = oGridIntersection.DisplayName;
                level = oGridIntersection.Level.DisplayName;
                listIntersection.Add(nameIntesect);
                listIntersection.Add(level);
            }
            return listIntersection;
        }

        public List<ClashResult> RecurseChildren(GroupItem group)
        {
            var testClashResult = new List<ClashResult>();
            foreach (SavedItem child in group.Children)
            {
                GroupItem child_group = child as GroupItem;
                if (child_group != null)
                {
                    RecurseChildren(child_group);
                }
                else
                {
                    ClashResult result = child as ClashResult;
                    testClashResult.Add(result);
                }
            }
            return testClashResult;
        }

        public int RandomNumber(int min, int max)
        {
            Random random = new Random();
            return random.Next(min, max);
        }
    }
}
