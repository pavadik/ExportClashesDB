using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
using exportdb.ViewForm;

namespace ExportClashesDB
{
    [Plugin("ExportClashes", "GRVN", DisplayName = "Экспорт пересечений")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class ExportClashesDB : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            var projectIdFform = new AddProjectIdForm();
            projectIdFform.ShowDialog();
            if (projectIdFform.DialogResult.Equals(DialogResult.OK) & projectIdFform.GetProjectId != " " & projectIdFform.GetProjectId != null)
            {
                DocumentClash getOldClashTests = Autodesk.Navisworks.Api.Application.MainDocument.GetClash();
                var filepath = Autodesk.Navisworks.Api.Application.MainDocument.FileName;
                var oldtests = getOldClashTests.TestsData.Tests;
                if (oldtests.Count != 0)
                {
                    foreach (ClashTest test in oldtests)
                    {
                        getOldClashTests.TestsData.TestsRunTest(test as ClashTest);
                    }
                    var tests = Autodesk.Navisworks.Api.Application.MainDocument.GetClash().TestsData.Tests;
                    using (var connection = new System.Data.SqlClient.SqlConnection(Properties.connectionDb.Default.DbConnector))
                    {
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
                            dtClashTests.Columns.Add(new DataColumn("ProjectId", typeof(string)));

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
                            newrow["ProjectId"] = projectIdFform.GetProjectId;
                            dtClashTests.Rows.Add(newrow);
                            DataTableBulkInsert(dtClashTests, connection, "ClashTests");

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

                            var clashResults = (test as ClashTest).Children;
                            foreach (var clashResult in clashResults)
                            {
                                ClashResult rt = clashResult as ClashResult;
                                if (rt.Guid != null && rt.Item1 != null && rt.Item2 != null && rt.Item1.InstanceGuid.ToString() != "00000000-0000-0000-0000-000000000000" && rt.Item2.InstanceGuid.ToString() != "00000000-0000-0000-0000-000000000000")
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
                                    listObjectsGuid.Add(rt.Item1.InstanceGuid.ToString());
                                    listObjectsGuid.Add(rt.Item2.InstanceGuid.ToString());
                                    dtClashResults.Rows.Add(newrowClashResults);
                                }
                            }
                            DataTableBulkInsert(dtClashResults, connection, "ClashResults");
                        }
                        var table = new DataTable();
                        table.Columns.Add("ItemGuid", typeof(string));
                        var distinctListObjectGuid = listObjectsGuid.Distinct().ToList();
                        var sourcelist = sourceData.AsEnumerable().ToList();
                        var excludedListObjectGuid = distinctListObjectGuid.Where(x => sourcelist.All(y => y.ItemArray[0].ToString() != x.ToString()));

                        DataTable dtClashObjects = new DataTable("ClashObjects");
                        dtClashObjects.Columns.Add(new DataColumn("ItemGuid", typeof(string)));
                        dtClashObjects.Columns.Add(new DataColumn("ID", typeof(string)));
                        dtClashObjects.Columns.Add(new DataColumn("WorksetName", typeof(string)));
                        dtClashObjects.Columns.Add(new DataColumn("Category", typeof(string)));
                        dtClashObjects.Columns.Add(new DataColumn("SourceFile", typeof(string)));
                        dtClashObjects.Columns.Add(new DataColumn("FamilyName", typeof(string)));
                        dtClashObjects.Columns.Add(new DataColumn("TypeName", typeof(string)));

                        Search search_2 = new Search();
                        search_2.Selection.SelectAll();
                        var oG1 = new System.Collections.Generic.List<SearchCondition>();
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
                                var newrowClashObject = dtClashObjects.NewRow();
                                newrowClashObject["ItemGuid"] = modelItem.InstanceGuid;
                                var elementId = modelItem.PropertyCategories.FindPropertyByName("LcRevitData_Element", "LcRevitPropertyElementId")?.Value;
                                newrowClashObject["ID"] = elementId.DataType == VariantDataType.Int32 ? elementId.ToInt32().ToString() : elementId.ToDisplayString();
                                newrowClashObject["WorksetName"] = modelItem.PropertyCategories.FindPropertyByDisplayName("Объект", "Рабочий набор")?.Value?.ToDisplayString();
                                newrowClashObject["Category"] = modelItem.PropertyCategories.FindPropertyByName("LcRevitData_Element", "LcRevitPropertyElementCategory")?.Value?.ToDisplayString();
                                newrowClashObject["SourceFile"] = modelItem.PropertyCategories.FindPropertyByName("LcOaNode", "LcOaNodeSourceFile")?.Value?.ToDisplayString();
                                newrowClashObject["FamilyName"] = modelItem.PropertyCategories.FindPropertyByName("LcRevitData_Element", "LcRevitPropertyElementFamily")?.Value?.ToDisplayString();
                                newrowClashObject["TypeName"] = modelItem?.PropertyCategories.FindPropertyByName("LcRevitData_Type", "LcRevitPropertyElementName")?.Value?.ToDisplayString();
                                dtClashObjects.Rows.Add(newrowClashObject);
                            }
                        DataTableBulkInsert(dtClashObjects, connection, "ClashObjects");
                    }
                }
                return 0;
            }
            else
            {
                return 0;
            }
        }
        public static void DataTableBulkInsert(DataTable Table, SqlConnection connection, string destinationDataTable)
        {
            SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connection);
            sqlBulkCopy.DestinationTableName = destinationDataTable;
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            sqlBulkCopy.WriteToServer(Table);
            connection.Close();
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
