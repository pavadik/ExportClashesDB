using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.ComponentModel;
using System.Reflection;
using System.Windows;
using System.Data.SqlClient;

namespace ExportClashesDB
{
    public static class Extentions
    {
        public static IEnumerable<TSource> Exclude<TSource, TKey>(this IEnumerable<TSource> source,
                          IEnumerable<TSource> exclude, Func<TSource, TKey> keySelector)
        {
            var excludedSet = new HashSet<TKey>(exclude.Select(keySelector));
            return source.Where(item => !excludedSet.Contains(keySelector(item)));
        }
        public static DataTable ConvertListToDataTable<T>(this List<T> iList)
        {
            DataTable dataTable = new DataTable();
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor propertyDescriptor = props[i];
                Type type = propertyDescriptor.PropertyType;

                if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    type = Nullable.GetUnderlyingType(type);

                dataTable.Columns.Add(propertyDescriptor.Name, type);
            }
            object[] values = new object[props.Count];
            foreach (T iListItem in iList)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(iListItem);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }
        public static DataTable ConvertToDatatable<T>(this IList<T> list)
        {
            DataTable t = new DataTable();
            Type elementType = typeof(T);
            //add a column to table for each public property on T
            foreach (var propInfo in elementType.GetProperties())
            {
                Type ColType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
                t.Columns.Add(propInfo.Name, ColType);
            }
            //go through each property on T and add each value to the table
            foreach (T item in list)
            {
                DataRow row = t.NewRow();
                foreach (var propInfo in elementType.GetProperties())
                {
                    try
                    {
                        row[propInfo.Name] = propInfo.GetValue(item, null) ?? DBNull.Value;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                t.Rows.Add(row);
            }
            return t;
        }
        public static void AddToDataTable<T>(this IEnumerable<T> enumerable, System.Data.DataTable table)
        {
            if (enumerable.FirstOrDefault() == null)
            {
                table.Rows.Add(new[] { string.Empty });
                return;
            }

            var properties = enumerable.FirstOrDefault().GetType().GetProperties();

            foreach (var item in enumerable)
            {
                var row = table.NewRow();
                foreach (var property in properties)
                {
                    row[property.Name] = item.GetType().InvokeMember(property.Name, BindingFlags.GetProperty, null, item, null);
                }
                table.Rows.Add(row);
            }
        }
        public static DataTable ConvertToDataTable<T>(this IList<T> list)
        {
            DataTable table = CreateTable<T>();
            var column = new DataColumn();
            column.ColumnName = "ItemGuid";
            table.Columns.Add(column);

            foreach (T item in list)
            {
                DataRow row = table.NewRow();
                row["ItemGuid"] = item;
                table.Rows.Add(row);
            }
            return table;
        }
        public static DataTable CreateTable<T>()
        {
            Type entityType = typeof(T);
            DataTable table = new DataTable(entityType.Name);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entityType);

            foreach (PropertyDescriptor prop in properties)
            {
                // HERE IS WHERE THE ERROR IS THROWN FOR NULLABLE TYPES
                table.Columns.Add(prop.Name, prop.PropertyType);
            }

            return table;
        }
        public static void DataTableBulkInsert(this DataTable Table, SqlConnection connection, string destinationDataTable)
        {
            SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connection);
            sqlBulkCopy.DestinationTableName = destinationDataTable;
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            sqlBulkCopy.WriteToServer(Table);
            connection.Close();
        }
        public static void DataTableBulkInsert(this DataTable Table, string connectionsString, string destinationDataTable)
        {
            SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connectionsString, SqlBulkCopyOptions.FireTriggers);
            sqlBulkCopy.DestinationTableName = destinationDataTable;
            sqlBulkCopy.WriteToServer(Table);
        }
    }
}
