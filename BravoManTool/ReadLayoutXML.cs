using System;
using Microsoft.SqlServer.Server;
using System.Data.SqlTypes;
using System.Xml;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace BravoManTool
{
    public  class ReadLayoutXML
    {
        [SqlFunction(FillRowMethodName = "FillRowExtract", TableDefinition = "CommandValue NVARCHAR(MAX), CommandPath NVARCHAR(MAX), ParentTagName1 NVARCHAR(MAX), ParentTagName2 NVARCHAR(MAX), ParentTagName3 NVARCHAR(MAX), ParentTagName4 NVARCHAR(MAX), ParentTagName5 NVARCHAR(MAX)")]
        public static IEnumerable ExtractValuesAndPaths(SqlXml xmlData, SqlString tagName)
        {
            var results = new List<Tuple<string, string, string, string, string, string, string>>();
            if (xmlData.IsNull || string.IsNullOrWhiteSpace(tagName.Value))
                return results;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlData.Value);
            XmlNodeList nodes = xmlDoc.SelectNodes($"//{tagName.Value}");

            foreach (XmlNode node in nodes)
            {
                string path = GetNodePath(node);
                string value = node.InnerText;
                string parentTagName1 = GetParentTagName(node, 1);
                string parentTagName2 = GetParentTagName(node, 2);
                string parentTagName3 = GetParentTagName(node, 3);
                string parentTagName4 = GetParentTagName(node, 4);
                string parentTagName5 = GetParentTagName(node, 5);
                results.Add(Tuple.Create(value, path, parentTagName1, parentTagName2, parentTagName3, parentTagName4, parentTagName5));
            }

            return results;
        }

        public static void FillRowExtract(object rowObj, out SqlString commandValue, out SqlString commandPath, out SqlString parentTagName1, out SqlString parentTagName2, out SqlString parentTagName3, out SqlString parentTagName4, out SqlString parentTagName5)
        {
            Tuple<string, string, string, string, string, string, string> row = (Tuple<string, string, string, string, string, string, string>)rowObj;
            commandValue = new SqlString(row.Item1);
            commandPath = new SqlString(row.Item2);
            parentTagName1 = new SqlString(row.Item3);
            parentTagName2 = new SqlString(row.Item4);
            parentTagName3 = new SqlString(row.Item5);
            parentTagName4 = new SqlString(row.Item6);
            parentTagName5 = new SqlString(row.Item7);
        }

        private static string GetNodePath(XmlNode node)
        {
            if (node == null || node.NodeType == XmlNodeType.Document)
            {
                return string.Empty;
            }

            string path = GetNodePath(node.ParentNode);
            return (path == string.Empty ? "" : path + "/") + node.Name;
        }

        private static string GetParentTagName(XmlNode node, int level)
        {
            XmlNode parentNode = node;
            for (int i = 0; i < level; i++)
            {
                if (parentNode.ParentNode != null && parentNode.ParentNode.NodeType != XmlNodeType.Document)
                {
                    parentNode = parentNode.ParentNode;
                }
                else
                {
                    return null;
                }
            }
            return parentNode.Name;
        }

        /*Xử lý đọc layout main lấy ra menu các chức năng*/
        [Microsoft.SqlServer.Server.SqlFunction(
        FillRowMethodName = "FillRow",
        TableDefinition = "FormName nvarchar(100), LayoutName nvarchar(100), FormText nvarchar(max), DLLName nvarchar(100), IsTemplate bit, PathName nvarchar(max)",
        DataAccess = DataAccessKind.Read)]
        public static IEnumerable FetchAndProcessData(SqlString serverName, SqlString databaseName)
        {
            string connectionString;
            if (!serverName.IsNull && !serverName.Value.Equals(string.Empty))
            {
                connectionString = $"Server={serverName};Integrated Security=true;";
            }
            else
            {
                connectionString = "context connection=true";
            }

            var databasePrefix = !databaseName.IsNull && !databaseName.Value.Equals(string.Empty) ? $"[{databaseName.Value}]." : string.Empty;

            var query = $@"
            SELECT 
                a.Id,
                a.IsTemplate,
                a.FormName,
                a.LayoutName,
                a.LayoutData,
                a.IsActive,
                dbo.ufn_sys_ToLayoutXml(a.LayoutData) AS LayoutDataXml,
                IIF(a.IsTemplate=0 AND a.FormName='MainWindow', 1, 0) IsLayoutMain
            FROM {databasePrefix}dbo.vB00Layout AS a
                INNER JOIN {databasePrefix}dbo.B00Command c ON c.CommandKey = a.FormName AND c.IsActive = 1";

            List<object[]> results = new List<object[]>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var layoutDataXml = reader["LayoutDataXml"].ToString();
                            var layoutName = reader["LayoutName"].ToString();
                            var isLayoutMain = (bool)reader["IsLayoutMain"];
                            var isTemplate = (bool)reader["IsTemplate"];
                            var formName = reader["FormName"].ToString();

                            if (isLayoutMain)
                            {
                                XmlDocument xmlDoc = new XmlDocument();
                                xmlDoc.LoadXml(layoutDataXml);
                                // Giả sử xmlDoc là XmlDocument đã load XML của bạn
                                XmlNode root = xmlDoc.DocumentElement;
                                if (root != null)
                                {
                                    // Xử lý menuStrip1
                                    XmlNode menuStrip1Node = root.SelectSingleNode("menuStrip1");
                                    if (menuStrip1Node != null)
                                    {
                                        ProcessXmlNode(menuStrip1Node, "LayoutNameMain", results, "MENU");
                                    }

                                    // Xử lý navigator
                                    XmlNode navigatorNode = root.SelectSingleNode("navigator");
                                    if (navigatorNode != null)
                                    {
                                        ProcessXmlNode(navigatorNode, "LayoutNameMain", results, "NAV");
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return results;
        }

        private static void ProcessXmlNode(XmlNode node, string layoutName, List<object[]> results, string currentPath)
        {
            // Đối với mỗi node, kiểm tra xem nó có thẻ con <Items> không.
            // Nếu không, đây là leaf node.
            bool hasItemChild = false;
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.Name == "Items")
                {
                    hasItemChild = true;
                    // Duyệt qua thẻ <Items>
                    foreach (XmlNode itemChild in childNode.ChildNodes)
                    {
                        ProcessXmlNode(itemChild, layoutName, results, currentPath);
                    }
                }
            }

            // Nếu node hiện tại không có thẻ con <Items>, coi nó là leaf node.
            if (!hasItemChild && node.Name != "root" && node.Name != "Text")
            {
                // Lấy tên hoặc giá trị Vietnamese nếu có
                var nodeText = node.SelectSingleNode("Text/Vietnamese")?.InnerText ?? node.Name;
                var newPath = string.IsNullOrEmpty(currentPath) ? nodeText : $"{currentPath}\\{nodeText}";

                // Thêm vào kết quả. Đối với leaf node, newPath chính là đường dẫn cuối cùng.
                results.Add(new object[] { node.Name, layoutName, newPath });
            }
        }

        public static void FillRow(object rowObj, out SqlString formName, out SqlString layoutName, out SqlString formText, out SqlString dllName, out SqlBoolean isTemplate, out SqlString pathName)
        {
            object[] row = (object[])rowObj;
            formName = new SqlString((string)row[0]);
            layoutName = new SqlString((string)row[1]);
            formText = new SqlString((string)row[2]);
            dllName = new SqlString((string)row[3]);
            isTemplate = new SqlBoolean((bool)row[4]);
            pathName = new SqlString((string)row[5]);
        }
    }
}
