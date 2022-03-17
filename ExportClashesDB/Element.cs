using Autodesk.Navisworks.Api;

namespace ExportClashesDB
{
    public class Element
    {
        private ModelItem x;
        public string ItemGuid { get; set; }

        public string Id { get; set; }

        public string WorksetName { get; set; }

        public string Category { get; set; }

        public string SourceFile { get; set; }

        public string FamilyName { get; set; }

        public string TypeName { get; set; }

        public Element(ModelItem x)
        {
            this.x = x;
            ItemGuid = x.InstanceGuid.ToString();
            var elementId = x?.PropertyCategories.FindPropertyByName("LcRevitData_Element", "LcRevitPropertyElementId")?.Value;
            Id =  elementId?.DataType == VariantDataType.Int32 ? elementId?.ToInt32().ToString() : elementId?.ToDisplayString();
            WorksetName = x?.PropertyCategories?.FindPropertyByDisplayName("Объект", "Рабочий набор")?.Value?.ToDisplayString();
            Category = x?.PropertyCategories?.FindPropertyByName("LcRevitData_Element", "LcRevitPropertyElementCategory")?.Value?.ToDisplayString();
            SourceFile = x?.PropertyCategories?.FindPropertyByName("LcOaNode", "LcOaNodeSourceFile")?.Value?.ToDisplayString();
            FamilyName = x?.PropertyCategories?.FindPropertyByName("LcRevitData_Element", "LcRevitPropertyElementFamily")?.Value?.ToDisplayString();
            TypeName = x?.PropertyCategories?.FindPropertyByName("LcRevitData_Type", "LcRevitPropertyElementName")?.Value?.ToDisplayString();
        }
    }
}