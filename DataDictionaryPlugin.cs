using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class DataDictionaryPlugin : IPlugin
{
    public void Execute(IServiceProvider serviceProvider)
    {
        // 1. Get context and input
        var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        var service = serviceFactory.CreateOrganizationService(context.UserId);

        string solutionUniqueName = context.InputParameters["SolutionUniqueName"] as string;
        if (string.IsNullOrEmpty(solutionUniqueName))
            throw new InvalidPluginExecutionException("SolutionUniqueName is required.");

        // 2. Get SolutionId
        var solution = GetSolutionByUniqueName(service, solutionUniqueName);
        if (solution == null)
            throw new InvalidPluginExecutionException($"Solution '{solutionUniqueName}' not found.");

        Guid solutionId = solution.Id;

        // 3. Get Entities and Attributes in Solution
        var entityIds = GetSolutionComponentIds(service, solutionId, 1); // Entities
        var attributeIds = GetSolutionComponentIds(service, solutionId, 2); // Attributes

        // 4. Get Web Resources in Solution
        var webResourceIds = GetSolutionComponentIds(service, solutionId, 61);

        // 5. Retrieve metadata for entities/fields
        var fields = GetFieldsMetadata(service, entityIds, attributeIds);
    }

    private List<Guid> GetSolutionComponentIds(IOrganizationService service, Guid solutionId, int componentType)
    {
        var query = new QueryExpression("solutioncomponent")
        {
            ColumnSet = new ColumnSet("objectid"),
            Criteria =
            {
                Conditions =
                {
                    new ConditionExpression("solutionid", ConditionOperator.Equal, solutionId),
                    new ConditionExpression("componenttype", ConditionOperator.Equal, componentType)
                }
            }
        };
        return query.AddOrder("objectid", OrderType.Ascending)
                    .RetrieveMultiple(service)
                    .Entities
                    .Select(e => e.GetAttributeValue<Guid>("objectid"))
                    .ToList();
    }

    private List<FieldInfo> GetFieldsMetadata(IOrganizationService service, List<Guid> entityIds, List<Guid> attributeIds)
    {
        var fields = new List<FieldInfo>();
        foreach (var entityId in entityIds)
        {
            var req = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Attributes,
                MetadataId = entityId
            };
            var resp = (RetrieveEntityResponse)service.Execute(req);
            var entityMeta = resp.EntityMetadata;
            foreach (var attr in entityMeta.Attributes)
            {
                if (!attr.IsCustomAttribute.GetValueOrDefault()) continue;
                fields.Add(new FieldInfo
                {
                    EntityLogicalName = entityMeta.LogicalName,
                    EntityDisplayName = entityMeta.DisplayName?.UserLocalizedLabel?.Label ?? entityMeta.LogicalName,
                    FieldSchemaName = attr.LogicalName,
                    FieldDisplayName = attr.DisplayName?.UserLocalizedLabel?.Label ?? attr.LogicalName,
                    DataType = attr.AttributeTypeName?.Value ?? attr.AttributeType?.ToString(),
                    RequiredLevel = attr.RequiredLevel?.Value.ToString(),
                    Description = attr.Description?.UserLocalizedLabel?.Label ?? "",
                });
            }
        }
        // Optionally, add fields from attributeIds not already included
        return fields;
    }

    private List<(string Name, string DisplayName, string Content)> GetWebResourceScripts(IOrganizationService service, List<Guid> webResourceIds)
    {
        var scripts = new List<(string, string, string)>();
        foreach (var id in webResourceIds)
        {
            var wr = service.Retrieve("webresource", id, new ColumnSet("name", "displayname", "content", "webresourcetype"));
            if (wr.GetAttributeValue<OptionSetValue>("webresourcetype")?.Value == 3) // JS
            {
                var content = wr.GetAttributeValue<string>("content");
                if (!string.IsNullOrEmpty(content))
                {
                    var js = Encoding.UTF8.GetString(Convert.FromBase64String(content));
                    scripts.Add((wr.GetAttributeValue<string>("name"), wr.GetAttributeValue<string>("displayname"), js));
                }
            }
        }
        return scripts;
    }

    private void MapScriptReferences(List<FieldInfo> fields, List<(string Name, string DisplayName, string Content)> scripts)
    {
        foreach (var field in fields)
        {
            foreach (var script in scripts)
            {
                if (Regex.IsMatch(script.Content, $@"\b{Regex.Escape(field.FieldSchemaName)}\b", RegexOptions.IgnoreCase))
                {
                    field.ScriptReferences.Add(new ScriptReference
                    {
                        WebResourceName = script.Name,
                        Note = $"Referenced in {script.Name}"
                    });
                }
            }
        }
    }

    private Guid SaveDocumentAsNote(IOrganizationService service, byte[] docBytes, string fileName, string subject)
    {
        var note = new Entity("annotation");
        note["subject"] = subject;
        note["filename"] = fileName;
        note["mimetype"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        note["documentbody"] = Convert.ToBase64String(docBytes);
        return service.Create(note);
    }
}

public class FieldInfo
{
    public string EntityLogicalName { get; set; }
    public string EntityDisplayName { get; set; }
    public string FieldSchemaName { get; set; }
    public string FieldDisplayName { get; set; }
    public string DataType { get; set; }
    public string RequiredLevel { get; set; }
    public string Description { get; set; }
    public List<ScriptReference> ScriptReferences { get; set; } = new List<ScriptReference>();
}

public class ScriptReference
{
    public string WebResourceName { get; set; }
    public string Note { get; set; }
}