using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions
{
    public class AttributeExpression : DynamicExpression
    {
        public AttributeExpression(string name)
        {
            this.Value = string.IsNullOrEmpty(name) ? string.Empty : name.Trim().ToLowerInvariant();
            this.ChildExpressions = new List<DynamicExpression>();
        }

        public override object Resolve(SymbolContext context)
        {
            if (context.StaticParameters !=null
                && context.StaticParameters.ContainsKey(this.Value))
                return context.StaticParameters[this.Value];

            return RetrieveAttributeValue(this, context.RecordRefence, context.OrganizationService);
        }

        private object RetrieveAttributeValue(DynamicExpression attributeExpression, EntityReference recordReference, IOrganizationService organizationService)
        {
            string expressionValue = attributeExpression.Value;

            if (string.IsNullOrEmpty(expressionValue))
                return string.Empty;

            Entity record = organizationService.Retrieve(recordReference.LogicalName, recordReference.Id, new ColumnSet(true));

            if (!record.Contains(expressionValue))
                return string.Empty;

            object crmAttribute = record[expressionValue];

            if (attributeExpression.ChildExpressions.Count == 0)
            {
                if (crmAttribute.GetType() == typeof(Money))
                    return ((Money)crmAttribute).Value;
                else if (crmAttribute.GetType() == typeof(EntityReference))
                    return ((EntityReference)crmAttribute).Name;
                else if (crmAttribute.GetType() == typeof(OptionSetValue))
                    return RetrieveOptionSetLabel(recordReference, expressionValue, record[expressionValue] as OptionSetValue, organizationService);
                else
                    return crmAttribute;
            }
            else
            {
                DynamicExpression childExpression = attributeExpression.ChildExpressions[0];
                string childExpressionValue = childExpression.Value;

                if (string.IsNullOrEmpty(childExpressionValue))
                    return string.Empty;

                if (crmAttribute.GetType() == typeof(EntityReference))
                {
                    if (childExpressionValue == "id")
                        return ((EntityReference)crmAttribute).Id;
                    else if (childExpressionValue == "name")
                        return ((EntityReference)crmAttribute).Name;
                    else
                        return RetrieveAttributeValue(attributeExpression.ChildExpressions[0], crmAttribute as EntityReference, organizationService);
                }
                else if (crmAttribute.GetType() == typeof(OptionSetValue))
                {
                    if (childExpressionValue == "value")
                        return ((OptionSetValue)crmAttribute).Value;
                    else if (childExpressionValue == "label")
                        return RetrieveOptionSetLabel(recordReference, expressionValue, (OptionSetValue)crmAttribute, organizationService);
                }
                else
                    return string.Empty; //Unexpected type
            }

            return string.Empty; //Unexpected type
        }

        private string RetrieveOptionSetLabel(EntityReference entityReference, string attributeLogicalName, OptionSetValue optionSetValue, IOrganizationService organizationService)
        {
            RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityReference.LogicalName,
                LogicalName = attributeLogicalName,
                RetrieveAsIfPublished = true
            };

            RetrieveAttributeResponse attributeResponse = organizationService.Execute(attributeRequest) as RetrieveAttributeResponse;

            OptionMetadataCollection options = ((PicklistAttributeMetadata)attributeResponse.AttributeMetadata).OptionSet.Options;
            foreach (OptionMetadata option in options)
            {
                if (option != null && option.Value == optionSetValue.Value)
                {
                    if (option.Label == null
                        || option.Label.UserLocalizedLabel == null)
                        return string.Empty;
                    else
                        return option.Label.UserLocalizedLabel.Label;
                }
            }

            return string.Empty;
        }
    }
}
