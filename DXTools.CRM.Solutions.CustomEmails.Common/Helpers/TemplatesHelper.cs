using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.Helpers
{
    public static class TemplatesHelper
    {
        public static string GetSubjectXML(string rowSubject)
        {
            return "<?xml version=\"1.0\" ?><xsl:stylesheet xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\" version=\"1.0\"><xsl:output method=\"text\" indent=\"no\"/><xsl:template match=\"/data\"><![CDATA[" + rowSubject + "]]></xsl:template></xsl:stylesheet>";
        }

        public static string GetSubjectPresentationXML(string rowSubject)
        {
            return "<template><text><![CDATA[" + rowSubject + "]]></text></template>";
        }

        public static string GetBodyXML(string rowBody)
        {
            return "<?xml version=\"1.0\" ?><xsl:stylesheet xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\" version=\"1.0\"><xsl:output method=\"text\" indent=\"no\"/><xsl:template match=\"/data\"><![CDATA[" + rowBody + "]]></xsl:template></xsl:stylesheet>";
        }

        public static string GetBodyPresentationXML(string rowBody)
        {
            return "<template><text><![CDATA[" + rowBody + "]]></text></template>";
        }

        /// <summary>
        /// Retrieve the first CRM email template (all columns) with the given name
        /// </summary>
        /// <param name="templateName"></param>
        /// <param name="organizationService"></param>
        /// <returns></returns>
        public static Entity RetrieveTemplateByName(string templateName, IOrganizationService organizationService)
        {
            QueryExpression query = new QueryExpression("template");
            query.NoLock = true;
            query.ColumnSet = new ColumnSet(true);
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition(new ConditionExpression("title", ConditionOperator.Equal, templateName));

            EntityCollection templateEntityCollection = organizationService.RetrieveMultiple(query);
            if (templateEntityCollection.Entities != null && templateEntityCollection.Entities.Count > 0)
                return templateEntityCollection.Entities[0];

            return null;
        }
    }
}
