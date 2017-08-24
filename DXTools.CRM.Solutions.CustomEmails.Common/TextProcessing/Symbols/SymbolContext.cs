using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions
{
    public class SymbolContext
    {
        public EntityReference RecordRefence { get; set; }

        public IOrganizationService OrganizationService { get; set; }

        public Dictionary<string, string> StaticParameters { get; set; }
    }
}
