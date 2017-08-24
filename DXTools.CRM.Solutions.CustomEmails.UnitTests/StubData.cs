using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.UnitTests
{
    public class StubData
    {
        public Entity Record { get; set; }
        public AttributeMetadata[] AttributesMetadata { get; set; }
    }
}
