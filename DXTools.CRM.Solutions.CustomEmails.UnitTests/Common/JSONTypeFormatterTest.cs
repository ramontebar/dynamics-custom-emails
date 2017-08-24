using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DXTools.CRM.Solutions.CustomEmails.Common;
using Microsoft.Xrm.Sdk.Fakes;
using Microsoft.Xrm.Sdk;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Fakes;
using Microsoft.QualityTools.Testing.Fakes;
using Newtonsoft.Json;

namespace DXTools.CRM.Solutions.CustomEmails.UnitTests.Common
{
    [TestClass]
    public class JSONTypeFormatterTest
    {
        [TestMethod]
        public void SerialiseTypeFormatters()
        {

            Dictionary<AttributeTypeCode, String> formatters = new Dictionary<AttributeTypeCode, string>();
            formatters.Add(AttributeTypeCode.Money, "0.00");
            formatters.Add(AttributeTypeCode.DateTime, "dd/MM/yyyy HH:mm");
            formatters.Add(AttributeTypeCode.Owner, "name");

            string resultJsonString = JsonConvert.SerializeObject(formatters);
            string expectedJsonTring = "{\"Money\":\"0.00\",\"DateTime\":\"dd/MM/yyyy HH:mm\",\"Owner\":\"name\"}";

            Assert.AreEqual(expectedJsonTring, resultJsonString);
            
        }

        [TestMethod]
        public void DeserialiseTypeFormatters()
        {
            string jsonTring = "{\"Money\":\"0.00\",\"DateTime\":\"dd/MM/yyyy HH:mm\",\"Owner\":\"name\"}";

            object typeFormatters = JsonConvert.DeserializeObject(jsonTring,typeof(Dictionary<AttributeTypeCode, String>));

            Dictionary<AttributeTypeCode, String> result = typeFormatters as Dictionary<AttributeTypeCode, String>;

            Assert.IsNotNull(result);

            Assert.IsTrue(result.Count == 3);

        }

    }
}
