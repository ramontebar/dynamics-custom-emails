using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;

namespace DXTools.CRM.Solutions.CustomEmails.UnitTests.Common
{
    [TestClass]
    public class CRMSdkBasis
    {
        [TestMethod]
        public void Test_CRM_SDK_Entity_Contains()
        {
            Entity entity = new Entity("contact");
            entity.Attributes.Add("lastname", "Smith");

            Assert.IsTrue(entity.Contains("lastname"));
            Assert.IsFalse(entity.Contains("LASTNAME"));
            Assert.IsFalse(entity.Contains("Lastname"));
        }
    }
}
