using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DXTools.CRM.Solutions.CustomEmails.Common;
using DXTools.CRM.Solutions.CustomEmails.Common.Helpers;

namespace DXTools.CRM.Solutions.CustomEmails.UnitTests.Common
{
    [TestClass]
    public class QueryStringHelperTest
    {
        [TestMethod]
        public void RetrieveParameter()
        {
            string url = "https://customemails.crm4.dynamics.com:443//main.aspx?etc=10005&id=00c9121e-2698-e311-b8ea-d89d67635d80&histKey=234045132&newWindow=true&pagetype=entityrecord";

            string expectedParameter1 = "10005";
            string actualParameter1 = QueryStringHelper.GetQueryStringParameter(url, "etc");

            string expectedParameter2 = "00c9121e-2698-e311-b8ea-d89d67635d80";
            string actualParameter2 = QueryStringHelper.GetQueryStringParameter(url, "id");

            Assert.AreEqual(expectedParameter1, actualParameter1);
            Assert.AreEqual(expectedParameter2, actualParameter2);
        }
    }
}
