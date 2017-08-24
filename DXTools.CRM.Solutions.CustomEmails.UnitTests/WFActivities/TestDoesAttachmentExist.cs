using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Workflow;
using Xrm.Framework.Test.Unit.Fakes;
using DXTools.CRM.Solutions.CustomEmails.Workflows;
using Microsoft.Xrm.Sdk.Query;


namespace DXTools.CRM.Solutions.CustomEmails.UnitTests.WFActivities
{
    [TestClass]
    public class TestDoesAttachmentExist : WFActivityUnitTest
    {
        #region Test Context
        
        private TestContext context;
        public TestContext TestContext
        {
            get { return context; }
            set { context = value; }
        }

        #endregion

        #region Setup

        protected override Activity SetupActivity()
        {
            string attachmentsName = context.DataRow[0].ToString();
            string existingAttachmentsInCRM = context.DataRow[1].ToString();

            //Input activity parameters
            DoesAttachmentExistActivity activity = new DoesAttachmentExistActivity();
            activity.AttachmentsFilename = attachmentsName;

            //Stub Organization Service according to test data
            this.OrganizationServiceStub.RetrieveMultipleQueryBase =
            (query) =>
            {
                QueryExpression queryExpr = query as QueryExpression;
                if (queryExpr != null)
                {
                    if (queryExpr.EntityName == "activitymimeattachment")
                    {
                        EntityCollection activitymimeattachmentCollection = new EntityCollection();

                        if (!string.IsNullOrEmpty(existingAttachmentsInCRM))
                        {
                            string[] existingAttachmentsInCRMArray = existingAttachmentsInCRM.Split(',');
                            if (existingAttachmentsInCRMArray != null && existingAttachmentsInCRMArray.Length > 0)
                            {
                                foreach (string att in existingAttachmentsInCRMArray)
                                {
                                    Entity activitymimeattachmentRecord = new Entity("activitymimeattachment");
                                    activitymimeattachmentRecord["filename"] = att;
                                    activitymimeattachmentCollection.Entities.Add(activitymimeattachmentRecord);
                                    activitymimeattachmentCollection.TotalRecordCount++;
                                }
                            }
                        }

                        return activitymimeattachmentCollection;
                    }
                }
                return null; 
            };

            return activity;
        }

        #endregion

        #region Test

        [TestMethod]
        [DeploymentItem("TestData\\Attachments.xlsx")]
        [DataSource("System.Data.Odbc",
        "Dsn=Excel Files;Driver={Microsoft Excel Driver(*.xlsx)};dbq=|DataDirectory|\\TestData\\Attachments.xlsx;defaultdir=.;driverid=790;maxbuffersize=2048;pagetimeout=5;readonly=true"
        , "AttachmentTests$"
        , DataAccessMethod.Sequential)]
        public void RunTestDoesAttachmentExist_NoAttachments()
        {
            base.Test();
        }

        #endregion

        #region Verify

        protected override void Verify()
        {
            Assert.IsNull(Error);

            bool expectedOutputAllAttachmenstFound = (bool)context.DataRow[2];
            string expectedOutputAttachmenstNotFound = context.DataRow[3].ToString();

            Assert.AreEqual(expectedOutputAllAttachmenstFound, (bool)this.Outputs["AllAttachmentsFound"]);
            Assert.AreEqual(expectedOutputAttachmenstNotFound, this.Outputs["AttachmentsNotFound"] != null ? this.Outputs["AttachmentsNotFound"].ToString() : string.Empty); 
        }

        #endregion
    }
}
