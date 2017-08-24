// <copyright file="SendCustomEmailActivity.cs" company="">
// Copyright (c) 2014 All Rights Reserved
// </copyright>
// <author></author>
// <date>1/29/2014 8:38:29 AM</date>
// <summary>Implements the SendCustomEmailActivity Workflow Activity.</summary>
namespace DXTools.CRM.Solutions.CustomEmails.Workflows
{
    using System;
    using System.Activities;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Workflow;
    using System.Text;
    using System.Collections.Generic;
    using DXTools.CRM.Solutions.CustomEmails.Common;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;
    using DXTools.CRM.Solutions.CustomEmails.Common.Helpers;

    /// <summary>
    /// Check if the given filename(s) exist as CRM email template attachments
    /// </summary>
    public sealed class DoesAttachmentExistActivity : BaseCodeActivity
    {
        [RequiredArgument]
        [Input("Attachments Filename")]
        [Default("")]
        public InArgument<String> AttachmentsFilename { get; set; }

        [Output("All Attachments Found")]
        [Default("false")]
        public OutArgument<Boolean> AllAttachmentsFound { get; set; }

        [Output("Attachments Not Found")]
        [Default("")]
        public OutArgument<String> AttachmentsNotFound { get; set; }


        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            if (this.AttachmentsFilename == null)
                throw new ArgumentNullException("AttachmentsFilename", "'AttachmentsFilename' input argument cannot be null");

            string attachments = this.AttachmentsFilename.Get(executionContext);

            if(string.IsNullOrEmpty(attachments))
            {
                this.AllAttachmentsFound.Set(executionContext, false);
                this.AttachmentsNotFound.Set(executionContext, "Attachments Filename is null or empty");
                return;
            }

            string[] attachmentsArray = attachments.Split(',');

            EntityCollection activityMimeAttachments = ActivityMimeAttachment.RetrieveTemplatesActivityMimeAttachmentsByName(attachmentsArray, this.OrganizationService);
            if (activityMimeAttachments == null || activityMimeAttachments.TotalRecordCount == 0)
            {
                this.AllAttachmentsFound.Set(executionContext, false);
                this.AttachmentsNotFound.Set(executionContext, attachments);
                return;
            }

            DataCollection<Entity> activityMimeAttachmentEntities = activityMimeAttachments.Entities;
            StringBuilder attachmentsNotFound = new StringBuilder();
            bool found = false;

            foreach (string attachment in attachmentsArray)
            {
                foreach (Entity activityMimeAttachment in activityMimeAttachmentEntities)
                {
                    if (activityMimeAttachment.Contains("filename")
                        && activityMimeAttachment["filename"].ToString().ToLowerInvariant() == attachment.ToLowerInvariant())
                    {
                        found = true;
                        break;
                    }
                }

                if (found == false)
                {
                    if (attachmentsNotFound.Length == 0)
                        attachmentsNotFound.Append(attachment);
                    else
                        attachmentsNotFound.Append(string.Format(",{0}", attachment));
                }
                found = false;
            }

            if (attachmentsNotFound.Length > 0)
            {
                this.AllAttachmentsFound.Set(executionContext, false);
                this.AttachmentsNotFound.Set(executionContext, attachmentsNotFound.ToString());
            }
            else
            {
                this.AllAttachmentsFound.Set(executionContext, true);
            }
        }
    }
}