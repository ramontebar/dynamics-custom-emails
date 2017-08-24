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
    using Microsoft.Xrm.Sdk.Metadata;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Metadata.Query;
    using Microsoft.Xrm.Sdk.Messages;
    using Newtonsoft.Json;
    using DXTools.CRM.Solutions.CustomEmails.Common.Helpers;
    using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing;
    using Microsoft.Crm.Sdk.Messages;

    /// <summary>
    /// Create a custom CRM email record based on the input parameters.
    /// 
    /// Optionally, the email can be sent instantly depending on "SendInstantly" property.
    /// </summary>
    public sealed class CreateCustomEmailActivity : BaseCodeActivity
    {
        #region In Arguments

        [Input("Email Template Name")]
        [Default("")]
        public InArgument<String> EmailTemplateName { get; set; }

        [Input("Sender Email Address")]
        [Default("")]
        public InArgument<String> SenderEmailAddress { get; set; }

        [Input("To Recipient Record URL")]
        [Default("")]
        public InArgument<String> ToRecipientRecordURL { get; set; }

        [Input("Email Regarding Record URL")]
        [Default("")]
        public InArgument<String> EmailRegardingRecordURL { get; set; }

        [Input("Record Context URL")]
        [Default("")]
        public InArgument<String> RecordContextURL { get; set; }

        [Input("Attachments")]
        [Default("")]
        public InArgument<String> Attachments { get; set; }

        [Input("Allow Duplicate Attachments")]
        [Default("false")]
        public InArgument<Boolean> AllowDuplicateAttachments { get; set; }

        [Input("Only Use Attachments In Template")]
        [Default("false")]
        public InArgument<Boolean> OnlyUseAttachmentsInTemplate { get; set; }

        [Input("Static Parameters")]
        [Default("")]
        public InArgument<String> StaticParameters { get; set; }

        #region Send email instantly

        /// <summary>
        /// Once the email is created, it will be sent instantly.
        /// </summary>
        [Input("Send Instantly")]
        [Default("false")]
        public InArgument<bool> SendInstantly { get; set; }

        /// <summary>
        /// Gets or sets whether to send the email, or to just record it as sent.
        /// 
        /// "true" if the email should be sent; otherwise, "false", just record it as sent
        /// </summary>
        [Input("Issue Send")]
        [Default("true")]
        public InArgument<bool> IssueSend { get; set; }

        [Input("Tracking Token")]
        [Default("")]
        public InArgument<String> TrackingToken { get; set; }

        #endregion

        #endregion

        #region Out Arguments

        [Output("New Created Email")]
        [ReferenceTarget("email")]
        public OutArgument<EntityReference> NewCreatedEmail { get; set; }

        /// <summary>
        /// Subject of the email sent, including tracking token. This attribute is only populated when "SendInstantly" input parameter is true.
        /// </summary>
        [Output("Subject of the sent email")]
        [Default("")]
        public OutArgument<String> Subject { get; set; }

        #endregion

        #region Main Execute Logic

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            string emailTemplateName = GetEmailTemplateName(executionContext);
            EntityReference fromSender = GetSender(executionContext);
            EntityReferenceCollection toRecipient = GetRecipient(executionContext);
            EntityReference regarding = GetRegarding(executionContext);
            EntityReference recordContext = GetRecordContext(executionContext);
            string[] attachmentsArray = GetAttachments(executionContext);
            bool allowDuplicateAttachments = GetAllowDuplicateAttachments(executionContext);
            bool onlyUseAttachmentsInTemplate = GetOnlyUseAttachmentsInTemplate(executionContext);
            Dictionary<string, string> staticParameters = GetStaticParameters(executionContext);
            
            CreateCustomEmail createCustomEmail = new CreateCustomEmail(this.OrganizationService, new TextProcessor(this.OrganizationService));
            Entity newEmail = createCustomEmail.CreateEmail(emailTemplateName, this.WorkflowContext.UserId, staticParameters, recordContext, regarding, fromSender, toRecipient, attachmentsArray, allowDuplicateAttachments, onlyUseAttachmentsInTemplate);
            EntityReference newEmailReference = newEmail.ToEntityReference();

            this.NewCreatedEmail.Set(executionContext, newEmailReference);

            SendEmail(newEmailReference, executionContext);
        }

        #endregion

        #region Private Helper Methods

        private bool GetAllowDuplicateAttachments(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting AllowDuplicateAttachments argument");
            if (this.WorkflowContext.InputParameters.Contains("AllowDuplicateAttachments"))
            {
                return (Boolean)this.WorkflowContext.InputParameters["AllowDuplicateAttachments"];
            }
            else
            {
                return this.AllowDuplicateAttachments.Get<Boolean>(executionContext);
            }
        }

        private bool GetOnlyUseAttachmentsInTemplate(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting OnlyUseAttachmentsInTemplate argument");
            if (this.WorkflowContext.InputParameters.Contains("OnlyUseAttachmentsInTemplate"))
            {
                return (Boolean)this.WorkflowContext.InputParameters["OnlyUseAttachmentsInTemplate"];
            }
            else
            {
                return this.OnlyUseAttachmentsInTemplate.Get<Boolean>(executionContext);
            }
        }

        private string[] GetAttachments(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting attachments");

            string attachments = null;

            if (this.WorkflowContext.InputParameters.Contains("Attachments"))
                attachments = this.WorkflowContext.InputParameters["Attachments"].ToString();
            else
                attachments = this.Attachments.Get<string>(executionContext);

            if (string.IsNullOrEmpty(attachments))
                return null;
            else
                return attachments.Split(',');
        }

        private EntityReference GetRecordContext(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting Record Context");

            EntityReference recordContext = null;
            if (this.WorkflowContext.InputParameters.Contains("RecordContext"))
            {
                recordContext = this.WorkflowContext.InputParameters["RecordContext"] as EntityReference;
                if (recordContext.Id.Equals(Guid.Empty))
                    recordContext = null;
                return recordContext;
            }
            else
            {
                string recordContextRecordURL = this.RecordContextURL.Get<String>(executionContext);
                if (string.IsNullOrEmpty(recordContextRecordURL))
                    return null;
                else
                {
                    return ConvertURLtoEntityReference(recordContextRecordURL);
                }
            }
        }

        private EntityReference GetRegarding(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting Regarding");
            if (this.WorkflowContext.InputParameters.Contains("Regarding"))
            {
                EntityReference regarding = this.WorkflowContext.InputParameters["Regarding"] as EntityReference;
                if (regarding.Id.Equals(Guid.Empty))
                    return null;
                else
                    return regarding;
            }
            else
            {
                string regardingRecordURL = this.EmailRegardingRecordURL.Get<String>(executionContext);
                if (string.IsNullOrEmpty(regardingRecordURL))
                    return null;
                else
                {
                    return ConvertURLtoEntityReference(regardingRecordURL);
                }
            }
        }

        private EntityReferenceCollection GetRecipient(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting Recipient");
            EntityReferenceCollection toRecipient = new EntityReferenceCollection();
            if (this.WorkflowContext.InputParameters.Contains("ToRecipient"))
            {
                EntityCollection toRecipientEntityCollection = this.WorkflowContext.InputParameters["ToRecipient"] as EntityCollection;
                foreach (Entity toRecipientEntity in toRecipientEntityCollection.Entities)
                    toRecipient.Add(toRecipientEntity.ToEntityReference());
                return toRecipient;
            }
            else
            {
                string toRecipientRecordURL = this.ToRecipientRecordURL.Get<String>(executionContext);
                if (string.IsNullOrEmpty(toRecipientRecordURL))
                    return null;
                else
                {
                    toRecipient.Add(ConvertURLtoEntityReference(toRecipientRecordURL));
                    return toRecipient;
                }
            }
        }

        private EntityReference GetSender(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting Sender");
            if (this.WorkflowContext.InputParameters.Contains("FromSender"))
            {
                EntityReference fromSender = this.WorkflowContext.InputParameters["FromSender"] as EntityReference;
                if (fromSender.Id.Equals(Guid.Empty))
                    return null;
                else
                    return fromSender;
            }
            else
            {
                string senderEmailAddress = this.SenderEmailAddress.Get<String>(executionContext);
                if (string.IsNullOrEmpty(senderEmailAddress))
                    throw new ArgumentNullException("Sender Email Address", "'Sender Email Address' cannot be null, a user or queue mail must be provided");

                QueryExpression userQuery = new QueryExpression("systemuser");
                userQuery.NoLock = true;
                userQuery.Criteria.AddCondition("internalemailaddress", ConditionOperator.Equal, senderEmailAddress);

                EntityCollection results = OrganizationService.RetrieveMultiple(userQuery);
                if (results.Entities.Count > 0)
                    return results.Entities[0].ToEntityReference();
                else
                {
                    QueryExpression queueQuery = new QueryExpression("queue");
                    queueQuery.NoLock = true;
                    queueQuery.Criteria.AddCondition("emailaddress", ConditionOperator.Equal, senderEmailAddress);

                    results = OrganizationService.RetrieveMultiple(queueQuery);
                    if (results.Entities.Count > 0)
                        return results.Entities[0].ToEntityReference();
                    else
                        throw new Exception(string.Format("User or Queue couldn't be found with email address {0}", senderEmailAddress));
                }
            }
        }

        private string GetEmailTemplateName(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting Email Template Name");
            if (this.WorkflowContext.InputParameters.Contains("EmailTemplateName"))
                return this.WorkflowContext.InputParameters["EmailTemplateName"] as string;
            else
                return this.EmailTemplateName.Get<String>(executionContext);

        }

        private Dictionary<string, String> GetStaticParameters(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting Static Parameters");
            string jsonTypeFormattersString;

            if (this.WorkflowContext.InputParameters.Contains("StaticParameters"))
                jsonTypeFormattersString = this.WorkflowContext.InputParameters["StaticParameters"] as string;
            else
                jsonTypeFormattersString = this.StaticParameters.Get<String>(executionContext);

            if (!string.IsNullOrEmpty(jsonTypeFormattersString))
            {
                TraceService.Trace("Static Parameters (JSON string): {0}", jsonTypeFormattersString);
                try
                {
                    Dictionary<string, String> staticParameters = JsonConvert.DeserializeObject(jsonTypeFormattersString, typeof(Dictionary<string, String>)) as Dictionary<string, String>;

                    TraceService.Trace("Static Parameters have been deserialised. Count: {0}", staticParameters == null ? "Null" : staticParameters.Count.ToString());

                    return staticParameters;
                }
                catch (Exception exc)
                {
                    throw new InvalidPluginExecutionException(string.Format("There was an exception deserialising the Static Parameters '{0}'. Error Message: {1}. Exception type: {2}", jsonTypeFormattersString, exc.Message, exc.GetType()));
                }
            }
            else
                return null;            
        }

        private EntityReference ConvertURLtoEntityReference(string recordContextRecordURL)
        {
            Uri validatedRecordContextRecordURL;
            if (Uri.TryCreate(recordContextRecordURL, UriKind.Absolute, out validatedRecordContextRecordURL))
            {
                const string guidStart = "%7b";
                const string guidEnd = "%7d";

                string id = QueryStringHelper.GetQueryStringParameter(recordContextRecordURL, "id").Replace(guidStart, string.Empty).Replace(guidEnd, string.Empty);
                string etn = QueryStringHelper.GetQueryStringParameter(recordContextRecordURL, "etn");

                if (string.IsNullOrEmpty(etn))
                {
                    string etc = QueryStringHelper.GetQueryStringParameter(recordContextRecordURL, "etc");
                    int etcInteger;
                    if (int.TryParse(etc, out etcInteger))
                    {
                        etn = CRMMetadataHelper.GetEntityNameFromETC(etcInteger, this.OrganizationService);
                    }
                    else
                        throw new ArgumentException(string.Format("Record Context URL has an unexpected etc parameter. Object Type Code '{0}' must be an integer", etc));
                }

                return new EntityReference(etn, new Guid(id));
            }
            else
                throw new ArgumentException(string.Format("Invalid Record Context URL '{0}'", recordContextRecordURL));
        }

        private void SendEmail(EntityReference newEmailReference, CodeActivityContext executionContext)
        {
            TraceService.Trace("Entering SendEmail method");
            if ((this.WorkflowContext.InputParameters.Contains("SendInstantly")
                && ((bool)this.WorkflowContext.InputParameters["SendInstantly"]) == true)
                ||
                (this.SendInstantly != null
                 &&
                 this.SendInstantly.Get<Boolean>(executionContext) == true)
                )
            {
                TraceService.Trace("Email is going to be sent");

                SendEmailRequest sendRequest = new SendEmailRequest()
                {
                    EmailId = newEmailReference.Id
                    ,
                    IssueSend = GetIssueSend(executionContext)
                    ,
                    TrackingToken = GetTrackingToken(executionContext)
                };

                SendEmailResponse sendEmailResponse = this.OrganizationService.Execute(sendRequest) as SendEmailResponse;

                this.Subject.Set(executionContext, sendEmailResponse.Subject);

                TraceService.Trace("Email has been sent correclty with subject " + sendEmailResponse.Subject);
            }
            else
                this.Subject.Set(executionContext, string.Empty);
        }

        private Boolean GetIssueSend(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting IssueSend input argument");
            if (this.WorkflowContext.InputParameters.Contains("IssueSend"))
                return (Boolean)this.WorkflowContext.InputParameters["IssueSend"];
            else
                return this.IssueSend.Get<Boolean>(executionContext);
        }

        private string GetTrackingToken(CodeActivityContext executionContext)
        {
            TraceService.Trace("Getting TrackingToken input argument");
            if (this.WorkflowContext.InputParameters.Contains("TrackingToken"))
                return this.WorkflowContext.InputParameters["TrackingToken"] as string;
            else
                return this.TrackingToken.Get<String>(executionContext);
        }

        #endregion
    }
}