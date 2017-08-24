using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing
{
    public interface ITextProcessor
    {
        string Process(string text, EntityReference entityReference);

        string Process(string text, Dictionary<string, string> staticParameters);

        string Process(string text, EntityReference entityReference, Dictionary<string, string> staticParameters);
    }
}
