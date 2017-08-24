using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Functions
{
    public class ToString : IDynFunctionHandler
    {
        public object Invoke(params object[] parameters)
        {
            if (parameters == null || parameters.Length == 0)
                return string.Empty;

            object dynamicType = parameters[0];

            if (dynamicType is decimal)
            {
                if (parameters.Length == 2)
                {
                    string format = parameters[1] as string;
                    return ((decimal)dynamicType).ToString(format);
                }
            }

            if (dynamicType is DateTime)
            {
                if (parameters.Length == 2)
                {
                    string format = parameters[1] as string;
                    DateTime dateTime = (DateTime)dynamicType;
                    return dateTime.ToString(format);
                }
            }

            if (dynamicType is EntityCollection)
            {
                EntityCollection entityCollection = dynamicType as EntityCollection;
                if (entityCollection.EntityName == "activityparty")
                {
                    string splitSeparator;
                    if (parameters.Length == 2)
                        splitSeparator = parameters[1] as string;
                    else
                        splitSeparator = ", ";

                    StringBuilder partyNameCollection = new StringBuilder();
                    int numberOfParties = entityCollection.Entities.Count;
                    Entity party;
                    for (int i = 0; i < numberOfParties; i++)
                    {
                        party = entityCollection.Entities[i];
                        if (party.Contains("partyid"))
                            partyNameCollection.Append(((EntityReference)party["partyid"]).Name);
                        if (i + 1 < numberOfParties)
                            partyNameCollection.Append(splitSeparator);
                    }

                    return partyNameCollection.ToString();
                }
            }

            return dynamicType.ToString();
            
        }
    }
}
