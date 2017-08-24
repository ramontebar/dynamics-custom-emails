using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.Helpers
{
    public class QueryStringHelper
    {
        public static string GetQueryStringParameter(string url, string parameterName)
        {
            int startOfParametersIndex = url.LastIndexOf('?');
            string parameters = url.Substring(startOfParametersIndex);
            string[] parametersArray = parameters.Split('&');
            foreach (string parameter in parametersArray)
            {
                if (parameter.Contains(parameterName + '='))
                {
                    string[] keyValueParameter = parameter.Split('=');
                    if (keyValueParameter.Length == 2)
                        return keyValueParameter[1];
                    else
                        return string.Empty;
                }
            }
            return null;
        }
    }
}
