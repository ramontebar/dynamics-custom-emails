using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Functions
{
    public class ToLocalDateTime : IDynFunctionHandler
    {
        public object Invoke(params object[] parameters)
        {

            if (parameters == null || parameters.Length == 0)
                return string.Empty;

            object dynamicType = parameters[0];

            if (dynamicType is DateTime)
            {
                if (parameters.Length == 2)
                {
                    string timeZoneId = parameters[1] as string;
                    DateTime dateTimeUTC = ((DateTime)dynamicType);
                    TimeZoneInfo localTimeZone = TimeZoneInfo.FindSystemTimeZoneById(timeZoneId);
                    DateTime localDateTime = TimeZoneInfo.ConvertTimeFromUtc(dateTimeUTC, localTimeZone);
                    return localDateTime;
                }
            }


            return dynamicType.ToString();

        }
    }
}
