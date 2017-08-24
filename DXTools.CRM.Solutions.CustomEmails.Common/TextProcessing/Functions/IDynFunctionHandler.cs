using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Functions
{
    public interface IDynFunctionHandler
    {
        object Invoke(params object[] parameters);
    }
}
