using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing
{
    public interface ISyntacticParser
    {
        Symbol Parse(string text);
    }
}
