using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing
{
    public class FreeText : Symbol
    {
        public FreeText(string text)
        {
            this.Value = text;
        }

        public override object Resolve(SymbolContext context)
        {
            return this.Value;
        }
    }
}
