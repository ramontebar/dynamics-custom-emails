using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions
{
    public abstract class Symbol
    {
        public string Value { get; protected set; }

        public abstract object Resolve(SymbolContext context);
    }

}
