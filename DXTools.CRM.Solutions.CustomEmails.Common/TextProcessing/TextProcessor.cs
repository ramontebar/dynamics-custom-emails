using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.Expressions;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing
{
    public class TextProcessor : ITextProcessor
    {
        public ISyntacticParser SyntacticParser { get; private set; }

        public IOrganizationService OrganizationService { get; private set; }

        public TextProcessor(IOrganizationService organizationService)
        {
            this.SyntacticParser = new SyntacticParser(new LexicalParser());
            this.OrganizationService = organizationService;
        }

        public string Process(string text, EntityReference entityReference)
        {
            return Process(text, entityReference, null);
        }

        public string Process(string text, Dictionary<string, string> staticParameters)
        {
            return Process(text, null, staticParameters);
        }

        public string Process(string text, EntityReference entityReference, Dictionary<string, string> staticParameters)
        {
            List<Symbol> symbols = GetListOfSymbols(text);

            SymbolContext symbolContext = new SymbolContext()
            {
                OrganizationService = this.OrganizationService
                ,
                RecordRefence = entityReference
                ,
                StaticParameters = staticParameters

            };

            StringBuilder processedTextBuilder = new StringBuilder();
            foreach (Symbol symbol in symbols)
            {
                processedTextBuilder.Append(symbol.Resolve(symbolContext));
            }

            return processedTextBuilder.ToString();
        }

        private List<Symbol> GetListOfSymbols(string text)
        {
            //Allow to create dynamic links. See https://crmcustomemails.codeplex.com/discussions/571696
            text = text.Replace("%7b", "{").Replace("%7d", "}"); 

            int index = 0;
            int textLength = text.Length;
            List<Symbol> symbols = new List<Symbol>();

            char currentCharacter;
            string currentSimbol = string.Empty;

            while (index < textLength)
            {
                currentCharacter = text[index];

                if (currentCharacter == '{')
                {
                    if (!string.IsNullOrEmpty(currentSimbol))
                    {
                        symbols.Add(new FreeText(currentSimbol));
                        currentSimbol = string.Empty;
                    }

                    do
                    {
                        index++;
                        currentCharacter = text[index];
                        if (currentCharacter != '}')
                            currentSimbol += currentCharacter;
                    }
                    while (currentCharacter != '}'
                          && index < textLength);

                    symbols.Add(SyntacticParser.Parse(System.Net.WebUtility.HtmlDecode(currentSimbol)));
                    currentSimbol = string.Empty;
                }
                else
                    currentSimbol += currentCharacter;

                index++;
            }
            if (!string.IsNullOrEmpty(currentSimbol))
                symbols.Add(new FreeText(currentSimbol));

            return symbols;
        }

    }
}
