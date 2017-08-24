using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;

namespace DXTools.CRM.Solutions.CustomEmails.Common
{
    public class TraceService
    {
        static TraceService _Current;
        ITracingService _TraceService;

        private TraceService(ITracingService Service)
        {
            _TraceService = Service;
        }

        public static TraceService Initialise(ITracingService Service)
        {
            _Current = new TraceService(Service);
            return _Current;
        }


        public static TraceService Current
        {
            get
            {
                return _Current;
            }
        }

        public static void Trace(string format, params object[] args)
        {
            if (Current != null)
                Current._TraceService.Trace(format, args);
        }
    }
}
