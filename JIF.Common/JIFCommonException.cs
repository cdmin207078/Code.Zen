using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace JIF.Common
{
    [Serializable]
    internal class JIFCommonException : Exception
    {
        public JIFCommonException()
            : base()
        {

        }

        public JIFCommonException(string message)
            : base(message)
        {

        }

        public JIFCommonException(string messageFormat, params object[] args)
            : base(string.Format(messageFormat, args))
        {

        }

        protected JIFCommonException(SerializationInfo
            info, StreamingContext context)
            : base(info, context)
        {

        }

        public JIFCommonException(string message, Exception innerException)
            : base(message, innerException)
        {

        }
    }
}