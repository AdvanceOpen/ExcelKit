using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelKit.Core.Infrastructure.Exceptions
{
	[Serializable]
	public class ExcelKitException : ApplicationException
	{
		public ExcelKitException()
		{

		}

		public ExcelKitException(string message)
			: base(message)
		{
		}

		public ExcelKitException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		public ExcelKitException(string messageFormat, params object[] args)
			: base(string.Format(messageFormat, args))
		{
		}
	}
}
