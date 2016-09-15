using System;
using System.Xml;

namespace UpgradeStubs
{
	public static class System_Exception
	{

		public static int geterrorCode(this Exception instance)
		{
			UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("MSXML2.IXMLDOMParseError.errorCode");
			return 0;
		}
	} 
	public static class System_Xml_XmlDocument
	{

		public static Exception getparseError(this XmlDocument instance)
		{
			UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("MSXML2.DOMDocument.parseError");
			return null;
		}
		public static int getreadyState(this XmlDocument instance)
		{
			UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("MSXML2.DOMDocument.readyState");
			return 0;
		}
	}
}