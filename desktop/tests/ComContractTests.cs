using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using Vefeyat;

// Reflection/XML checks only; this does not start or automate Word.
internal static class ComContractTests
{
    private static int assertions;
    public static int Main()
    {
        try
        {
            OfficeContracts();
            RibbonCallbacks();
            Console.WriteLine("Office COM contract and ribbon checks passed: " + assertions);
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static void OfficeContracts()
    {
        Check(typeof(IDTExtensibility2).GUID == new Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), "Office lifecycle IID");
        Check(typeof(IRibbonExtensibility).GUID == new Guid("000C0396-0000-0000-C000-000000000046"), "Office ribbon IID");
        foreach (Type contract in new Type[] { typeof(IDTExtensibility2), typeof(IRibbonExtensibility) })
        {
            var kind = (InterfaceTypeAttribute)Attribute.GetCustomAttribute(contract, typeof(InterfaceTypeAttribute));
            Check(kind != null && kind.Value == ComInterfaceType.InterfaceIsDual, "Office may call via dual vtable: " + contract.Name);
        }
        string[] lifecycle = { "OnConnection", "OnDisconnection", "OnAddInsUpdate", "OnStartupComplete", "OnBeginShutdown" };
        for (int i = 0; i < lifecycle.Length; i++)
        {
            MethodInfo method = typeof(IDTExtensibility2).GetMethod(lifecycle[i]);
            var dispatch = (DispIdAttribute)Attribute.GetCustomAttribute(method, typeof(DispIdAttribute));
            Check(dispatch != null && dispatch.Value == i + 1, "Office lifecycle dispatch ID: " + lifecycle[i]);
            ParameterInfo[] args = method.GetParameters();
            Check(args[args.Length - 1].ParameterType == typeof(Array).MakeByRefType(), "Host custom argument is SAFEARRAY by reference");
        }
        Check(typeof(Connect).GUID == new Guid("2B0AF7B7-BD0D-4C70-9136-BED2DBD8F8F3"), "Installed CLSID matches class");
        var progId = (ProgIdAttribute)Attribute.GetCustomAttribute(typeof(Connect), typeof(ProgIdAttribute));
        Check(progId != null && progId.Value == "Vefeyat.Connect", "Installed ProgID matches class");
        var exposed = (ClassInterfaceAttribute)Attribute.GetCustomAttribute(typeof(Connect), typeof(ClassInterfaceAttribute));
        Check(exposed != null && exposed.Value == ClassInterfaceType.AutoDispatch, "Ribbon callbacks are exposed through IDispatch");
    }

    private static void RibbonCallbacks()
    {
        var xml = new XmlDocument();
        xml.XmlResolver = null;
        xml.LoadXml(new Connect().GetCustomUI("Microsoft.Word.Document"));
        var namespaces = new XmlNamespaceManager(xml.NameTable);
        namespaces.AddNamespace("r", "http://schemas.microsoft.com/office/2009/07/customui");
        XmlNode tab = xml.SelectSingleNode("/r:customUI/r:ribbon/r:tabs/r:tab", namespaces);
        Check(tab != null && tab.Attributes["label"].Value == "Vefeyât", "Dedicated Vefeyât Word tab");
        var identifiers = new HashSet<string>();
        foreach (XmlElement element in xml.SelectNodes("//*[@id]"))
            Check(identifiers.Add(element.GetAttribute("id")), "Ribbon IDs must be unique");
        XmlNodeList buttons = xml.SelectNodes("//r:button", namespaces);
        Check(buttons.Count >= 2, "Search and browse commands are present");
        foreach (XmlElement button in buttons)
        {
            string callback = button.GetAttribute("onAction");
            MethodInfo method = typeof(Connect).GetMethod(callback, BindingFlags.Public | BindingFlags.Instance);
            Check(method != null, "Every ribbon callback resolves: " + callback);
            Check(method.ReturnType == typeof(void), "Button callback returns void: " + callback);
            ParameterInfo[] args = method.GetParameters();
            Check(args.Length == 1 && args[0].ParameterType == typeof(object), "Button callback accepts its COM control: " + callback);
        }
    }

    private static void Check(bool condition, string message)
    {
        assertions++;
        if (!condition) throw new Exception("FAIL: " + message);
    }

}
