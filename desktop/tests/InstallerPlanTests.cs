using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Vefeyat.Setup;

// Pure registration-plan checks. No Windows registry, Word process, certificate
// store, user catalog, or installation directory is opened by this application.
internal static class InstallerPlanTests
{
    private const string ProgId = "Vefeyat.Connect";
    private const string ClassId = "{2B0AF7B7-BD0D-4C70-9136-BED2DBD8F8F3}";
    private static int assertions;

    public static int Main()
    {
        try
        {
            string installation = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "Vefeyat plan with spaces"));
            string library = Path.Combine(installation, "Vefeyat.WordAddin.dll");
            string installer = Path.Combine(installation, "Vefeyat.Kurulum.exe");
            AssemblyName assembly = new AssemblyName("Vefeyat.WordAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null");
            Dictionary<string, Dictionary<string, object>> plan = InstallerEngine.RegistrationPlan(assembly, library, installer, "Test Publisher");
            ComRegistration(plan, assembly, library);
            UninstallRegistration(plan, installer);
            ScopeAndDataPreservation(plan);
            RejectWrongPayloadIdentity(library, installer);
            Console.WriteLine("Installer registration-plan checks passed: " + assertions);
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static void ComRegistration(Dictionary<string, Dictionary<string, object>> plan, AssemblyName assembly, string library)
    {
        string clsid = "Software\\Classes\\CLSID\\" + ClassId;
        Check(Text(plan, clsid, "") == ProgId, "COM class identity must match Word's exposed class");
        Check(Text(plan, clsid + "\\ProgId", "") == ProgId, "CLSID must map back to the registered ProgID");
        Check(Text(plan, "Software\\Classes\\" + ProgId + "\\CLSID", "") == ClassId, "ProgID must map to the correct CLSID");

        foreach (string path in new[] { clsid + "\\InprocServer32", clsid + "\\InprocServer32\\" + assembly.Version })
        {
            Check(Text(plan, path, "") == "mscoree.dll", "Managed COM must activate through the CLR shim");
            Check(Text(plan, path, "Class") == ProgId, "Managed class must match the eklenti COM class");
            Check(Text(plan, path, "Assembly") == assembly.FullName, "COM must load the exact assembly identity");
            Check(Text(plan, path, "RuntimeVersion") == "v4.0.30319", ".NET Framework 4 CLR must be selected");
            Check(Text(plan, path, "ThreadingModel") == "Both", "Both Word bitness registrations use the managed threading model");
            string codeBase = Text(plan, path, "CodeBase");
            Uri location;
            Check(Uri.TryCreate(codeBase, UriKind.Absolute, out location) && location.IsFile, "CodeBase must be an absolute file URI");
            Check(codeBase.IndexOf("%20", StringComparison.Ordinal) >= 0, "CodeBase must escape spaces in the user's installation path");
            Check(String.Equals(location.LocalPath, library, StringComparison.OrdinalIgnoreCase), "CodeBase must point at the installed DLL");
        }

        string addin = "Software\\Microsoft\\Office\\Word\\Addins\\" + ProgId;
        Check(Text(plan, addin, "FriendlyName") == "Vefeyât", "Word must display the expected add-in name");
        object load = Value(plan, addin, "LoadBehavior");
        Check(load is int && (int)load == 3, "LoadBehavior must be DWORD 3 for startup activation");
        object commandLine = Value(plan, addin, "CommandLineSafe");
        Check(commandLine is int && (int)commandLine == 0, "CommandLineSafe must remain a numeric value");
    }

    private static void UninstallRegistration(Dictionary<string, Dictionary<string, object>> plan, string installer)
    {
        string path = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Vefeyat";
        Check(Text(plan, path, "DisplayName").StartsWith("Vefeyât", StringComparison.Ordinal), "Windows installed-apps list must identify the application");
        Check(Text(plan, path, "Publisher") == "Test Publisher", "Installed-apps publisher must come from the verified publisher");
        Check(Text(plan, path, "UninstallString") == "\"" + installer + "\" --uninstall", "Uninstall executable path must be quoted and select the uninstall UI");
        string command = Text(plan, path, "UninstallString");
        Check(command.IndexOf("cmd.exe", StringComparison.OrdinalIgnoreCase) < 0 && command.IndexOf("powershell", StringComparison.OrdinalIgnoreCase) < 0,
            "Uninstall must directly launch the installer instead of a shell script");
    }

    private static void RejectWrongPayloadIdentity(string library, string installer)
    {
        Reject(delegate { InstallerEngine.RegistrationPlan(null, library, installer, "Test Publisher"); }, "A missing assembly identity must not be installed");
        Reject(delegate { InstallerEngine.RegistrationPlan(new AssemblyName("Unrelated.Library, Version=1.0.0.0"), library, installer, "Test Publisher"); },
            "An unrelated signed DLL must not be registered as Vefeyât");
    }

    private static void Reject(Action action, string message)
    {
        try { action(); }
        catch (InvalidDataException) { assertions++; return; }
        throw new Exception("FAIL: " + message);
    }

    private static void ScopeAndDataPreservation(Dictionary<string, Dictionary<string, object>> plan)
    {
        string classRoot = "Software\\Classes\\CLSID\\" + ClassId;
        string progRoot = "Software\\Classes\\" + ProgId;
        string wordRoot = "Software\\Microsoft\\Office\\Word\\Addins\\" + ProgId;
        string uninstallRoot = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Vefeyat";
        Check(plan.Count >= 7, "Plan must contain COM activation, Word loading and Windows uninstall registration");
        foreach (KeyValuePair<string, Dictionary<string, object>> entry in plan)
        {
            string key = entry.Key;
            bool owned = IsWithin(key, classRoot) || IsWithin(key, progRoot) || IsWithin(key, wordRoot) || IsWithin(key, uninstallRoot);
            Check(owned, "Every planned key must belong to this application's registration: " + key);
            Check(key.IndexOf("HKEY_LOCAL_MACHINE", StringComparison.OrdinalIgnoreCase) < 0 && key.IndexOf("HKLM", StringComparison.OrdinalIgnoreCase) < 0,
                "No machine hive may be encoded in the per-user registration plan");
            Check(key.IndexOf("Policies", StringComparison.OrdinalIgnoreCase) < 0 && key.IndexOf("CodeIntegrity", StringComparison.OrdinalIgnoreCase) < 0 &&
                key.IndexOf("Certificates", StringComparison.OrdinalIgnoreCase) < 0, "Registration must not alter security policy or certificate trust");
            foreach (object value in entry.Value.Values)
            {
                string text = value as string;
                Check(text == null || (text.IndexOf("catalog.xml", StringComparison.OrdinalIgnoreCase) < 0 &&
                    text.IndexOf("Vefeyat\\Data", StringComparison.OrdinalIgnoreCase) < 0), "Registration must leave the user's catalog outside its mutation scope");
            }
        }
    }

    private static bool IsWithin(string path, string root)
    {
        return String.Equals(path, root, StringComparison.OrdinalIgnoreCase) || path.StartsWith(root + "\\", StringComparison.OrdinalIgnoreCase);
    }

    private static object Value(Dictionary<string, Dictionary<string, object>> plan, string path, string name)
    {
        Dictionary<string, object> values;
        object result;
        if (!plan.TryGetValue(path, out values) || !values.TryGetValue(name, out result))
            throw new Exception("FAIL: missing registry value " + path + " [" + name + "]");
        return result;
    }

    private static string Text(Dictionary<string, Dictionary<string, object>> plan, string path, string name)
    {
        object value = Value(plan, path, name);
        Check(value is string, "Registry string value must stay a string: " + path + " [" + name + "]");
        return value as string;
    }

    private static void Check(bool condition, string message)
    {
        assertions++;
        if (!condition) throw new Exception("FAIL: " + message);
    }
}
