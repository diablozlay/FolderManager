// ============================================================================
// 檔案管理工具 Folder Manager — 單檔啟動器（host）
// 功能：讀取內嵌資源 FileManager.ps1，以 STA Runspace 執行；
//       結束時聚合 PowerShell 錯誤流，寫入 logs\runtime_errors.log 並提示。
// 建置：csc.exe /target:winexe /resource:FileManager.ps1 Program.cs
//       （詳見 build.cmd）
// ============================================================================
using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

[assembly: AssemblyTitle("檔案管理工具 Folder Manager")]
[assembly: AssemblyDescription("ZW3D 檔案搜尋與清理工具（單檔可攜版）")]
[assembly: AssemblyProduct("檔案管理工具")]
[assembly: AssemblyVersion("1.5.2.0")]
[assembly: AssemblyFileVersion("1.5.2.0")]

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        string baseDir;
        try { baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); }
        catch { baseDir = Directory.GetCurrentDirectory(); }
        if (string.IsNullOrEmpty(baseDir)) baseDir = Directory.GetCurrentDirectory();
        Environment.SetEnvironmentVariable("FM_BASEDIR", baseDir);

        bool selfTest = false;
        bool perf = false;
        foreach (string a in args)
        {
            string t = (a ?? string.Empty).TrimStart('-', '/').ToLowerInvariant();
            if (t == "selftest") selfTest = true;
            if (t == "perf") { perf = true; selfTest = true; }
        }

        PowerShell ps = null;
        Runspace rs = null;
        try
        {
            string script = LoadScript();
            if (script == null)
            {
                Fatal(baseDir, "內嵌的腳本資源讀取失敗（資源可能損毀）。", selfTest);
                return 1;
            }

            InitialSessionState iss = InitialSessionState.CreateDefault();
            iss.ApartmentState = System.Threading.ApartmentState.STA;
            rs = RunspaceFactory.CreateRunspace(iss);
            rs.Open();

            ps = PowerShell.Create();
            ps.Runspace = rs;
            ps.AddScript(script);
            if (selfTest) ps.AddParameter("SelfTest", true);
            if (perf) ps.AddParameter("Perf", true);
            ps.Invoke();

            var errors = ps.Streams.Error;
            if (errors != null && errors.Count > 0)
            {
                var sb = new StringBuilder();
                foreach (ErrorRecord er in errors)
                    sb.AppendLine("[" + er.FullyQualifiedErrorId + "] " + er.Exception.Message);
                LogText(baseDir, "runtime_errors.log", sb.ToString());
                if (!selfTest)
                {
                    MessageBox.Show(
                        "執行期間發生錯誤，詳細內容已寫入 logs\\runtime_errors.log\n\n" + sb.ToString(),
                        "檔案管理工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            return 0;
        }
        catch (Exception ex)
        {
            Fatal(baseDir, ex.ToString(), selfTest);
            return 2;
        }
        finally
        {
            if (ps != null) ps.Dispose();
            if (rs != null) rs.Dispose();
        }
    }

    // 讀取內嵌的 .ps1 資源（找不到回傳 null）
    private static string LoadScript()
    {
        Assembly asm = Assembly.GetExecutingAssembly();
        string name = null;
        foreach (string n in asm.GetManifestResourceNames())
        {
            if (n.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase)) { name = n; break; }
        }
        if (name == null) return null;
        using (Stream s = asm.GetManifestResourceStream(name))
        {
            if (s == null) return null;
            using (var r = new StreamReader(s, Encoding.UTF8, true))
                return r.ReadToEnd();
        }
    }

    // 致命錯誤：寫 crash.log 並彈窗（彈窗失敗就靜默結束）
    private static void Fatal(string dir, string msg, bool quiet)
    {
        LogText(dir, "crash.log", msg);
        if (quiet) return;
        try
        {
            MessageBox.Show("程式無法啟動：\n\n" + msg + "\n\n記錄檔位置：\n" + dir,
                "檔案管理工具", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch { }
    }

    // logs\<file> 附加一行（UTF-8 BOM），失敗靜默
    private static void LogText(string dir, string file, string text)
    {
        try
        {
            string logDir = Path.Combine(dir, "logs");
            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
            File.AppendAllText(Path.Combine(logDir, file),
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + text + Environment.NewLine,
                new UTF8Encoding(true));
        }
        catch { }
    }
}
