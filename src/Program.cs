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

[assembly: AssemblyTitle("ZW3D 無用檔案清除程式 (ZW3D Junk File Cleaner)")]
[assembly: AssemblyDescription("ZW3D 檔案搜尋與清理工具（單檔可攜版）")]
[assembly: AssemblyProduct("ZW3D 無用檔案清除程式")]
[assembly: AssemblyVersion("1.6.0.0")]
[assembly: AssemblyFileVersion("1.6.0.0")]

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
                    string[] t = HostText(baseDir);
                    MessageBox.Show(
                        t[1] + "\n\n" + sb.ToString(),
                        t[0], MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

    // 啟動器彈窗文案：依 exe 同目錄 language.txt 的 ui= 介面語言（與腳本相同設定），預設繁中。
    // 順序：{ 視窗標題, 執行期錯誤說明, 無法啟動, 記錄檔位置 }
    private static readonly System.Collections.Generic.Dictionary<string, string[]> HostStrings =
        new System.Collections.Generic.Dictionary<string, string[]>
    {
        { "zh_TW", new[] { "ZW3D 無用檔案清除程式", "執行期間發生錯誤，詳細內容已寫入 logs\\runtime_errors.log", "程式無法啟動：", "記錄檔位置：" } },
        { "zh_CN", new[] { "ZW3D 无用文件清除程序", "运行期间发生错误，详细内容已写入 logs\\runtime_errors.log", "程序无法启动：", "日志位置：" } },
        { "en_US", new[] { "ZW3D Junk File Cleaner", "An error occurred while running. Details were written to logs\\runtime_errors.log", "The application could not start:", "Log location:" } },
        { "de_DE", new[] { "ZW3D Bereinigung unnötiger Dateien", "Während der Ausführung ist ein Fehler aufgetreten. Details stehen in logs\\runtime_errors.log", "Das Programm konnte nicht gestartet werden:", "Speicherort des Protokolls:" } },
        { "fr_FR", new[] { "ZW3D Nettoyeur de fichiers inutiles", "Une erreur s'est produite pendant l'exécution. Détails dans logs\\runtime_errors.log", "Impossible de démarrer l'application :", "Emplacement du journal :" } },
        { "es_ES", new[] { "ZW3D Limpiador de archivos innecesarios", "Se produjo un error durante la ejecución. Detalles en logs\\runtime_errors.log", "No se pudo iniciar la aplicación:", "Ubicación del registro:" } },
        { "it_IT", new[] { "ZW3D Pulizia file inutili", "Si è verificato un errore durante l'esecuzione. Dettagli in logs\\runtime_errors.log", "Impossibile avviare l'applicazione:", "Posizione del log:" } },
        { "pt_PT", new[] { "ZW3D Limpeza de ficheiros inúteis", "Ocorreu um erro durante a execução. Detalhes em logs\\runtime_errors.log", "Não foi possível iniciar a aplicação:", "Localização do registo:" } },
        { "ja_JP", new[] { "ZW3D 不要ファイル削除ツール", "実行中にエラーが発生しました。詳細は logs\\runtime_errors.log に記録されています", "プログラムを起動できません：", "ログの場所：" } },
        { "ko_KR", new[] { "ZW3D 불필요한 파일 정리 도구", "실행 중 오류가 발생했습니다. 자세한 내용은 logs\\runtime_errors.log에 기록되었습니다", "프로그램을 시작할 수 없습니다:", "로그 위치:" } },
        { "ru_RU", new[] { "ZW3D Очистка ненужных файлов", "Во время работы произошла ошибка. Подробности записаны в logs\\runtime_errors.log", "Не удалось запустить программу:", "Расположение журнала:" } },
        { "pl_PL", new[] { "ZW3D Czyszczenie zbędnych plików", "Podczas działania wystąpił błąd. Szczegóły zapisano w logs\\runtime_errors.log", "Nie można uruchomić programu:", "Lokalizacja dziennika:" } },
        { "cs_CZ", new[] { "ZW3D Čištění nepotřebných souborů", "Během běhu došlo k chybě. Podrobnosti jsou v logs\\runtime_errors.log", "Program nelze spustit:", "Umístění protokolu:" } },
        { "hu_HU", new[] { "ZW3D Felesleges fájlok törlése", "Futás közben hiba történt. A részletek: logs\\runtime_errors.log", "A program nem indítható:", "Napló helye:" } },
        { "tr_TR", new[] { "ZW3D Gereksiz Dosya Temizleyici", "Çalışma sırasında bir hata oluştu. Ayrıntılar logs\\runtime_errors.log dosyasına yazıldı", "Program başlatılamadı:", "Günlük konumu:" } },
    };

    private static string[] HostText(string dir)
    {
        string id = "zh_TW";
        try
        {
            string f = Path.Combine(dir ?? string.Empty, "language.txt");
            if (File.Exists(f))
            {
                foreach (string raw in File.ReadAllLines(f))
                {
                    string ln = (raw ?? string.Empty).Trim();
                    if (ln.StartsWith("ui=")) { id = ln.Substring(3).Trim(); break; }
                }
            }
        }
        catch { }
        string[] t;
        return HostStrings.TryGetValue(id, out t) ? t : HostStrings["zh_TW"];
    }

    // 致命錯誤：寫 crash.log 並彈窗（彈窗失敗就靜默結束）
    private static void Fatal(string dir, string msg, bool quiet)
    {
        LogText(dir, "crash.log", msg);
        if (quiet) return;
        try
        {
            string[] t = HostText(dir);
            MessageBox.Show(t[2] + "\n\n" + msg + "\n\n" + t[3] + "\n" + dir,
                t[0], MessageBoxButtons.OK, MessageBoxIcon.Error);
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
