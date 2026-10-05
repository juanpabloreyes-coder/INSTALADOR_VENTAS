// Instalador-Addins-VENTAS.exe: extrae el paquete (zip incrustado) a una carpeta temporal y corre instalar.cmd.
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;

class Programa
{
    static int Main()
    {
        string dir = Path.Combine(Path.GetTempPath(), "AddinsVENTAS_" + Guid.NewGuid().ToString("N").Substring(0, 8));
        try
        {
            Directory.CreateDirectory(dir);
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("paquete.zip"))
            using (ZipArchive z = new ZipArchive(s, ZipArchiveMode.Read))
                z.ExtractToDirectory(dir);
            ProcessStartInfo p = new ProcessStartInfo("cmd.exe", "/c \"" + Path.Combine(dir, "instalar.cmd") + "\"");
            p.WorkingDirectory = dir;
            p.UseShellExecute = false;
            Process proc = Process.Start(p);
            proc.WaitForExit();
            return proc.ExitCode;
        }
        catch (Exception ex)
        {
            Console.WriteLine("No se pudo instalar: " + ex.Message);
            Console.WriteLine("Avisale a Juan Pablo con una captura de esta ventana.");
            Console.ReadLine();
            return 1;
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }
}
