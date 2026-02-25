using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using iText.Kernel.Pdf;

class Program
{

  static string rootPath = "/home/desarrollo/Documentos/GESTOR_CTS/doc_pruebas";

  static void Main(string[] args)
  {
    Console.WriteLine("🚀 Iniciando proceso de escaneo...");

    if (!Directory.Exists(rootPath))
    {
      Console.WriteLine($"❌ Error: La ruta {rootPath} no existe.");
      return;
    }

    RecursiveWalk(rootPath);

    Console.WriteLine("\n✅ Proceso finalizado con éxito.");
  }

  static void RecursiveWalk(string currentPath)
  {
    try
    {      
      var pdfFiles = Directory.EnumerateFiles(currentPath, "*.*")
          .Where(f => f.ToLower().EndsWith(".pdf")).ToList();
      Console.WriteLine($"📂 {currentPath} ({pdfFiles.Count} archivos)");
      Console.WriteLine($"  {string.Join("\n  ", pdfFiles)}");
      if (pdfFiles.Any())
      {
        ExportToCsv(currentPath, pdfFiles);
      }

      
      foreach (string subDir in Directory.EnumerateDirectories(currentPath))
      {
        RecursiveWalk(subDir);
      }
    }
    catch (UnauthorizedAccessException) {}
    catch (Exception ex) { Console.WriteLine($"⚠️ Error en {currentPath}: {ex.Message}"); }
  }

  static void ExportToCsv(string folderPath, List<string> files)
  {
    string csvPath = Path.Combine(folderPath, "CARGA_SCRIPT.csv");
    string relativePath = Path.GetRelativePath(rootPath, folderPath);
    string[] levels = relativePath == "." ? new string[] { "Raiz" } : relativePath.Split(Path.DirectorySeparatorChar);

    using (var sw = new StreamWriter(csvPath, false, Encoding.UTF8))
    {
      
      StringBuilder header = new StringBuilder();
      for (int i = 0; i < levels.Length; i++) header.Append($"Nivel_{i + 1};");
      header.Append("Nombre_Archivo;Paginas");
      sw.WriteLine(header.ToString());

      foreach (var file in files)
      {
        int pgs = GetPages(file);
        string fileName = Path.GetFileName(file);
        string rowPath = string.Join(";", levels);

        sw.WriteLine($"{rowPath};{fileName};{pgs}");
      }
    }
    Console.WriteLine($"=====================[CARGA_SCRIPT] creado en: {folderPath} ({files.Count} archivos)");
  }

  static int GetPages(string path)
  {
    try
    {
      using (var reader = new PdfReader(path))
      using (var pdfDoc = new PdfDocument(reader))
      {
        return pdfDoc.GetNumberOfPages();
      }
    }
    catch { return 0; } 
  }
}