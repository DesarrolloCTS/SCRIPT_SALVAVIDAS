// using System;
// using System.IO;
// using System.Linq;
// using System.Text;
// using System.Collections.Generic;
// using iText.Kernel.Pdf;

// class Program
// {

//   static string rootPath = "/media/desarrollo/HD710 PRO/DISCO 2 V2 final";

//   static void Main(string[] args)
//   {
//     Console.WriteLine("🚀 Iniciando proceso de escaneo...");

//     if (!Directory.Exists(rootPath))
//     {
//       Console.WriteLine($"❌ Error: La ruta {rootPath} no existe.");
//       return;
//     }
//     RecursiveWalk(rootPath);
//     Console.WriteLine("\n✅ Proceso finalizado con éxito.");
//   }

//   static void RecursiveWalk(string currentPath)
//   {
//     try
//     {      
//       var pdfFiles = Directory.EnumerateFiles(currentPath, "*.*")
//           .Where(f => f.ToLower().EndsWith(".pdf")).ToList();
//       Console.WriteLine($"📂 {currentPath} ({pdfFiles.Count} archivos)");
//       Console.WriteLine($"  {string.Join("\n  ", pdfFiles)}");
//       if (pdfFiles.Any())
//       {
//         ExportToCsv(currentPath, pdfFiles);
//       }


//       foreach (string subDir in Directory.EnumerateDirectories(currentPath))
//       {
//         RecursiveWalk(subDir);
//       }
//     }
//     catch (UnauthorizedAccessException) {}
//     catch (Exception ex) { Console.WriteLine($"⚠️ Error en {currentPath}: {ex.Message}"); }
//   }

//   static void ExportToCsv(string folderPath, List<string> files)
//   {
//     string csvPath = Path.Combine(folderPath, "CARGA_SCRIPT.csv");
//     string relativePath = Path.GetRelativePath(rootPath, folderPath);
//     string[] levels = relativePath == "." ? new string[] { "Raiz" } : relativePath.Split(Path.DirectorySeparatorChar);

//     using (var sw = new StreamWriter(csvPath, false, Encoding.UTF8))
//     {

//       StringBuilder header = new StringBuilder();
//       for (int i = 0; i < levels.Length; i++) header.Append($"Nivel_{i + 1};");
//       header.Append("Nombre_Archivo;Paginas");
//       sw.WriteLine(header.ToString());

//       foreach (var file in files)
//       {
//         int pgs = GetPages(file);
//         string fileName = Path.GetFileName(file);
//         string rowPath = string.Join(";", levels);

//         sw.WriteLine($"{rowPath};{fileName};{pgs}");
//       }
//     }
//     Console.WriteLine($"=====================[CARGA_SCRIPT] creado en: {folderPath} ({files.Count} archivos)");
//   }

//   static int GetPages(string path)
//   {
//     try
//     {
//       using (var reader = new PdfReader(path))
//       using (var pdfDoc = new PdfDocument(reader))
//       {
//         return pdfDoc.GetNumberOfPages();
//       }
//     }
//     catch { return 0; } 
//   }
// }
// using System;
// using System.IO;
// using System.Linq;
// using System.Text;
// using System.Collections.Generic;
// using iText.Kernel.Pdf;

// class Program
// {

//   static string rootPath = "/media/desarrollo/HD710 PRO/DISCO 2 V2 final";

//   static void Main(string[] args)
//   {
//     Console.WriteLine("🚀 Iniciando proceso de escaneo...");

//     if (!Directory.Exists(rootPath))
//     {
//       Console.WriteLine($"❌ Error: La ruta {rootPath} no existe.");
//       return;
//     }
//     RecursiveWalk(rootPath);
//     Console.WriteLine("\n✅ Proceso finalizado con éxito.");
//   }

//   static void RecursiveWalk(string currentPath)
//   {
//     try
//     {      
//       var pdfFiles = Directory.EnumerateFiles(currentPath, "*.*")
//           .Where(f => f.ToLower().EndsWith(".pdf")).ToList();
//       Console.WriteLine($"📂 {currentPath} ({pdfFiles.Count} archivos)");
//       Console.WriteLine($"  {string.Join("\n  ", pdfFiles)}");
//       if (pdfFiles.Any())
//       {
//         ExportToCsv(currentPath, pdfFiles);
//       }


//       foreach (string subDir in Directory.EnumerateDirectories(currentPath))
//       {
//         RecursiveWalk(subDir);
//       }
//     }
//     catch (UnauthorizedAccessException) {}
//     catch (Exception ex) { Console.WriteLine($"⚠️ Error en {currentPath}: {ex.Message}"); }
//   }

//   static void ExportToCsv(string folderPath, List<string> files)
//   {
//     string csvPath = Path.Combine(folderPath, "CARGA_SCRIPT.csv");
//     string relativePath = Path.GetRelativePath(rootPath, folderPath);
//     string[] levels = relativePath == "." ? new string[] { "Raiz" } : relativePath.Split(Path.DirectorySeparatorChar);

//     using (var sw = new StreamWriter(csvPath, false, Encoding.UTF8))
//     {

//       StringBuilder header = new StringBuilder();
//       for (int i = 0; i < levels.Length; i++) header.Append($"Nivel_{i + 1};");
//       header.Append("Nombre_Archivo;Paginas");
//       sw.WriteLine(header.ToString());

//       foreach (var file in files)
//       {
//         int pgs = GetPages(file);
//         string fileName = Path.GetFileName(file);
//         string rowPath = string.Join(";", levels);

//         sw.WriteLine($"{rowPath};{fileName};{pgs}");
//       }
//     }
//     Console.WriteLine($"=====================[CARGA_SCRIPT] creado en: {folderPath} ({files.Count} archivos)");
//   }

//   static int GetPages(string path)
//   {
//     try
//     {
//       using (var reader = new PdfReader(path))
//       using (var pdfDoc = new PdfDocument(reader))
//       {
//         return pdfDoc.GetNumberOfPages();
//       }
//     }
//     catch { return 0; } 
//   }
// }
using System;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using ClosedXML.Excel;

class Program
{
  static string rootPath = @"/media/desarrollo/HD710 PRO/DISCO 2 V2 final";
  static string excelPath = @"/media/desarrollo/HD710 PRO/bases final final/DISCO 0002 FINAL Val-2 14-1-2020.xlsx";

  static Dictionary<string, List<string>> cargas = new();
  static Dictionary<string, List<string>> indicePdf = new();

  static void Main()
  {
    Console.WriteLine("🚀 Iniciando proceso...\n");

    IndexarPDFs();

    using (var workbook = new XLWorkbook(excelPath))
    {
      ProcessSheet(workbook, "INSURGENTES");
      ProcessSheet(workbook, "TLAHUAC");

      Console.WriteLine("\n💾 Guardando Excel...");
      workbook.Save();
    }

    CrearCargaScripts();

    Console.WriteLine("\n✅ Proceso terminado.");
  }

  static void IndexarPDFs()
  {
    Console.WriteLine("📂 Indexando PDFs...");

    var archivos = Directory
        .EnumerateFiles(rootPath, "*", SearchOption.AllDirectories)
        .Where(f => Path.GetExtension(f)
        .Equals(".pdf", StringComparison.OrdinalIgnoreCase));

    foreach (var file in archivos)
    {
      string nombre = Limpiar(Path.GetFileNameWithoutExtension(file));

      if (!indicePdf.ContainsKey(nombre))
        indicePdf[nombre] = new List<string>();

      indicePdf[nombre].Add(file);
    }

    Console.WriteLine($"📊 PDFs indexados: {indicePdf.Count}\n");
  }

  static void ProcessSheet(XLWorkbook workbook, string sheetName)
  {
    Console.WriteLine($"📑 Procesando hoja: {sheetName}");

    var ws = workbook.Worksheets.FirstOrDefault(w => w.Name == sheetName);

    if (ws == null)
    {
      Console.WriteLine("❌ Hoja no encontrada");
      return;
    }

    ws.Cell("K1").Value = "RUTA_PDF";

    var rows = ws.RangeUsed().RowsUsed().Skip(1).ToList();

    Console.WriteLine($"📊 Registros: {rows.Count}");

    int contador = 0;

    foreach (var row in rows)
    {
      contador++;

      string nombreExcel = row.Cell("A").GetString();

      if (string.IsNullOrWhiteSpace(nombreExcel))
        continue;

      string nombre = Limpiar(nombreExcel);

      Console.WriteLine($"🔎 [{contador}/{rows.Count}] {nombre}");

      if (indicePdf.ContainsKey(nombre))
      {
        var rutas = indicePdf[nombre];

        row.Cell("K").Value = string.Join(",", rutas);

        foreach (var ruta in rutas)
        {
          string carpeta = Path.GetDirectoryName(ruta);

          string jefatura = row.Cell("B").GetString();
          string area = row.Cell("C").GetString();
          string anio = row.Cell("D").GetString();
          string archivo = row.Cell("A").GetString();

          string registro = $"{jefatura};{area};{anio};{archivo};{ruta}";

          if (!cargas.ContainsKey(carpeta))
            cargas[carpeta] = new List<string>();

          cargas[carpeta].Add(registro);
        }

        Console.WriteLine("   ✅ Encontrado");
      }
      else
      {
        row.Cell("K").Value = "";
        Console.WriteLine("   ❌ No encontrado");
      }
    }

    Console.WriteLine($"✔ Hoja {sheetName} terminada\n");
  }

  static void CrearCargaScripts()
  {
    Console.WriteLine("\n📄 Creando CARGA_SCRIPT...");

    foreach (var carpeta in cargas.Keys)
    {
      string csvPath = Path.Combine(carpeta, "CARGA_SCRIPT.csv");

      using (var sw = new StreamWriter(csvPath))
      {
        sw.WriteLine("JEFATURA;AREA;ANIO;NOMBRE_ARCHIVO;RUTA");

        foreach (var linea in cargas[carpeta])
          sw.WriteLine(linea);
      }

      Console.WriteLine($"✔ {csvPath}");
    }
  }

  static string Limpiar(string texto)
  {
    if (string.IsNullOrWhiteSpace(texto))
      return "";

    texto = texto
        .Replace('\u00A0', ' ')
        .Trim();

    while (texto.Contains("  "))
      texto = texto.Replace("  ", " ");

    texto = QuitarAcentos(texto);

    return texto.ToUpper();
  }

  static string QuitarAcentos(string texto)
  {
    var normalized = texto.Normalize(NormalizationForm.FormD);
    var sb = new StringBuilder();

    foreach (var c in normalized)
    {
      if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
        sb.Append(c);
    }

    return sb.ToString().Normalize(NormalizationForm.FormC);
  }
}