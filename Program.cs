using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Collections.Generic;
using ClosedXML.Excel;

class Program
{
  static string rootPath = @"/media/desarrollo/HD710 PRO/DISCO 3 V2 final/Nueva carpeta/OTROS";
  static string excelPath = @"/media/desarrollo/HD710 PRO/bases final final/DISCO 0003 MODIFICADO VA-1.xlsx";

  static Dictionary<string, List<string>> indicePdf =
      new(StringComparer.OrdinalIgnoreCase);

  static Dictionary<string, List<string>> pdfsPorCarpeta =
      new(StringComparer.OrdinalIgnoreCase);

  static Dictionary<string, List<Dictionary<string, string>>> cargas =
      new(StringComparer.OrdinalIgnoreCase);

  static HashSet<string> pdfsUsados = new();

  static void Main()
  {
    Console.WriteLine("🚀 Iniciando proceso\n");

    if (!Directory.Exists(rootPath))
    {
      Console.WriteLine("❌ No existe la ruta raíz");
      return;
    }

    if (!File.Exists(excelPath))
    {
      Console.WriteLine("❌ No existe el Excel");
      return;
    }

    IndexarPDFs();

    using (var workbook = new XLWorkbook(excelPath))
    {
      var hoja = workbook.Worksheets.FirstOrDefault();

      if (hoja == null)
      {
        Console.WriteLine("❌ El Excel no tiene hojas");
        return;
      }

      ProcessSheet(workbook, hoja.Name);

      Console.WriteLine("\n💾 Guardando cambios en el Excel maestro...");
      workbook.Save();
    }

    CrearCargaScripts();
    CrearCargaScriptVerificar();
    CrearResumenGlobal();

    Console.WriteLine("\n✅ Proceso terminado");
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

      string carpeta = NormalizarRuta(Path.GetDirectoryName(file));

      if (!pdfsPorCarpeta.ContainsKey(carpeta))
        pdfsPorCarpeta[carpeta] = new List<string>();

      pdfsPorCarpeta[carpeta].Add(file);
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

    var usedRange = ws.RangeUsed();

    if (usedRange == null)
    {
      Console.WriteLine("⚠ Hoja vacía");
      return;
    }

    var encabezados = ws.Row(1)
        .CellsUsed()
        .Select(c => c.GetString())
        .Where(c => !string.IsNullOrWhiteSpace(c))
        .ToList();

    int colRuta = encabezados.Count + 1;
    ws.Cell(1, colRuta).Value = "RUTA_PDF";

    var rows = usedRange.RowsUsed().Skip(1).ToList();

    int total = rows.Count;
    int contador = 0;

    foreach (var row in rows)
    {
      contador++;

      string nombreExcel = row.Cell(1).GetString();

      if (string.IsNullOrWhiteSpace(nombreExcel))
        continue;

      string nombre = Limpiar(nombreExcel);

      Console.WriteLine($"🔎 [{contador}/{total}] {nombreExcel}");

      // 🚫 Ignorar nombres muy cortos (evita basura tipo "15")
      if (nombre.Length < 5)
      {
        Console.WriteLine("   ⚠ Ignorado (muy corto)");
        continue;
      }

      var palabras = nombre.Split(' ', StringSplitOptions.RemoveEmptyEntries);

      var coincidencias = indicePdf
          .Where(kvp => palabras.All(p => kvp.Key.Contains(p)))
          .SelectMany(kvp => kvp.Value)
          .ToList();

      if (coincidencias.Any())
      {
        // ✅ Elegir SOLO 1 (evita error de 32k caracteres)
        var rutaElegida = coincidencias
            .OrderBy(r => Math.Abs(
                Path.GetFileNameWithoutExtension(r).Length - nombre.Length))
            .First();

        row.Cell(colRuta).Value = rutaElegida;

        pdfsUsados.Add(rutaElegida);

        string carpeta = NormalizarRuta(Path.GetDirectoryName(rutaElegida));

        var registro = new Dictionary<string, string>();

        for (int i = 0; i < encabezados.Count; i++)
          registro[encabezados[i]] = row.Cell(i + 1).GetString();

        registro["RUTA_PDF"] = rutaElegida;

        if (!cargas.ContainsKey(carpeta))
          cargas[carpeta] = new List<Dictionary<string, string>>();

        cargas[carpeta].Add(registro);

        Console.WriteLine("   ✅ Encontrado");
      }
      else
      {
        row.Cell(colRuta).Value = "";
        Console.WriteLine("   ❌ No encontrado");
      }
    }

    Console.WriteLine($"✔ Hoja {sheetName} terminada\n");
  }

  static void CrearCargaScripts()
  {
    Console.WriteLine("\n📄 Creando CARGA_SCRIPT.xlsx...");

    foreach (var carpeta in cargas.Keys)
    {
      Console.WriteLine($"📂 Carpeta destino: {carpeta}");

      var registros = cargas[carpeta];

      string path = Path.Combine(carpeta, "CARGA_SCRIPT.xlsx");

      if (File.Exists(path))
        File.Delete(path);

      string nombreHoja = LimpiarNombreHoja(new DirectoryInfo(carpeta).Name);

      using (var wb = new XLWorkbook())
      {
        var ws = wb.AddWorksheet(nombreHoja);

        var headers = registros
            .First()
            .Keys
            .Where(h => h != "RUTA_PDF")
            .ToList();

        for (int i = 0; i < headers.Count; i++)
          ws.Cell(1, i + 1).Value = headers[i];

        int fila = 2;

        foreach (var reg in registros)
        {
          int col = 1;

          foreach (var h in headers)
          {
            ws.Cell(fila, col).Value =
                reg.ContainsKey(h) && !string.IsNullOrWhiteSpace(reg[h])
                ? reg[h]
                : "N/A";

            col++;
          }

          fila++;
        }

        wb.SaveAs(path);
      }

      Console.WriteLine($"✔ Creado: {path}");
    }
  }

  static void CrearCargaScriptVerificar()
  {
    Console.WriteLine("\n🔎 Creando CARGA_SCRIPT_VERIFICAR...");

    foreach (var carpeta in pdfsPorCarpeta.Keys)
    {
      var todos = pdfsPorCarpeta[carpeta];

      var usados = new HashSet<string>();

      if (cargas.ContainsKey(carpeta))
        usados = cargas[carpeta].Select(r => r["RUTA_PDF"]).ToHashSet();

      var faltantes = todos
          .Where(p => !usados.Contains(p))
          .ToList();

      if (!faltantes.Any())
        continue;

      string path = Path.Combine(carpeta, "CARGA_SCRIPT_VERIFICAR.xlsx");

      if (File.Exists(path))
        File.Delete(path);

      using (var wb = new XLWorkbook())
      {
        var ws = wb.AddWorksheet("VERIFICAR");

        ws.Cell(1, 1).Value = "ARCHIVO";
        ws.Cell(1, 2).Value = "Ruta";

        int fila = 2;

        foreach (var f in faltantes)
        {
          ws.Cell(fila, 1).Value = Path.GetFileName(f);
          ws.Cell(fila, 2).Value = f;
          fila++;
        }

        wb.SaveAs(path);
      }

      Console.WriteLine($"⚠ Creado verificar: {path}");
    }
  }

  static void CrearResumenGlobal()
  {
    Console.WriteLine("\n📊 Generando RESUMEN_GLOBAL.xlsx");

    string path = Path.Combine(rootPath, "RESUMEN_GLOBAL.xlsx");

    if (File.Exists(path))
      File.Delete(path);

    int totalDisco = indicePdf.Values.Sum(x => x.Count);
    int totalUsados = pdfsUsados.Count;
    int totalFaltantes = totalDisco - totalUsados;

    using (var wb = new XLWorkbook())
    {
      var ws = wb.AddWorksheet("RESUMEN");

      ws.Cell(1, 1).Value = "METRICA";
      ws.Cell(1, 2).Value = "VALOR";

      ws.Cell(2, 1).Value = "PDFs en disco";
      ws.Cell(2, 2).Value = totalDisco;

      ws.Cell(3, 1).Value = "PDFs encontrados en Excel";
      ws.Cell(3, 2).Value = totalUsados;

      ws.Cell(4, 1).Value = "PDFs no referenciados en Excel";
      ws.Cell(4, 2).Value = totalFaltantes;

      ws.Cell(5, 1).Value = "Carpetas con CARGA_SCRIPT";
      ws.Cell(5, 2).Value = cargas.Count;

      wb.SaveAs(path);
    }

    Console.WriteLine($"✔ RESUMEN creado: {path}");
  }

  static string NormalizarRuta(string ruta)
  {
    if (string.IsNullOrWhiteSpace(ruta))
      return "";

    return Path.GetFullPath(ruta).Trim();
  }

  static string Limpiar(string texto)
  {
    if (string.IsNullOrWhiteSpace(texto))
      return "";

    texto = texto.Replace('\u00A0', ' ').Trim();

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

  static string LimpiarNombreHoja(string nombre)
  {
    if (string.IsNullOrWhiteSpace(nombre))
      return "HOJA";

    char[] invalidos = { '\\', '/', '?', '*', '[', ']', ':' };

    foreach (var c in invalidos)
      nombre = nombre.Replace(c.ToString(), "");

    nombre = nombre.Trim();

    if (nombre.Length > 31)
      nombre = nombre.Substring(0, 31);

    return string.IsNullOrWhiteSpace(nombre) ? "HOJA" : nombre;
  }
}