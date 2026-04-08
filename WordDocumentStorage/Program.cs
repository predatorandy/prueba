using Microsoft.EntityFrameworkCore;
using WordDocumentStorage.Models;

namespace WordDocumentStorage;

public class Program
{
    public static async Task Main(string[] args)
    {
        // Configurar el contexto de base de datos SQLite
        var optionsBuilder = new DbContextOptionsBuilder<ApplicationDbContext>();
        var connectionString = "Data Source=WordDocuments.db";
        
        optionsBuilder.UseSqlite(connectionString);
        
        using var context = new ApplicationDbContext(optionsBuilder.Options);
        
        // Crear la base de datos si no existe
        await context.Database.EnsureCreatedAsync();
        
        Console.WriteLine("Base de datos SQLite lista (WordDocuments.db).");
        Console.WriteLine("Ejemplo de uso del servicio:");
        Console.WriteLine("");
        Console.WriteLine("// Registrar el servicio en Program.cs (ASP.NET Core):");
        Console.WriteLine("builder.Services.AddScoped<IWordDocumentService, WordDocumentService>();");
        Console.WriteLine("");
        Console.WriteLine("// Configurar SQLite en Program.cs:");
        Console.WriteLine("builder.Services.AddDbContext<ApplicationDbContext>(options =>");
        Console.WriteLine("    options.UseSqlite(\"Data Source=WordDocuments.db\"));");
        Console.WriteLine("");
        Console.WriteLine("// Uso en un controlador o servicio:");
        Console.WriteLine("var document = await _wordDocumentService.ReadAndSaveDocumentAsync(@\"C:\\ruta\\archivo.docx\", \"Descripción opcional\");");
    }
}
