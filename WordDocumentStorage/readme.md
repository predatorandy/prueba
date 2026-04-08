# Servicio de Almacenamiento de Documentos Word en .NET 10

Este proyecto proporciona un servicio completo para leer documentos Word (.docx) y guardarlos en una base de datos **SQLite** utilizando .NET 10.

## Estructura del Proyecto

```
WordDocumentStorage/
├── Models/
│   └── ApplicationDbContext.cs    # Modelo de datos y contexto EF Core
├── Services/
│   └── WordDocumentService.cs     # Servicio principal para manejar documentos Word
├── Program.cs                      # Punto de entrada y ejemplo de uso
└── WordDocumentStorage.csproj      # Archivo de proyecto
```

## Paquetes NuGet Requeridos

- **DocumentFormat.OpenXml** (v3.1.1): Para leer y validar documentos Word
- **Microsoft.EntityFrameworkCore.Sqlite** (v9.0.0): Para la conexión a SQLite
- **Microsoft.EntityFrameworkCore.Tools** (v9.0.0): Para herramientas de EF Core

## Características

El servicio `IWordDocumentService` proporciona las siguientes funcionalidades:

### Métodos Principales

1. **ReadAndSaveDocumentAsync(string filePath, string? description)**
   - Lee un documento Word desde una ruta de archivo
   - Lo valida como documento Word válido
   - Lo guarda en la base de datos con metadatos

2. **ReadAndSaveDocumentAsync(Stream documentStream, string fileName, string? description)**
   - Lee un documento Word desde un stream
   - Útil para uploads web o archivos en memoria

3. **GetDocumentByIdAsync(int id)**
   - Recupera un documento almacenado por su ID

4. **GetAllDocumentsAsync()**
   - Obtiene todos los documentos almacenados ordenados por fecha

5. **DeleteDocumentAsync(int id)**
   - Elimina un documento de la base de datos

6. **ExtractTextFromWord(string filePath)**
   - Extrae el contenido de texto plano de un documento Word

## Modelo de Datos

La entidad `StoredDocument` contiene:

| Propiedad | Tipo | Descripción |
|-----------|------|-------------|
| Id | int | Clave primaria autoincremental |
| FileName | string | Nombre original del archivo |
| ContentType | string | Tipo MIME del documento |
| FileContent | byte[] | Contenido binario del archivo |
| FileSize | long | Tamaño del archivo en bytes |
| CreatedDate | DateTime | Fecha de creación/almacenamiento |
| Description | string? | Descripción opcional del documento |

## Configuración en ASP.NET Core

### 1. Registrar el DbContext

```csharp
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlite(builder.Configuration.GetConnectionString("DefaultConnection")));
```

### 2. Registrar el Servicio

```csharp
builder.Services.AddScoped<IWordDocumentService, WordDocumentService>();
```

### 3. Connection String (appsettings.json)

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Data Source=WordDocuments.db"
  }
}
```

## Ejemplo de Uso en un Controlador

```csharp
[ApiController]
[Route("api/[controller]")]
public class DocumentsController : ControllerBase
{
    private readonly IWordDocumentService _documentService;

    public DocumentsController(IWordDocumentService documentService)
    {
        _documentService = documentService;
    }

    [HttpPost("upload")]
    public async Task<IActionResult> UploadDocument(IFormFile file, string? description)
    {
        if (file == null || file.Length == 0)
            return BadRequest("No se proporcionó ningún archivo");

        await using var stream = file.OpenReadStream();
        var document = await _documentService.ReadAndSaveDocumentAsync(
            stream, 
            file.FileName, 
            description);

        return Ok(new { document.Id, document.FileName, document.FileSize });
    }

    [HttpGet("{id}")]
    public async Task<IActionResult> GetDocument(int id)
    {
        var document = await _documentService.GetDocumentByIdAsync(id);
        
        if (document == null)
            return NotFound();

        return File(document.FileContent, document.ContentType, document.FileName);
    }

    [HttpGet]
    public async Task<IActionResult> GetAllDocuments()
    {
        var documents = await _documentService.GetAllDocumentsAsync();
        return Ok(documents.Select(d => new 
        { 
            d.Id, 
            d.FileName, 
            d.FileSize, 
            d.CreatedDate,
            d.Description 
        }));
    }

    [HttpDelete("{id}")]
    public async Task<IActionResult> DeleteDocument(int id)
    {
        var result = await _documentService.DeleteDocumentAsync(id);
        
        if (!result)
            return NotFound();

        return NoContent();
    }
}
```

## Ejemplo de Uso Directo

```csharp
// Guardar un documento desde una ruta
var document = await _wordDocumentService.ReadAndSaveDocumentAsync(
    @"C:\Documentos\informe.docx", 
    "Informe anual 2024");

Console.WriteLine($"Documento guardado con ID: {document.Id}");

// Extraer texto de un documento
var text = _wordDocumentService.ExtractTextFromWord(@"C:\Documentos\informe.docx");
Console.WriteLine($"Texto extraído: {text.Substring(0, Math.Min(100, text.Length))}...");

// Recuperar un documento
var storedDoc = await _wordDocumentService.GetDocumentByIdAsync(document.Id);
if (storedDoc != null)
{
    await File.WriteAllBytesAsync(@"C:\Temp\copiado.docx", storedDoc.FileContent);
}
```

## Migraciones de Entity Framework

Para crear la base de datos y las tablas:

```bash
# Agregar una migración
dotnet ef migrations add InitialCreate

# Actualizar la base de datos (creará el archivo WordDocuments.db)
dotnet ef database update
```

La base de datos SQLite se creará automáticamente en el directorio de la aplicación como un archivo llamado `WordDocuments.db`.

## Notas Importantes

1. **Validación**: El servicio valida que el archivo sea un documento Word válido (.docx) antes de guardarlo.

2. **Almacenamiento**: Los documentos se guardan como BLOBs (Binary Large Objects) en la columna `FileContent`. Para documentos muy grandes, considera almacenarlos en el sistema de archivos y guardar solo la ruta en la base de datos.

3. **Seguridad**: Implementa validaciones adicionales según tus necesidades (tamaño máximo, tipos de archivo permitidos, autenticación, etc.).

4. **Rendimiento**: Para consultas que listan documentos, evita traer el `FileContent` usando proyecciones selectivas.

5. **SQLite**: La base de datos es un archivo local (`WordDocuments.db`), ideal para desarrollo, pruebas o aplicaciones de pequeño/mediano tamaño. Para producción con alta concurrencia, considera SQL Server o PostgreSQL.

## Requisitos

- .NET 10 SDK
- Entity Framework Core Tools
