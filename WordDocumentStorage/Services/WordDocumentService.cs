using DocumentFormat.OpenXml.Packaging;
using Microsoft.EntityFrameworkCore;
using WordDocumentStorage.Models;

namespace WordDocumentStorage.Services;

public interface IWordDocumentService
{
    /// <summary>
    /// Lee un documento Word desde una ruta y lo guarda en la base de datos
    /// </summary>
    Task<StoredDocument> ReadAndSaveDocumentAsync(string filePath, string? description = null);

    /// <summary>
    /// Lee un documento Word desde un stream y lo guarda en la base de datos
    /// </summary>
    Task<StoredDocument> ReadAndSaveDocumentAsync(Stream documentStream, string fileName, string? description = null);

    /// <summary>
    /// Obtiene un documento por su ID
    /// </summary>
    Task<StoredDocument?> GetDocumentByIdAsync(int id);

    /// <summary>
    /// Obtiene todos los documentos almacenados
    /// </summary>
    Task<List<StoredDocument>> GetAllDocumentsAsync();

    /// <summary>
    /// Elimina un documento de la base de datos
    /// </summary>
    Task<bool> DeleteDocumentAsync(int id);

    /// <summary>
    /// Extrae el contenido de texto de un documento Word
    /// </summary>
    string ExtractTextFromWord(string filePath);
}

public class WordDocumentService : IWordDocumentService
{
    private readonly ApplicationDbContext _context;
    private readonly ILogger<WordDocumentService> _logger;

    public WordDocumentService(ApplicationDbContext context, ILogger<WordDocumentService> logger)
    {
        _context = context;
        _logger = logger;
    }

    /// <summary>
    /// Lee un documento Word desde una ruta y lo guarda en la base de datos
    /// </summary>
    public async Task<StoredDocument> ReadAndSaveDocumentAsync(string filePath, string? description = null)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"El archivo no existe: {filePath}");

        await using var fileStream = File.OpenRead(filePath);
        var fileName = Path.GetFileName(filePath);
        
        return await ReadAndSaveDocumentAsync(fileStream, fileName, description);
    }

    /// <summary>
    /// Lee un documento Word desde un stream y lo guarda en la base de datos
    /// </summary>
    public async Task<StoredDocument> ReadAndSaveDocumentAsync(Stream documentStream, string fileName, string? description = null)
    {
        try
        {
            // Validar que sea un documento Word válido
            ValidateWordDocument(documentStream);

            // Leer todo el contenido del archivo
            byte[] fileContent;
            if (documentStream.CanSeek)
            {
                documentStream.Position = 0;
                fileContent = new byte[documentStream.Length];
                await documentStream.ReadAsync(fileContent.AsMemory(0, (int)documentStream.Length));
            }
            else
            {
                await using var memoryStream = new MemoryStream();
                await documentStream.CopyToAsync(memoryStream);
                fileContent = memoryStream.ToArray();
            }

            // Crear la entidad del documento
            var storedDocument = new StoredDocument
            {
                FileName = fileName,
                ContentType = GetContentType(fileName),
                FileContent = fileContent,
                FileSize = fileContent.Length,
                CreatedDate = DateTime.UtcNow,
                Description = description
            };

            // Guardar en la base de datos
            _context.StoredDocuments.Add(storedDocument);
            await _context.SaveChangesAsync();

            _logger.LogInformation("Documento guardado exitosamente: {FileName} ({FileSize} bytes)", 
                fileName, fileContent.Length);

            return storedDocument;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al guardar el documento: {FileName}", fileName);
            throw;
        }
    }

    /// <summary>
    /// Obtiene un documento por su ID
    /// </summary>
    public async Task<StoredDocument?> GetDocumentByIdAsync(int id)
    {
        return await _context.StoredDocuments.FindAsync(id);
    }

    /// <summary>
    /// Obtiene todos los documentos almacenados
    /// </summary>
    public async Task<List<StoredDocument>> GetAllDocumentsAsync()
    {
        return await _context.StoredDocuments
            .OrderByDescending(d => d.CreatedDate)
            .ToListAsync();
    }

    /// <summary>
    /// Elimina un documento de la base de datos
    /// </summary>
    public async Task<bool> DeleteDocumentAsync(int id)
    {
        var document = await _context.StoredDocuments.FindAsync(id);
        if (document == null)
            return false;

        _context.StoredDocuments.Remove(document);
        await _context.SaveChangesAsync();
        
        _logger.LogInformation("Documento eliminado: {FileName}", document.FileName);
        return true;
    }

    /// <summary>
    /// Extrae el contenido de texto de un documento Word
    /// </summary>
    public string ExtractTextFromWord(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"El archivo no existe: {filePath}");

        try
        {
            using var wordDoc = WordprocessingDocument.Open(filePath, false);
            var body = wordDoc.MainDocumentPart?.Document?.Body;
            return body?.InnerText ?? string.Empty;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al extraer texto del documento: {FilePath}", filePath);
            throw;
        }
    }

    /// <summary>
    /// Valida que el stream contenga un documento Word válido
    /// </summary>
    private void ValidateWordDocument(Stream stream)
    {
        try
        {
            bool canSeek = stream.CanSeek;
            long originalPosition = canSeek ? stream.Position : 0;

            if (canSeek)
                stream.Position = 0;

            using var wordDoc = WordprocessingDocument.Open(stream, false);
            
            // Verificar que tenga una parte principal de documento
            if (wordDoc.MainDocumentPart == null)
                throw new InvalidDataException("El documento no tiene una parte principal válida");

            if (canSeek)
                stream.Position = originalPosition;
        }
        catch (Exception ex) when (ex is not InvalidDataException)
        {
            throw new InvalidDataException("El archivo no es un documento Word válido (.docx)", ex);
        }
    }

    /// <summary>
    /// Obtiene el tipo de contenido según la extensión del archivo
    /// </summary>
    private static string GetContentType(string fileName)
    {
        var extension = Path.GetExtension(fileName).ToLowerInvariant();
        return extension switch
        {
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".docm" => "application/vnd.ms-word.document.macroEnabled.12",
            ".dotx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
            _ => "application/octet-stream"
        };
    }
}
