using Microsoft.EntityFrameworkCore;

namespace WordDocumentStorage.Models;

public class ApplicationDbContext : DbContext
{
    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
        : base(options)
    {
    }

    public DbSet<StoredDocument> StoredDocuments { get; set; } = null!;

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        modelBuilder.Entity<StoredDocument>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.FileName).IsRequired().HasMaxLength(255);
            entity.Property(e => e.ContentType).IsRequired().HasMaxLength(100);
            entity.Property(e => e.FileContent).IsRequired();
            entity.Property(e => e.FileSize).IsRequired();
            entity.Property(e => e.CreatedDate).IsRequired();
            entity.Property(e => e.Description).HasMaxLength(500);
        });
    }
}

public class StoredDocument
{
    public int Id { get; set; }
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
    public byte[] FileContent { get; set; } = Array.Empty<byte>();
    public long FileSize { get; set; }
    public DateTime CreatedDate { get; set; }
    public string? Description { get; set; }
}
