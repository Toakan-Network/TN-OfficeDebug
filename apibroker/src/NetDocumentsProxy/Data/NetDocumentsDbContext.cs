using Microsoft.EntityFrameworkCore;
using NetDocumentsProxy.Models;

namespace NetDocumentsProxy.Data;

public class NetDocumentsDbContext : DbContext
{
    public NetDocumentsDbContext(DbContextOptions<NetDocumentsDbContext> options) : base(options)
    {
    }

    public DbSet<UserToken> UserTokens { get; set; }
    public DbSet<ApiLog> ApiLogs { get; set; }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<UserToken>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.UserId).IsRequired().HasMaxLength(100);
            entity.Property(e => e.AccessToken).IsRequired();
            entity.Property(e => e.RefreshToken).IsRequired();
            entity.Property(e => e.ExpiresAt).IsRequired();
            entity.Property(e => e.CreatedAt).IsRequired();
            entity.Property(e => e.UpdatedAt).IsRequired();
            
            entity.HasIndex(e => e.UserId).IsUnique();
        });

        modelBuilder.Entity<ApiLog>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.UserId).HasMaxLength(100);
            entity.Property(e => e.Endpoint).IsRequired().HasMaxLength(500);
            entity.Property(e => e.Method).IsRequired().HasMaxLength(10);
            entity.Property(e => e.StatusCode).IsRequired();
            entity.Property(e => e.RequestedAt).IsRequired();
            entity.Property(e => e.ResponseTime).IsRequired();
            
            entity.HasIndex(e => e.RequestedAt);
            entity.HasIndex(e => e.UserId);
        });

        base.OnModelCreating(modelBuilder);
    }
}