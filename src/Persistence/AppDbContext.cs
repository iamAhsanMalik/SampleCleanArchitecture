using Application.Models.IdentityModels;

namespace Persistence;

internal class AppDbContext : IdentityDbContext<ApplicationUser>
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }
    protected override void OnModelCreating(ModelBuilder builder)
    {
        base.OnModelCreating(builder);

        // Customize the ASP.NET Identity model and override the defaults if needed.
        // Add your customizations after calling base.OnModelCreating(builder);
        // For example, here we are renaming the ASP.NET Identity table names.

        builder.Entity<ApplicationUser>(b =>
          b.ToTable("Users"));

        builder.Entity<IdentityUserClaim<string>>(b =>
            b.ToTable("UserClaims"));

        builder.Entity<IdentityUserLogin<string>>(b =>
            b.ToTable("UserLogins"));

        builder.Entity<IdentityUserToken<string>>(b =>
            b.ToTable("UserTokens"));

        builder.Entity<IdentityRole>(b => b.ToTable("Roles"));

        builder.Entity<IdentityRoleClaim<string>>(b =>
            b.ToTable("RoleClaims"));

        builder.Entity<IdentityUserRole<string>>(b => b.ToTable("UserRoles"));
    }
}
