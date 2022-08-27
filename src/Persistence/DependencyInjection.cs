using Application.Models.IdentityModels;
using Domain.Constants;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Persistence;

namespace Application;
public static class DependencyInjection
{
    public static IServiceCollection AddPersistenceServices(this IServiceCollection services, IConfiguration configuration)
    {
        // DB Context Setting
        services.AddDbContext<AppDbContext>(options =>
            options.UseSqlServer(configuration.GetConnectionString("DHA_Accounts")));

        // Authentication Configurations Options
        services.Configure<IdentityOptions>(options =>
        {
            // Password settings.
            options.Password.RequireDigit = true;
            options.Password.RequireLowercase = true;
            options.Password.RequireNonAlphanumeric = true;
            options.Password.RequireUppercase = true;
            options.Password.RequiredLength = 6;
            options.Password.RequiredUniqueChars = 1;

            // Lockout settings.
            options.Lockout.DefaultLockoutTimeSpan = TimeSpan.FromMinutes(5);
            options.Lockout.MaxFailedAccessAttempts = 5;
            options.Lockout.AllowedForNewUsers = true;

            // User settings.
            options.User.AllowedUserNameCharacters = AppConstants.AllowedUserNameCharacters;
            options.User.RequireUniqueEmail = false;
        });

        // Identity Services
        services.AddIdentity<ApplicationUser, IdentityRole>().AddEntityFrameworkStores<AppDbContext>().AddDefaultTokenProviders();

        services.AddAuthentication()
           .AddFacebook(facebookOptions =>
           {
               IConfigurationSection facebookAuthNSection = configuration.GetSection("ExternalAuthenticators:Faceboook");
               facebookOptions.ClientId = facebookAuthNSection["AppId"];
               facebookOptions.ClientSecret = facebookAuthNSection["AppSecret"];
               facebookOptions.AccessDeniedPath = "/AccessDeniedPathInfo";
           })
           .AddGoogle(googleOptions =>
           {
               IConfigurationSection googleAuthNSection = configuration.GetSection("ExternalAuthenticators:Google");
               googleOptions.ClientId = googleAuthNSection["AppId"];
               googleOptions.ClientSecret = googleAuthNSection["AppSecret"];
           })
           .AddMicrosoftAccount(microsoftOptions =>
           {
               IConfigurationSection microsoftAuthNSection = configuration.GetSection("ExternalAuthenticators:Microsoft");
               microsoftOptions.ClientId = microsoftAuthNSection["AppSecret"];
               microsoftOptions.ClientSecret = microsoftAuthNSection["AppSecret"];
           });

        return services;
    }
}