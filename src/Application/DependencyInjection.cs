namespace Application;
public static class DependencyInjection
{
    public static IServiceCollection AddApplicationServices(this IServiceCollection services)
    {
        services.AddSingleton<ICommonHelpers, CommonHelpers>();
        services.AddSingleton<IDateTimeHelpers, DateTimeHelpers>();
        services.AddSingleton<IEscaperHelpers, EscaperHelpers>();
        services.AddSingleton<IGeneralHelpers, GeneralHelpers>();
        services.AddSingleton<IGeneratorHelpers, GeneratorHelpers>();
        services.AddSingleton<ISecurityHelpers, SecurityHelpers>();
        services.AddSingleton<IValidatorHelpers, ValidatorHelpers>();
        return services;
    }
}