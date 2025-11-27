using _20251013_FirstWeigh_Blazor.Components;
using FirstWeigh.Services;
using Serilog;
using Serilog.Events;

namespace _20251013_FirstWeigh_Blazor
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // ? SERILOG CONFIGURATION - Configure logging first
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .MinimumLevel.Override("Microsoft", LogEventLevel.Information)
                .MinimumLevel.Override("Microsoft.AspNetCore", LogEventLevel.Warning)
                .Enrich.FromLogContext()
                .WriteTo.Console(
                    outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
                .WriteTo.File(
                    path: "Logs/firstweigh_.log",
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 30,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {SourceContext}{NewLine}{Message:lj}{NewLine}{Exception}",
                    shared: true)
                .CreateLogger();

            try
            {
                Log.Information("Starting FirstWeigh application...");

                var builder = WebApplication.CreateBuilder(args);

                // ? Use Serilog for all logging
                builder.Host.UseSerilog();

                // Register UserService
                builder.Services.AddScoped<UserService>();
                builder.Services.AddScoped<BrowserStorageService>();
                builder.Services.AddSingleton<IngredientService>();
                builder.Services.AddSingleton<BowlService>();
                builder.Services.AddSingleton<ModbusScaleService>();
                builder.Services.AddSingleton<RecipeService>();
                builder.Services.AddScoped<IBatchService, BatchService>();
                builder.Services.AddScoped<IWeighingService, WeighingService>();
                builder.Services.AddSingleton<ReportService>();
                builder.Services.AddScoped<AuthenticationService>();
                builder.Services.AddScoped<LoginAttemptService>();
                builder.Services.AddScoped<AuditLogService>();

                builder.Services.AddRazorComponents()
                    .AddInteractiveServerComponents();

                var app = builder.Build();

                if (!app.Environment.IsDevelopment())
                {
                    app.UseExceptionHandler("/Error");
                    app.UseHsts();
                }

                app.UseHttpsRedirection();
                app.UseSerilogRequestLogging();
                app.UseAntiforgery();

                app.MapStaticAssets();
                app.MapRazorComponents<App>()
                    .AddInteractiveServerRenderMode();

                app.Run();
            }
            catch (Exception ex)
            {
                Log.Fatal(ex, "Application terminated unexpectedly");
            }
            finally
            {
                Log.CloseAndFlush();
            }
        }
    }
}