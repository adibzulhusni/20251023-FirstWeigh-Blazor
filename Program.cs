using _20251013_FirstWeigh_Blazor.Components;
using FirstWeigh.Services;

namespace _20251013_FirstWeigh_Blazor
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);
            // Register UserService
            builder.Services.AddScoped<UserService>();
            // Add this line:
            builder.Services.AddScoped<BrowserStorageService>();
            // Register IngredientService
            builder.Services.AddSingleton<IngredientService>();
            // Register BowlService
            builder.Services.AddSingleton<BowlService>();
            // Register ModbusScaleService
            builder.Services.AddSingleton<ModbusScaleService>();
            // Register RecipeService
            builder.Services.AddSingleton<RecipeService>();
            builder.Services.AddScoped<IBatchService, BatchService>();
            builder.Services.AddScoped<IWeighingService, WeighingService>();
            builder.Services.AddSingleton<ReportService>();
            // Add this line where you register services:
            builder.Services.AddScoped<AuthenticationService>();
            // Add services to the container.
            builder.Services.AddRazorComponents()
                .AddInteractiveServerComponents();

            var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (!app.Environment.IsDevelopment())
            {
                app.UseExceptionHandler("/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            app.UseHttpsRedirection();

            app.UseAntiforgery();

            app.MapStaticAssets();
            app.MapRazorComponents<App>()
                .AddInteractiveServerRenderMode();

            app.Run();
        }
    }
}
