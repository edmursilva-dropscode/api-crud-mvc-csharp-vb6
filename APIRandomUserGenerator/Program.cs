using Microsoft.EntityFrameworkCore;
using APIRandomUserGenerator.Context;


namespace APIRandomUserGenerator
{
    public class Program
    {
        public static void Main(string[] args)
        {

            var builder = WebApplication.CreateBuilder(args);

            //Adicionando servicos do framework MVC
            builder.Services.AddControllersWithViews();

            //adicionando string de conexao do banco de dados            
            builder.Services.AddEntityFrameworkNpgsql() .AddDbContext<Contexto>(options => options.UseNpgsql("Host=localhost;Port=5432;Pooling=true;Database=PaschoalottoDesafio;User Id=postgres;Password=admin;"));

            var app = builder.Build();

            //Configura solicitação HTTP.
            if (!app.Environment.IsDevelopment())
            {
                app.UseExceptionHandler("/Home/Error");
                app.UseHsts();
            }

            app.UseHttpsRedirection();

            app.UseStaticFiles();

            app.UseRouting();

            app.UseAuthorization();

            //Mapeia padrão MVC
            app.MapControllerRoute(
                name: "default",
                pattern: "{controller=Home}/{action=Index}/{id?}");

            app.Run();

        }
    }
}