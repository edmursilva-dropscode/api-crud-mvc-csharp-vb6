using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using Microsoft.EntityFrameworkCore;
using APIUsuario.Context;
using APIUsuario.Swagger;


namespace APIUsuario
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var myAllowSpecificOrigins = "myAllowSpecificOrigins";   // Adicionando nível/origem de conexão de sistema
            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.
            builder.Services.AddControllers();

            //adicionando string de conexao do banco de dados            
            builder.Services.AddEntityFrameworkNpgsql().AddDbContext<Contexto>(options => options.UseNpgsql("Host=localhost;Port=5432;Pooling=true;Database=PaschoalottoDesafio;User Id=postgres;Password=admin;"));

            // Habilitar CORS para ser usado pelo sistema
            builder.Services.AddCors(options => options.AddPolicy(name: myAllowSpecificOrigins, builder => builder.WithOrigins("http://localhost:4200").AllowAnyMethod().AllowAnyHeader()));

            // Configuração do Swagger
            builder.Services.AddEndpointsApiExplorer();

            //builder.Services.AddSwaggerGen();
            builder.Services.AddInfrastructureSwagger();      //foi desabilidado o builder.Services.AddSwaggerGen() para configurar o layout do Swagger

            // Configuração do JWT
            var key = Encoding.ASCII.GetBytes(builder.Configuration["Jwt:Key"]);
            builder.Services.AddAuthentication(options =>
            {
                options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
                options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
            }).AddJwtBearer(options =>
            {
                options.RequireHttpsMetadata = false;
                options.SaveToken = true;
                options.TokenValidationParameters = new TokenValidationParameters
                {
                    ValidateIssuer = true,
                    ValidateAudience = true,
                    ValidateLifetime = true,
                    ValidateIssuerSigningKey = true,
                    ValidIssuer = builder.Configuration["Jwt:Issuer"],
                    ValidAudience = builder.Configuration["Jwt:Audience"],
                    IssuerSigningKey = new SymmetricSecurityKey(key)
                };
            });

            var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseHttpsRedirection();

            app.UseCors(myAllowSpecificOrigins);     // Habilitar CORS

            app.UseAuthentication();  // Adicionar autenticação
            app.UseAuthorization();

            app.MapControllers();

            app.Run();
        }
    }
}
