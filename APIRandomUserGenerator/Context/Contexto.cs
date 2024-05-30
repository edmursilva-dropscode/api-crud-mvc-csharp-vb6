using APIRandomUserGenerator.Models;
using Microsoft.EntityFrameworkCore;

namespace APIRandomUserGenerator.Context
{
    public class Contexto : DbContext
    {
        //Construtor de ligação da base de dados
        public Contexto(DbContextOptions<Contexto> options) : base(options)
        {
        }

        //Propriedade representando a entidade/tabela 
        public DbSet<UsuarioModel>? Usuarios { get; set; }     //Usuarios

        //Altera o nome da tabela a ser criada no banco de dados
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<UsuarioModel>().ToTable("Usuario");
        }

    }
}
