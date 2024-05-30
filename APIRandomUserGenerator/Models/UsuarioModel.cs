using System.ComponentModel.DataAnnotations;

namespace APIRandomUserGenerator.Models
{
    public class UsuarioModel
    {
        //propriedades
        [Key]
        public int IdUsuario { get; set; }

        [Required, MaxLength(100)]
        [Display(Name = "Nome")]
        public string Nome { get; set; } = string.Empty;

        [Required, MaxLength(60)]
        [Display(Name = "Sobrenome")]
        public string Sobrenome { get; set; } = string.Empty;

        [Required, MaxLength(60)]
        [Display(Name = "Senha")]
        public string Senha { get; set; } = string.Empty;

        [Required, MaxLength(250)]
        [Display(Name = "Email")]
        public string Email { get; set; } = string.Empty;

        [Required, MaxLength(100)]
        [Display(Name = "Telefone")]
        public string Telefone { get; set; } = string.Empty;

        [Required, MaxLength(50)]
        [Display(Name = "Gênero")]
        public string Genero { get; set; } = string.Empty;
    }
}
