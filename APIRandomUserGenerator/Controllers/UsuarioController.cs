using APIRandomUserGenerator.Context;
using APIRandomUserGenerator.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json.Linq;
using PdfSharpCore.Drawing.Layout;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System.Text;


namespace APIRandomUserGenerator.Controllers
{
    public class UsuarioController : Controller
    {
        //Atributo somente leitura para guardar conexao com bnco de dados
        private readonly Contexto _context;
        private readonly HttpClient _httpClient;

        //Construtor
        public UsuarioController(Contexto context)
        {
            this._context = context;
            _httpClient = new HttpClient();
        }

        //Retorna listagem de usuarios ordernada por nome
        public async Task<IActionResult> Index()
        {
            return View(await _context.Usuarios!.OrderBy(x => x.Nome).AsNoTracking().ToListAsync());
        }

        //Cadastrar usuario
        [HttpGet]
        public async Task<IActionResult> Cadastrar(int? id)
        {
            if (id.HasValue)
            {
                var usuario = await _context.Usuarios!.FindAsync(id);
                if (usuario == null)
                {
                    return NotFound();
                }
                return View(usuario);
            }
            return View(new UsuarioModel());
        }

        //Funcao validacao se o usuario existe
        private bool UsuarioExiste(int id)
        {
            return _context.Usuarios!.Any(x => x.IdUsuario == id);
        }

        //Alterar usuario
        [HttpPost]
        public async Task<IActionResult> Cadastrar(int? id, [FromForm] UsuarioModel usuario)
        {
            if (ModelState.IsValid)
            {
                if (id.HasValue)
                {
                    if (UsuarioExiste(id.Value))
                    {
                        _context.Usuarios?.Update(usuario);
                        if (await _context.SaveChangesAsync() > 0)
                        {
                            TempData["mensagem"] = MensagemModel.Serializar("Usuário alterado com sucesso.");
                        }
                        else
                        {
                            TempData["mensagem"] = MensagemModel.Serializar("Erro ao alterar usuário.", TipoMensagem.Erro);
                        }
                    }
                    else
                    {
                        TempData["mensagem"] = MensagemModel.Serializar("Usuário não encontrado.", TipoMensagem.Erro);
                    }
                }
                else
                {
                    _context.Usuarios?.Add(usuario);
                    if (await _context.SaveChangesAsync() > 0)
                    {
                        TempData["mensagem"] = MensagemModel.Serializar("Usuário cadastrado com sucesso.");
                    }
                    else
                    {
                        TempData["mensagem"] = MensagemModel.Serializar("Erro ao cadastrar usuário.", TipoMensagem.Erro);
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            else
            {
                return View(usuario);
            }
        }

        //Funcão validacao se usuario foi informado
        [HttpGet]
        public async Task<IActionResult> Excluir(int? id)
        {
            if (!id.HasValue)
            {
                TempData["mensagem"] = MensagemModel.Serializar("Usuário não informado.", TipoMensagem.Erro);
                return RedirectToAction(nameof(Index));
            }

            var usuario = await _context.Usuarios!.FindAsync(id);
            if (usuario == null)
            {
                TempData["mensagem"] = MensagemModel.Serializar("Usuário não encontrado.", TipoMensagem.Erro);
                return RedirectToAction(nameof(Index));
            }

            return View(usuario);
        }

        //Excluir usuario
        [HttpPost]
        public async Task<IActionResult> Excluir(int id)
        {
            var usuario = await _context.Usuarios!.FindAsync(id);
            if (usuario != null)
            {
                _context.Usuarios.Remove(usuario);
                if (await _context.SaveChangesAsync() > 0)
                    TempData["mensagem"] = MensagemModel.Serializar("Usuário excluído com sucesso.");
                else
                    TempData["mensagem"] = MensagemModel.Serializar("Não foi possível excluir o usuário.", TipoMensagem.Erro);
                return RedirectToAction(nameof(Index));
            }
            else
            {
                TempData["mensagem"] = MensagemModel.Serializar("Usuário não encontrado.", TipoMensagem.Erro);
                return RedirectToAction(nameof(Index));
            }
        }


        [HttpPost]
        public async Task<IActionResult> AddRandomUsers()
        {
            try
            {
                int countToAdd = 20; // Defina o número desejado de usuários aleatórios a serem adicionados

                // Verificar se já existem 20 ou mais usuários na base de dados
                if (_context.Usuarios!.Count() >= 20)
                {
                    //return BadRequest("Já existem 20 ou mais usuários na base de dados.");
                    TempData["mensagem"] = MensagemModel.Serializar("Já existem 20 ou mais usuários na base de dados.", TipoMensagem.Erro);
                    return RedirectToAction(nameof(Index));
                }

                // Calcular quantos usuários ainda podem ser adicionados
                int usersToAdd = countToAdd - _context.Usuarios!.Count();

                var httpClient = new HttpClient();
                var response = await httpClient.GetAsync($"https://randomuser.me/api/?results={usersToAdd}");

                if (!response.IsSuccessStatusCode)
                    return StatusCode((int)response.StatusCode, "Falha ao obter usuários aleatórios da API.");

                var jsonString = await response.Content.ReadAsStringAsync();
                var json = JObject.Parse(jsonString);

                var randomUsers = json["results"]!.ToObject<List<JObject>>();

                //Valida de email já cadastrado na base de dados
                foreach (var userJson in randomUsers!)
                {
                    var email = userJson["email"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(email) && !_context.Usuarios!.Any(u => u.Email == email))
                    {
                        var name = userJson["name"];
                        var firstName = name?["first"]?.ToString() ?? "";
                        var lastName = name?["last"]?.ToString() ?? "";

                        var user = new UsuarioModel
                        {
                            Nome = firstName,
                            Sobrenome = lastName,
                            Senha = SenhaAleatoria(),
                            Email = email,
                            Telefone = userJson["phone"]?.ToString() ?? "",
                            Genero = userJson["gender"]?.ToString() ?? ""
                        };
                        _context.Usuarios!.Add(user);
                    }
                }

                await _context.SaveChangesAsync();

                TempData["mensagem"] = MensagemModel.Serializar($"Adicionados {usersToAdd} usuários aleatórios à base de dados.");
                //return Ok($"Adicionados {usersToAdd} usuários aleatórios à base de dados.");
            }
            catch (Exception ex)
            {
                TempData["mensagem"] = MensagemModel.Serializar($"Ocorreu um erro: {ex.Message}", TipoMensagem.Erro);
                //return StatusCode(500, $"Ocorreu um erro: {ex.Message}");
            }

            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public IActionResult GerarRelatorioPDF()
        {
            var usuarios = _context.Usuarios!.OrderBy(x => x.Nome).ToList();

            using (MemoryStream stream = new MemoryStream())
            {
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();

                // Define a orientação da página para paisagem
                page.Orientation = PdfSharpCore.PageOrientation.Landscape;

                XGraphics gfx = XGraphics.FromPdfPage(page);

                // Define as fontes
                XFont titleFont = new XFont("Verdana", 12, XFontStyle.Bold);
                XFont dataFont = new XFont("Verdana", 9, XFontStyle.Regular);

                int yPos = 50;

                // Desenha o título do relatório
                gfx.DrawString("Relatório de Usuários", titleFont, XBrushes.Black, new XPoint(50, yPos));

                yPos += 20;

                // Definindo posições x para cada coluna
                double nomeX = 50;
                double sobrenomeX = 150;
                double senhaX = 250;
                double emailX = 380;
                double telefoneX = 590;
                double generoX = 700;

                // Desenha os cabeçalhos das colunas
                gfx.DrawString("Nome", dataFont, XBrushes.Black, new XPoint(nomeX, yPos));
                gfx.DrawString("Sobrenome", dataFont, XBrushes.Black, new XPoint(sobrenomeX, yPos));
                gfx.DrawString("Senha", dataFont, XBrushes.Black, new XPoint(senhaX, yPos));
                gfx.DrawString("Email", dataFont, XBrushes.Black, new XPoint(emailX, yPos));
                gfx.DrawString("Telefone", dataFont, XBrushes.Black, new XPoint(telefoneX, yPos));
                gfx.DrawString("Genero", dataFont, XBrushes.Black, new XPoint(generoX, yPos));

                yPos += 20;

                // Desenha os dados dos usuários
                foreach (var usuario in usuarios)
                {
                    gfx.DrawString(usuario.Nome, dataFont, XBrushes.Black, new XPoint(nomeX, yPos));
                    gfx.DrawString(usuario.Sobrenome, dataFont, XBrushes.Black, new XPoint(sobrenomeX, yPos));
                    gfx.DrawString(usuario.Senha, dataFont, XBrushes.Black, new XPoint(senhaX, yPos));
                    gfx.DrawString(usuario.Email, dataFont, XBrushes.Black, new XPoint(emailX, yPos));
                    gfx.DrawString(usuario.Telefone, dataFont, XBrushes.Black, new XPoint(telefoneX, yPos));
                    gfx.DrawString(usuario.Genero, dataFont, XBrushes.Black, new XPoint(generoX, yPos));

                    yPos += 20;
                }

                pdf.Save(stream, false);
                stream.Position = 0;

                return File(stream.ToArray(), "application/pdf");
            }
        }

        private string SenhaAleatoria()
        {
            const string caracteresPermitidos = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var random = new Random();
            var senha = new StringBuilder();

            // Gerar uma senha com até 15 caracteres
            for (int i = 0; i < 15; i++)
            {
                int index = random.Next(caracteresPermitidos.Length);
                senha.Append(caracteresPermitidos[index]);
            }

            return senha.ToString();
        }

    }
    public class RandomUserResponse
    {
        public List<RandomUser>? Results { get; set; }
    }

    public class RandomUser
    {
        public RandomUserName? Name { get; set; }
        public string Senha { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Phone { get; set; } = string.Empty;
        public string Gender { get; set; } = string.Empty;
    }

    public class RandomUserName
    {
        public string First { get; set; } = string.Empty;
        public string Last { get; set; } = string.Empty;
    }

}
