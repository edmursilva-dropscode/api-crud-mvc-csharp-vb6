USE [PaschoalottoDesafio]
GO
/****** Object:  Table [dbo].[Usuario]    Script Date: 28/05/2024 10:29:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Usuario](
	[IdUsuario] [int] IDENTITY(1,1) NOT NULL,
	[Nome] [nvarchar](100) NOT NULL,
	[Sobrenome] [nvarchar](60) NOT NULL,
	[Senha] [nvarchar](60) NOT NULL,
	[Email] [nvarchar](250) NOT NULL,
	[Telefone] [nvarchar](100) NOT NULL,
	[Genero] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK_Usuario] PRIMARY KEY CLUSTERED 
(
	[IdUsuario] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[Usuario] ON 

INSERT [dbo].[Usuario] ([IdUsuario], [Nome], [Sobrenome], [Senha], [Email], [Telefone], [Genero]) VALUES (1, N'Edmur', N'Silva', N'1234', N'edmurgsjr@hotmail.com', N'12345', N'masculino')
INSERT [dbo].[Usuario] ([IdUsuario], [Nome], [Sobrenome], [Senha], [Email], [Telefone], [Genero]) VALUES (3, N'Fernando', N'Gomes', N'1234', N'fernando.gomes@hotmail.com', N'12121', N'masculino')
SET IDENTITY_INSERT [dbo].[Usuario] OFF
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_D_Usuario]    Script Date: 28/05/2024 10:29:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[SP_Teste_D_Usuario](@Id INT)
AS
BEGIN
    
    DELETE FROM Usuario 
    WHERE IdUsuario = @Id;
         
END 
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_I_Usuario]    Script Date: 28/05/2024 10:29:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[SP_Teste_I_Usuario](@Nome VARCHAR(100), @Sobrenome VARCHAR(60), @Senha VARCHAR(60), @Email VARCHAR(250), @Telefone VARCHAR(100), @Genero VARCHAR(50))
AS
BEGIN
        
       INSERT INTO Usuario 
       (Nome, Sobrenome, Senha, Email, Telefone, Genero)
       VALUES 
       (@Nome, @Sobrenome, @Senha, @Email, @Telefone, @Genero);
       
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_S_Usuario]    Script Date: 28/05/2024 10:29:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[SP_Teste_S_Usuario](@Id INT)
AS
BEGIN
        
    IF(@Id = 0)            
		SELECT IdUsuario, Nome, sobrenome, Senha, Email, telefone, Genero 
		FROM Usuario (NOLOCK) 
		ORDER BY IdUsuario
        
    IF(@Id <> 0)           
		SELECT IdUsuario, Nome, sobrenome, Senha, Email, telefone, Genero
		FROM Usuario (NOLOCK) 
		WHERE IdUsuario = @Id 
     	ORDER BY IdUsuario;
END
GO
/****** Object:  StoredProcedure [dbo].[SP_Teste_U_Usuario]    Script Date: 28/05/2024 10:29:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[SP_Teste_U_Usuario](@Id INT, @Nome VARCHAR(100), @Sobrenome VARCHAR(60), @Senha VARCHAR(60), @Email VARCHAR(250), @Telefone VARCHAR(100), @Genero VARCHAR(100))
AS
BEGIN
    
    UPDATE Usuario SET   
	Nome = @Nome,
	Sobrenome = @Sobrenome, 
	Senha = @Senha, 
	Email = @Email,
	Telefone = @Telefone,
	Genero = @Genero
	WHERE IdUsuario = @Id;   
END
GO
