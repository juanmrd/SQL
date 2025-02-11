USE [ModeloIntegral]
GO
/****** Object:  StoredProcedure [dbo].[DimCalendario]    Script Date: 11-02-2025 18:03:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Dimesión Calendario
ALTER   PROCEDURE [dbo].[DimCalendario]
AS
	DECLARE @Fec_incio DATE;   -- Declarar Variable
	DECLARE @Fec_fin DATE;   -- Declarar Variable
	DECLARE @anio INT;   -- Declarar Variable
	DECLARE @trimestre INT;   -- Declarar Variable
	DECLARE @mes_num INT;   -- Declarar Variable
	DECLARE @mes_corto VARCHAR (3);   -- Declarar Variable
	DECLARE @mes_largo VARCHAR (20);   -- Declarar Variable
	DECLARE @dia INT;   -- Declarar Variable
	
	CREATE TABLE #DIN_CALENDARIO
	(
		Fecha DATE,
		Año INT,
		Trimestre INT,
		NumMes INT,
		Mes VARCHAR (3),
		MesLargo VARCHAR (20),
		Dia INT
	);
	SELECT 
	@Fec_incio = MIN(OrderDate),
	@Fec_fin = 	 MAX(OrderDate)
	FROM  [Sales].[SalesOrderHeader];

	WHILE @Fec_incio <= @Fec_fin
	BEGIN 
		SET @anio = (SELECT YEAR(@Fec_incio));
		SET @trimestre = (SELECT DATEPART(QUARTER, @Fec_incio));
		SET @mes_num = (SELECT MONTH(@Fec_incio));
		SET @mes_corto = (SELECT FORMAT(@Fec_incio, 'MMM', 'es-pe'));
		SET @mes_largo = (SELECT FORMAT(@Fec_incio, 'MMMM', 'es-pe'));
		SET @dia = (SELECT DAY(@Fec_incio));
		
		INSERT INTO #DIN_CALENDARIO (Fecha, Trimestre, Año, NumMes, Mes, MesLargo, Dia)
		SELECT @Fec_incio, @trimestre, @anio, @mes_num, @mes_corto, @mes_largo, @dia;

		SET @Fec_incio = DATEADD (DAY, 1, @Fec_incio)
	END;

	SELECT * FROM #DIN_CALENDARIO
