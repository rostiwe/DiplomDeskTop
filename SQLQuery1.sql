﻿select [Id_клиента],[Id_Типа],SUM([Изменение_накопления]) as Накопение from [dbo].[Накопление_клиентов] group by [Id_клиента],[Id_Типа] order by [Id_клиента]