# TransferExcelToDatabase

目的：用C#代码将文件夹中的excel文件中的信息抓取出来，插入到数据库中（数据库script如下）

CREATE TABLE [dbo].[APS_ToolList](
	[File] [nvarchar](max) NULL,
	[APS_No] [nvarchar](max) NULL,
	[Sheet_Name] [nvarchar](max) NULL,
	[Setup_Type] [nvarchar](max) NULL,
	[Setup_No] [nvarchar](max) NULL,
	[Mod_Date] [nvarchar](max) NULL,
	[Line_No] [nvarchar](max) NULL,
	[Tool_No] [nvarchar](max) NULL,
	[Tool_Holder] [nvarchar](max) NULL,
	[Stick_Out] [nvarchar](max) NULL,
	[Insert] [nvarchar](max) NULL,
	[P/E] [nvarchar](max) NULL,
	[Comment] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

所要获取的字段参见数据库字段。

请以3005文件中030 sheet name为例，获取数据如下（复制粘贴到excel更直观）

File	APS_No	Sheet_Name	Setup_Type	Setup_No	Mod_Date	Line_No	Tool_No	Tool_Holder	Stick_Out	Insert	P/E	Comment
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	1	04	SCLCR164D	n/a	CCGT431-UM H10A	Check	Kennametal
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	2	09	A20T-DWLNR4	2.2	WNMG431-PP IC10	Check	Iscar
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	3	08	TF1165 (3005-080-T01)	2.2	.187 W, R.030 Brazed Carbide	Check	Face Groove, manufactured by Avitec
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	4	07	TF1166 (3005-080-T02)	2.2	.126 W, R .063 Brazed Carbide	Check	Full Rad. I.D Groove, manufactured by Avitec
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	5	01	TF1167 (3005-080-T04)	2.2	.128 W, R .030 Brazed Carbide	Check	O.D Groove, manufactured by Avitec but modified in-house (.040 off back)
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	6	11	A20-NEL3	2.2	NG3062RK KC730	Check	Kennametal
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	1	04	SCLCR164D	n/a	CCGT431-UM H10A	Check	Kennametal
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	2	09	A20T-DWLNR4	2.2	WNMG431-PP IC10	Check	Iscar
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	3	08	TF1165 (3005-080-T01)	2.2	.187 W, R.030 Brazed Carbide	Check	Face Groove, manufactured by Avitec
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	4	07	TF1166 (3005-080-T02)	2.2	.126 W, R .063 Brazed Carbide	Check	Full Rad. I.D Groove, manufactured by Avitec
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	5	01	TF1167 (3005-080-T04)	2.2	.128 W, R .030 Brazed Carbide	Check	O.D Groove, manufactured by Avitec but modified in-house (.040 off back)
L:\Data_Setups\3005\docs\3005.xlsx	3005	030	Lathe Setup Sheet	3005-030-H	Nov 15, 2016	6	11	A20-NEL3	2.2	NG3062RK KC730	Check	Kennametal

问题：
1.有些文件抓取不到，比如6457，因为sheet name 为 020-B
2.有些文件列名（在excel中，比如tool holder）写成 Tool holder, SP 2 (right side)， 因此C# 抓取不到。例子为6457的020-B
3.在对文件后缀的判断上，依据版本的不同，有xls和xlsx，我认为源代码有bug，在180行，但我还不知道怎样兼顾两种文件格式。

目前调试情况：
我建立了数据库，设置了参数。并将要获取的excel文件设置路径为D:\Data_Setups\3005\docs（要放在根目录下）
调试过程中有可能需要插件AccessDatabaseEngine.exe
现在可以获取到文件信息，但卡在372行，提示有语法错误，但我找不到错误所在。







