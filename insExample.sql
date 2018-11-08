insert into LoadDataDef
(
  LoadDataName
, SourceConnType
, SourceConnStr
, SourceQuery
, DestConnStr
, DestDbName
, DestSchemaName
, DestTableName
, LoadType
, BatchSize
, ActiveFlg
, GroupName
)
values
(
  'ExampleMSSQL'
, 'MSSQL'  
, 'Integrated Security=SSPI;Initial Catalog=K2;Data Source=ahs-bi-db;'
, 'select * from cnobiliTest'
, 'Integrated Security=SSPI;Initial Catalog=K2;Data Source=ahs-bi-dbt;'
, 'K2'
, 'dbo'
, 'cnobiliTest'
, 'FULL'
, 100000
, 'Y'
, 'TEST_MSSQL'
);

insert into LoadDataDef
(
  LoadDataName
, SourceConnType
, SourceConnStr
, SourceQuery
, DestConnStr
, DestDbName
, DestSchemaName
, DestTableName
, LoadType
, BatchSize
, ActiveFlg
, GroupName
)
values
(
  'ExampleODBC'
, 'ODBC'  
, 'Driver={SQL Server};Server=ahs-bi-db;Database=K2;Trusted_Connection=yes;'
, 'select * from cnobiliTest'
, 'Integrated Security=SSPI;Initial Catalog=K2;Data Source=ahs-bi-dbt;'
, 'K2'
, 'dbo'
, 'cnobiliTest'
, 'APPEND'
, 100000
, 'Y'
, 'TEST_ODBC'
);

insert into LoadDataDef
(
  LoadDataName
, SourceConnType
, SourceConnStr
, SourceQuery
, DestConnStr
, DestDbName
, DestSchemaName
, DestTableName
, LoadType
, BatchSize
, ActiveFlg
, GroupName
)
values
(
  'ExampleOLEDB'
, 'OLEDB' 
, 'Provider=SQLOLEDB;Data Source=ahs-bi-db;Initial Catalog=K2;Integrated Security=SSPI;'
, 'select * from cnobiliTest'
, 'Integrated Security=SSPI;Initial Catalog=K2;Data Source=ahs-bi-dbt;'
, 'K2'
, 'dbo'
, 'cnobiliTest'
, 'APPEND'
, 100000
, 'Y'
, 'TEST_OLEDB'
);

insert into LoadDataDef
(
  LoadDataName
, SourceConnType
, SourceConnStr
, SourceQuery
, DestConnStr
, DestDbName
, DestSchemaName
, DestTableName
, LoadType
, BatchSize
, Directory
, FilePattern
, Header
, Delimiter
, ActiveFlg
, GroupName
)
values
(
  'ExampleCSV'
, 'CSV' 
, 'N/A'
, null
, 'Integrated Security=SSPI;Initial Catalog=K2;Data Source=ahs-bi-dbt;'
, 'K2'
, 'dbo'
, 'cnobiliTest'
, 'FULL'
, 100000
, 'c:\Temp'
, '*test.csv'
, 'Y'
, ','
, 'Y'
, 'TEST_CSV'
);

insert into LoadDataDef
(
  LoadDataName
, SourceConnType
, SourceConnStr
, SourceQuery
, DestConnStr
, DestDbName
, DestSchemaName
, DestTableName
, LoadType
, BatchSize
, Directory
, FilePattern
, Header
, Delimiter
, ActiveFlg
, GroupName
)
values
(
  'ExampleFLATFILE'
, 'FLATFILE' 
, 'N/A'
, null
, 'Integrated Security=SSPI;Initial Catalog=K2;Data Source=ahs-bi-dbt;'
, 'K2'
, 'dbo'
, 'cnobiliTest'
, 'APPEND'
, 100000
, 'c:\Temp'
, '*test.txt'
, 'Y'
, '|'
, 'Y'
, 'TEST_FLATFILE'
);

insert into LoadDataDef
(
  LoadDataName
, SourceConnType
, SourceConnStr
, SourceQuery
, DestConnStr
, DestDbName
, DestSchemaName
, DestTableName
, LoadType
, BatchSize
, Directory
, FilePattern
, Worksheet
, Header
, Delimiter
, ActiveFlg
, GroupName
)
values
(
  'ExampleEXCEL'
, 'EXCEL' 
, 'N/A'
, null
, 'Integrated Security=SSPI;Initial Catalog=K2;Data Source=ahs-bi-dbt;'
, 'K2'
, 'dbo'
, 'cnobiliTest'
, 'FULL'
, 100000
, 'c:\Temp'
, 'excelTes*'
, 'Sheet1'
, 'N'
, null
, 'Y'
, 'TEST_EXCEL'
);
