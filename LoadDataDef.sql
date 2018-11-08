-- -----------------------------------------------------------------------------
--
-- Script: LoadDataDef.sql
--
-- purpose: Create LoadDataDef table.
--
-- -----------------------------------------------------------------------------
--
-- rev log
--
-- date:  27-FEB-2018
-- author:  Craig Nobili
-- desc: original
--
-- ---------------------------------------------------------------------------  

if object_id('LoadDataDef') is not null
  drop table LoadDataDef

go

create table LoadDataDef
(
  LoadDataDefKey         bigint primary key identity
, LoadDataName           varchar(256)                 not null
, SourceConnType         varchar(25)                  not null                 
, SourceConnStr          varchar(256)
, SourceQuery            varchar(max)  
, DestConnStr            varchar(256)                 not null
, DestDbName             varchar(256)                 not null
, DestSchemaName         varchar(256)                 not null
, DestTableName          varchar(256)                 not null
, LoadType               varchar(25)                  not null
, BatchSize              bigint  
, Directory              varchar(256)
, FilePattern            varchar(256)
, Worksheet              varchar(256)
, Header                 varchar(1)
, Delimiter              varchar(10)
, ActiveFlg              varchar(1)                   not null
, GroupName              varchar(256)
, constraint LoadDataDef_UK1
  unique (LoadDataName)
, constraint LoadDataDef_CK1
  check (ActiveFlg in ('Y', 'N'))
, constraint LoadDataDef_CK2
  check (SourceConnType in ('MSSQL', 'OLEDB', 'ODBC', 'CSV', 'FLATFILE', 'EXCEL'))
, constraint LoadDataDef_CK3
  check (LoadType in ('FULL', 'APPEND'))
, constraint LoadDataDef_CK4
  check (Header in ('Y', 'N'))
)
;
