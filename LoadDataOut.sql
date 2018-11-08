-- -----------------------------------------------------------------------------
--
-- Script: LoadDataOut.sql
--
-- purpose: Create LoadDataOut table.
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

if object_id('LoadDataOut') is not null
  drop table LoadDataOut

go

create table LoadDataOut
(
  LoadDataOutKey         bigint primary key identity
, LoadDataName           varchar(256)  
, ProgramName            varchar(256)
, SqlStatement           varchar(8000)
, BegDateTime            datetime
, EndDateTime            datetime
, Status                 varchar(20)
, NumRecs                bigint
, DestConnStr            varchar(256) 
, DestDbName             varchar(256) 
, DestSchemaName         varchar(256) 
, DestTableName          varchar(256) 
, LoadType               varchar(25)  
, BatchSize              bigint    
, Directory              varchar(256)
, FilePattern            varchar(256)
, Worksheet              varchar(256)
, Header                 varchar(1)
, Delimiter              varchar(10)
, Filename               varchar(256)
, LoadDataDefKey         bigint        default 0 -- links back to LoadDataDef.LoadDataDefKey (implied foreign key but optional to populate)
, ErrorMsg               varchar(8000) default null
, constraint LoadDataOut_CK1
  check (Status in ('RUNNING', 'SUCCESS', 'FAILURE'))
)
;
