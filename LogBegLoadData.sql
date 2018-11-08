if object_id('LogBegLoadData') is not null
  drop procedure LogBegLoadData
  
go

create procedure LogBegLoadData
(
  @ProgramName       varchar(256)
, @LoadDataName      varchar(256)  
, @SqlStatement      varchar(8000)
, @BegDateTime       datetime
, @DestConnStr       varchar(256)
, @DestDbName        varchar(256)
, @DestSchemaName    varchar(256)
, @DestTableName     varchar(256)
, @LoadType          varchar(256)
, @BatchSize         bigint
, @Directory         varchar(256)
, @FilePattern       varchar(256)
, @Worksheet         varchar(256)
, @Header            varchar(1)
, @Delimiter         varchar(10)
, @Filename          varchar(256)
, @LoadDataDefKey    bigint = 0
)
-- -----------------------------------------------------------------------------
--
-- procedure: LogBegLoadData
--
-- purpose: Logs the beginning of a data load.
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
as
begin

  insert into LoadDataOut
  (
    LoadDataName
  , ProgramName
  , SqlStatement
  , BegDateTime
  , Status
  , LoadDataDefKey  
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
  , Filename
  )
  output inserted.LoadDataOutKey
  values
  (
    @LoadDataName
  , @ProgramName    
  , @SqlStatement
  , @BegDateTime
  , 'RUNNING'
  , @LoadDataDefKey
  , @DestConnStr
  , @DestDbName 
  , @DestSchemaName
  , @DestTableName 
  , @LoadType      
  , @BatchSize
  , @Directory
  , @FilePattern  
  , @Worksheet     
  , @Header        
  , @Delimiter     
  , @Filename
  );
  
end
