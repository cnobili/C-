if object_id('SqlExtractData') is not null
  drop procedure SqlExtractData
  
go
create proc SqlExtractData
  @outputfile   nvarchar(256) -- note that varchar does not match to a CLR data type
, @delim        nvarchar(10)
, @columnHeader nvarchar(1)
, @query        nvarchar(max)
, @server       nvarchar(256)
, @dbname       nvarchar(256)
, @user         nvarchar(256)
, @pass         nvarchar(256)
as
-- ---------------------------------------------------------------------------
--
-- Procedure:  SqlExtractData
--
-- purpose: T-SQL wrapper for CLR Procedure StoredProcedures.SqlExtractData
--
-- exec SqlExtractData
--   @outputfile   = 'c:\temp\x.txt'
-- , @delim        = '|'
-- , @columnHeader = 'T'
-- , @query        = 'select * from Epic.Patient.Patient'
-- , @server       = 'HCDBDEV'
-- , @dbname       = 'Epic'
-- , @user         = null
-- , @pass         = null
--
-- If user/pass is null, uses trusted connection.
--
-- Note: If uses trusted connection, the server can only be local to where this
--   procedure is running on.
--
-- ---------------------------------------------------------------------------
--
-- rev log
--
-- date:  15-MAR-2018
-- author:  Craig Nobili
-- desc: original
--
-- ---------------------------------------------------------------------------  
external name CLR_TSQL.CLRTSQL.SqlExtractData
