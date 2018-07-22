if object_id('IsMatch') is not null
  drop function IsMatch
  
go

create function IsMatch
(
  @str nvarchar(4000)  -- note that varchar does not match to a CLR data type
, @regex nvarchar(256) -- note that varchar does not match to a CLR data type  
) 
returns int
as
-- ---------------------------------------------------------------------------
--
-- Procedure:  IsMatch
--
-- purpose: T-SQL wrapper for CLR Function StoredProcedures.IsMatch
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
external name CLR_TSQL.CLRTSQL.IsMatch

