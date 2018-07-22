if object_id('RunOsCmd') is not null
  drop procedure RunOsCmd
  
go

create proc RunOsCmd 
  @cmd nvarchar(4000) -- note that varchar does not match to a CLR data type
as
-- ---------------------------------------------------------------------------
--
-- Procedure:  RunOsCmd
--
-- purpose: T-SQL wrapper for CLR Procedure CLRTSQL.RunSystemCmd
-- 
-- Note: to run OS commands via CLR:
--
-- Assembly must be create with PERMISSION_SET = UNSAFE
-- And database must be set to trustworthy > ALTER DATABASE foobar SET TRUSTWORTHY ON
--
-- exec RunOsCmd @cmd = 'c:\C#\Csv2Table\Csv2Table c:\C#\Csv2Table\orders.dat , hey'
--
-- One can also use the xp_cmdshell command to run OS commands:
-- (database instance must allow T-SQL to execute os commands XPCmdShellEnable = True)
--
--   exec xp_cmdshell 'dir > c:\C#\x.log'
--   exec xp_cmdshell 'c:\C#\Csv2Table\Csv2Table c:\C#\Csv2Table\orders.dat , hey'
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
external name CLR_TSQL.CLRTSQL.RunCmdSync
