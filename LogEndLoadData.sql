if object_id('LogEndLoadData') is not null
  drop procedure LogEndLoadData
  
go

create procedure LogEndLoadData
(
  @LoadDataOutKey    bigint
, @EndDateTime       datetime
, @Status            varchar(20) -- Success or Failure
, @NumRecs           bigint
, @ErrorMsg          varchar(8000) = null
)
-- -----------------------------------------------------------------------------
--
-- procedure: LogEndLoadData
--
-- purpose: Logs the end of a data load.
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

  update LoadDataOut set
    EndDateTime   = GetDate()
  , Status        = @Status
  , NumRecs       = @NumRecs
  , ErrorMsg      = @ErrorMsg
  where LoadDataOutKey = @LoadDataOutKey
  ;

end
