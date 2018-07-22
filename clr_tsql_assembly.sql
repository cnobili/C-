drop assembly if exists clr_tsql;

create assembly clr_tsql
from 'c:\C#\CLR_TSQL\CLR_TSQL.dll'
WITH PERMISSION_SET = UNSAFE
