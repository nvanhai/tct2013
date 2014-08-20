spool C:\Temp\Log_NTK_326_DDL.txt
/
Whenever sqlerror continue
/
PROM '===== ALTER - CREATE ====='
Prompt 01_CREATE_SEQ_ESBMSG_ID.sql
@@01_CREATE_SEQ_ESBMSG_ID.sql;
Prompt 02_RCV_IHTKK_TKHAI.sql
@@02_RCV_IHTKK_TKHAI.sql;
Prompt 03_RCV_IHTKK_TKHAI_LOG.sql
@@03_RCV_IHTKK_TKHAI_LOG.sql;
Prompt 04_ALTER_IHTKK_TTIN_TTHAI.sql
@@04_ALTER_IHTKK_TTIN_TTHAI.sql;
Prompt 07_PCK_QLT_NTK_IHTKK.sql
@@07_PCK_QLT_NTK_IHTKK.sql;
Exit
/