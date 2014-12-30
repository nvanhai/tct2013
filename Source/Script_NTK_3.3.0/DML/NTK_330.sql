Spool C:\Temp\Log_NTK_330_DML.txt
/
Whenever sqlerror continue
/
PROM '===== INSERT - UPDATE ====='
Prompt 1_rcv_dm_tkhai.sql
@@1_rcv_dm_tkhai.sql;
Prompt 2_rcv_map_tkhai.sql
@@2_rcv_map_tkhai.sql;
Prompt 3_rcv_gdien_tkhai.sql
@@3_rcv_gdien_tkhai.sql;
Prompt 4_rcv_map_ctieu.sql
@@4_rcv_map_ctieu.sql;
Prompt 5_update_cg_reg_code.sql
@@5_update_cg_reg_code.sql;
/
Exit
/
spool off;