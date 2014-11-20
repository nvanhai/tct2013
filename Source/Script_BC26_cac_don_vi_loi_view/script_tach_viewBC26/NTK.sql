spool C:\Temp\Log_NTK_DDL.txt
/
Whenever sqlerror continue
/
Prompt RCV_V_BC26_AC_t1.sql
@@RCV_V_BC26_AC_t1.sql;
Prompt RCV_V_BC26_AC_t2.sql
@@RCV_V_BC26_AC_t2.sql;
Prompt RCV_V_BC26_AC.sql
@@RCV_V_BC26_AC.sql;

Prompt rcv_v_bc26_ac_blp_t1.sql
@@rcv_v_bc26_ac_blp_t1.sql;

Prompt rcv_v_bc26_ac_blp_t2.sql
@@rcv_v_bc26_ac_blp_t2.sql;

Prompt rcv_v_bc26_ac_blp.sql
@@rcv_v_bc26_ac_blp.sql;

Exit
/