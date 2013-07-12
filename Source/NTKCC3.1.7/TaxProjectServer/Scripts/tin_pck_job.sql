-- Start of DDL Script for Package Body QLT_OWNER.TIN_PCK_JOB
-- Generated 8-Dec-2005 18:51:59 from QLT_OWNER@QLT_93

CREATE OR REPLACE 
PACKAGE tin_pck_job is
procedure Prc_Check_Job;
procedure Prc_Run_Job;
procedure Prc_Remove_Job;
Procedure Prc_ReCreate_Job;
End TIN_PCK_JOB;
/


CREATE OR REPLACE 
PACKAGE BODY tin_pck_job
IS
----------------------------
   PROCEDURE prc_check_job
   IS
/******************************************************************************
Author: ThanhMT
Purpose: Check Job status, rerun the job if it is broken
Create date: 11/11/2003
******************************************************************************/
      CURSOR c_job
      IS
         SELECT broken
           FROM user_jobs
          WHERE schema_user = 'QLT_OWNER';

      v_broken   VARCHAR2 (10);
   BEGIN
      OPEN c_job;

      FETCH c_job
       INTO v_broken;

      IF c_job%NOTFOUND
      THEN
         prc_run_job;

         CLOSE c_job;

         RETURN;
      END IF;

      CLOSE c_job;

      IF v_broken = 'Y'
      THEN
         prc_remove_job;
         prc_run_job;
      END IF;
   END;

------------------------------------------------------------------------------
   PROCEDURE prc_run_job
   IS
/******************************************************************************
Purpose: Run Job receive, send file in TCT
******************************************************************************/
      CURSOR c_job_exist
      IS
         SELECT 1
           FROM user_jobs
          WHERE schema_user = 'QLT_OWNER';

      v_intvl      VARCHAR2 (50);
      v_intvlnum   NUMBER;
      v_num        NUMBER;
      v_sql        VARCHAR2 (1000);
      v_exist      INTEGER;
   BEGIN
      OPEN c_job_exist;

      FETCH c_job_exist
       INTO v_exist;

      IF c_job_exist%FOUND
      THEN
         CLOSE c_job_exist;

         RETURN;
      END IF;

      CLOSE c_job_exist;

      v_intvl := '60';                                        --default 1 phut
      v_sql := 'ALTER SESSION SET NLS_DATE_FORMAT = "DD/MM/YYYY"';

      EXECUTE IMMEDIATE v_sql;

      DBMS_JOB.submit (v_num,
                       'RCV_PCK_CHUYEN_DLIEU_QLT.prc_chuyen_dlieu_qlt;',
                       SYSDATE,
                       'SYSDATE+' || v_intvl || '/(60*60*24)'
                      );
      COMMIT;
   EXCEPTION
      WHEN OTHERS
      THEN
         ROLLBACK;
   END;

----------------------------
   PROCEDURE prc_remove_job
   IS
/******************************************************************************
Purpose: Remove Recv_Send_Job that is running in Center or the Job whose name specified by P_job_name
******************************************************************************/
      v_num   NUMBER;

      CURSOR c_job
      IS
         SELECT job
           FROM user_jobs
          WHERE (UPPER (priv_user) = 'QLT_OWNER')
            AND schema_user = 'QLT_OWNER';

      v_job   c_job%ROWTYPE;
   BEGIN
      FOR v_job IN c_job
      LOOP
         DBMS_JOB.remove (v_job.job);
      END LOOP;

      COMMIT;
   EXCEPTION
      WHEN OTHERS
      THEN
         ROLLBACK;
   END;

---------------------------
   PROCEDURE prc_recreate_job
   IS
      CURSOR c_job
      IS
         SELECT job, what, INTERVAL
           FROM user_jobs
          WHERE broken = 'Y';

      v_job     c_job%ROWTYPE;
      v_jobid   NUMBER;
   BEGIN
      FOR v_job IN c_job
      LOOP
         DBMS_JOB.remove (v_job.job);
         DBMS_JOB.submit (v_jobid, v_job.what, SYSDATE, v_job.INTERVAL);
      END LOOP;
      COMMIT;
   EXCEPTION
      WHEN OTHERS
      THEN
         ROLLBACK;
   END;

END tin_pck_job;
/


-- End of DDL Script for Package Body QLT_OWNER.TIN_PCK_JOB

