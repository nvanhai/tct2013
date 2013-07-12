CREATE OR REPLACE 
PACKAGE rcv_pck_job is
procedure Prc_Check_Job(job_name varchar2);
procedure Prc_Run_Job(job_name varchar2);
procedure Prc_Remove_Job(job_name varchar2);
Procedure Prc_ReCreate_Job;
End rcv_pck_job;
/

CREATE OR REPLACE 
PACKAGE BODY rcv_pck_job IS
Procedure Prc_Check_Job(job_name varchar2) is
Cursor c_Job Is
    Select Broken from user_jobs
    Where what = job_name;
    v_Broken  Varchar2(10);
BEGIN
    Open c_Job;
    Fetch c_Job into v_Broken;
    If c_job%NOTFOUND then
      Prc_Run_Job(job_name);
      Close c_Job;
      return;
    End If;
    Close c_Job;

    If v_Broken = 'Y' then
      Prc_Remove_Job(job_name);
      Prc_Run_Job(job_name);
    End If;
END;


procedure Prc_Run_Job(job_name varchar2) is
Cursor c_Job_Exist is
    Select 1 from User_Jobs
    Where upper(what) = upper(job_name);

    v_Intvl Varchar2(150);
    v_IntvlNum Number;
    V_num number;
    v_Sql Varchar2(1000);
    v_Exist  Integer;
BEGIN
    Open c_Job_Exist;
    Fetch c_Job_Exist into v_Exist;
    If c_Job_Exist%FOUND then
      Close c_Job_Exist;
      return;
    End If;
    Close c_Job_Exist;
    v_Intvl := 'SYSDATE+5/(60*24)';
    v_Sql := 'ALTER SESSION SET NLS_DATE_FORMAT = "DD/MM/YYYY"';
    Execute Immediate v_Sql;
	DBMS_JOB.SUBMIT(V_num,job_name,SYSDATE,v_Intvl);
	Commit;
Exception
  When Others Then
	Rollback;
END;
---------------------------
procedure Prc_Remove_Job(job_name varchar2) is
V_num number;
Cursor C_Job is
    Select Job from USER_Jobs
    Where (upper(PRIV_USER)='QLT_RECV')
    And what = job_name;

	V_job C_Job%rowtype;
BEGIN
	For V_job In C_Job loop
		DBMS_JOB.REMOVE(V_Job.Job);
	End Loop;
    Commit;
Exception
	When Others Then
		Rollback;
END;
---------------------------
Procedure Prc_ReCreate_Job is
    Cursor C_Job is
    Select Job, what, interval from USER_Jobs
    Where upper(PRIV_USER)='QLT_RECV'
    And broken = 'Y';
	V_job C_Job%rowtype;
	v_jobid number;
	v_retry boolean;
	v_count number(1);
BEGIN
	For V_job In C_Job loop
	v_retry := true;
	while (v_retry) loop
	begin
		DBMS_JOB.REMOVE(V_Job.Job);
		Select count(1) into v_count from USER_Jobs where Job = V_Job.Job;
		If v_count = 0 Then
			v_retry := false;
		End If;
	end;
	end loop;
		dbms_job.submit(v_jobid,v_job.what, sysdate,v_job.interval);
	End Loop;
    Commit;
Exception
	When Others Then
		Rollback;
end;
END;
/
