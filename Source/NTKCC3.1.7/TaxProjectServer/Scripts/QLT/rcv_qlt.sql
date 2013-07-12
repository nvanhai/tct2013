--Connect bang user QLT_RECV
SPOOL c:\temp\logrcv_qlt120.txt

Prom '-----Phan TABLES ---'
Prom 'rcv_map_tkhai'
@@rcv_map_tkhai.sql
Prom 'rcv_map_tkhai_data'
@@rcv_map_tkhai_data.sql
Prom 'rcv_dm_tkhai'
@@rcv_dm_tkhai.sql
Prom 'rcv_dm_tkhai_data'
@@rcv_dm_tkhai_data.sql
Prom 'rcv_gdien_tkhai'
@@rcv_gdien_tkhai.sql
Prom 'rcv_gdien_tkhai_data'
@@rcv_gdien_tkhai_data.sql
Prom 'rcv_map_ctieu'
@@rcv_map_ctieu.sql
Prom 'rcv_map_ctieu_data'
@@rcv_map_ctieu_data.sql
Prom 'rcv_tkhai_hdr'
@@rcv_tkhai_hdr.sql
Prom 'rcv_tkhai_dtl'
@@rcv_tkhai_dtl.sql
Prom 'rcv_thamso'
@@rcv_thamso.sql
Prom 'rcv_thamso_data'
@@rcv_thamso_data.sql
Prom 'rcv_map_ctieu_bthue'
@@rcv_map_ctieu_bthue.sql
Prom 'rcv_map_ctieu_bthue_data'
@@rcv_map_ctieu_bthue_data.sql

Prom '----Phan SEQUENCES---- '
Prom 'rcv_xltk_hdr_seq'
@@rcv_xltk_hdr_seq.sql
Prom 'rcv_xltk_dtl_seq'
@@rcv_xltk_dtl_seq.sql

Prom '--- Phan VIEWS ----'
Prom 'rcv_v_tkhai_gtgt_kt'
@@rcv_v_tkhai_gtgt_kt.sql
Prom 'rcv_v_tkhai_gtgt_kt_pluc2a'
@@rcv_v_tkhai_gtgt_kt_pluc2a.sql
Prom 'rcv_v_tkhai_gtgt_kt_pluc2b'
@@rcv_v_tkhai_gtgt_kt_pluc2b.sql
Prom 'rcv_v_tkhai_gtgt_kt_pluc2c'
@@rcv_v_tkhai_gtgt_kt_pluc2c.sql
Prom 'rcv_v_dtnt'
@@rcv_v_dtnt.sql
Prom 'rcv_v_tkhai_tndn_quy'
@@rcv_v_tkhai_tndn_quy.sql
Prom 'rcv_v_tkhai_qtoan_tndn'
@@rcv_v_tkhai_qtoan_tndn.sql
Prom 'rcv_v_pluc_qtoan_tndn_01ab'
@@rcv_v_pluc_qtoan_tndn_01ab.sql
Prom 'rcv_v_pluc_qtoan_tndn_02_13'
@@rcv_v_pluc_qtoan_tndn_02_13.sql
Prom 'rcv_v_pluc_qtoan_tndn_14'
@@rcv_v_pluc_qtoan_tndn_14.sql
Prom 'rcv_v_tkhai_gtgt_tt'
@@rcv_v_tkhai_gtgt_tt.sql
Prom 'rcv_v_tkhai_qtoan_gtgt_tt'
@@rcv_v_tkhai_qtoan_gtgt_tt.sql
Prom 'rcv_v_tkhai_tnguyen'
@@rcv_v_tkhai_tnguyen.sql
Prom 'rcv_v_tkhai_qtoan_tnguyen'
@@rcv_v_tkhai_qtoan_tnguyen.sql
Prom 'rcv_v_tkhai_ttdb'
@@rcv_v_tkhai_ttdb.sql
Prom 'rcv_v_pluc_ttdb_01c'
@@rcv_v_pluc_ttdb_01c.sql

Prom '---Gan quyen va tao synonym---'
Prom 'Synonym_And_Grant'
@@Synonym_And_Grant.sql

Prom '--- Phan PACKAGES ----'
Prom 'rcv_pck_chuyen_dlieu_qlt'
@@rcv_pck_chuyen_dlieu_qlt.sql
Prom 'rcv_pck_job'
@@rcv_pck_job.sql

Prom '--- CREATE JOB ----'
Prom 'rcv_create_job'
@@rcv_create_job.sql
Prom 'rcv_recreate_job'
@@rcv_recreate_job.sql

Prom ' KET THUC DATABASE'
commit
/
exit
