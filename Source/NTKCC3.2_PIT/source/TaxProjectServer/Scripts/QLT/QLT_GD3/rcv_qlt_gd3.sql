--Connect bang user QLT_RECV
SPOOL c:\temp\logrcv_gd3.txt

Prom '-----Phan TABLES ---'
Prom 'alter_rcv_tkhai_hdr'
@@alter_rcv_tkhai_hdr.sql
Prom 'rcv_trg_auid_tkhai_hdr'
@@rcv_trg_auid_tkhai_hdr.sql
Prom 'rcv_tkhai_dtl_ind'
@@rcv_tkhai_dtl_ind.sql
Prom 'alter_rcv_map_tkhai'
@@alter_rcv_map_tkhai.sql
Prom 'rcv_map_tkhai_data'
@@rcv_map_tkhai_data.sql
Prom 'rcv_dm_tkhai_alter'
@@rcv_dm_tkhai_alter.sql
Prom 'rcv_dm_tkhai_data'
@@rcv_dm_tkhai_data.sql
Prom 'rcv_gdien_tkhai_data'
@@rcv_gdien_tkhai_data.sql
Prom 'rcv_map_ctieu_data'
@@rcv_map_ctieu_data.sql
prom 'rcv_map_ctieu_bthue'
@@rcv_map_ctieu_bthue.sql
prom 'rcv_map_ctieu_bthue_data'
@@rcv_map_ctieu_bthue_data.sql
Prom 'rcv_thamso_data'
@@rcv_thamso_data.sql
Prom 'Synonym_And_Grant'
@@Synonym_And_Grant.sql

Prom '--- Phan VIEWS ----'
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

Prom '--- Phan PACKAGES ----'
Prom 'rcv_pck_chuyen_dlieu_qlt'
@@rcv_pck_chuyen_dlieu_qlt.sql

Prom ' KET THUC DATABASE'
commit
/
exit
