--drop table RCV_IHTKK_TKHAI_LOG
CREATE TABLE RCV_IHTKK_TKHAI_LOG
    (id_tkhai                      NUMBER,
    ngay_gui                       DATE,
    ngay_thuc_hien                 DATE,
    tbao_loi                       VARCHAR2(1000),
    tbao_loi_ctiet                  VARCHAR2(4000)
    )
  TABLESPACE  qlt_recv_data
    PARTITION BY RANGE (ngay_thuc_hien)
      (
       PARTITION rcv_chuyen_dlieu_xml_log_2013
       VALUES LESS THAN (TO_DATE(' 2014-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA,
       PARTITION rcv_chuyen_dlieu_xml_log_2014
       VALUES LESS THAN (TO_DATE(' 2015-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA,
       PARTITION rcv_chuyen_dlieu_xml_log_2015
       VALUES LESS THAN (TO_DATE(' 2016-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA,
       PARTITION rcv_chuyen_dlieu_xml_log_2016
       VALUES LESS THAN (TO_DATE(' 2017-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA,
       PARTITION rcv_chuyen_dlieu_xml_log_2017
       VALUES LESS THAN (TO_DATE(' 2018-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA,
       PARTITION rcv_chuyen_dlieu_xml_log_2018
       VALUES LESS THAN (TO_DATE(' 2019-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA,
       PARTITION rcv_chuyen_dlieu_xml_log_2019
       VALUES LESS THAN (TO_DATE(' 2020-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA,
       PARTITION rcv_chuyen_dlieu_xml_log_20
       VALUES LESS THAN (MAXVALUE)
       TABLESPACE  QLT_RECV_DATA
      )
/
CREATE INDEX RCV_IHTKK_TKHAI_LOG_id_ind ON RCV_IHTKK_TKHAI_LOG
(
   id_tkhai               ASC
)
TABLESPACE  qlt_recv_data
/
CREATE INDEX ngay_thuc_hien_log_ind ON RCV_IHTKK_TKHAI_LOG
(
  ngay_thuc_hien                          ASC
)
TABLESPACE  qlt_recv_data
/
