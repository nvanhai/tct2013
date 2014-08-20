CREATE SEQUENCE RCV_XLTK_ID_SEQ
  INCREMENT BY 1
  START WITH 1
  MINVALUE 1
  MAXVALUE 999999999999999999999999999
  NOCYCLE
  NOORDER
  
 /
-- drop table RCV_IHTKK_TKHAI
CREATE TABLE RCV_IHTKK_TKHAI
    (id                            NUMBER,
	tkhai_id                       VARCHAR2(500),
    ngay_gui                       DATE,
    ngay_nhan                      DATE,
    ma_tkhai                       VARCHAR2(3),
    ten_tkhai                      VARCHAR2(800),
    ky_kkhai                       VARCHAR2(10),
    dlieu_tkhai                    CLOB,
    ngay_nop                       DATE,
    lan_nop                        NUMBER(5),
    tin                            VARCHAR2(14),
    ten_NNT                        VARCHAR2(800),
    tin_dly                        VARCHAR2(14),
    ten_dly                        VARCHAR2(800),
	msg_id						   VARCHAR2(500),
	hthuc_nop					   VARCHAR2(1), --I :la ihtkk; N là từ NTK
    da_nhan                        VARCHAR2(1) default 'N',
	tthai_tms                     VARCHAR2(10),
	mota_loi_tms  VARCHAR2(1000)
    )
  TABLESPACE  qlt_recv_data
    PARTITION BY RANGE (ngay_nhan)
      (
       PARTITION rcv_chuyen_dlieu_xml_2013
       VALUES LESS THAN (TO_DATE(' 2014-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA
       LOB ("DLIEU_TKHAI") STORE AS (TABLESPACE  QLT_RECV_DATA),
       PARTITION rcv_chuyen_dlieu_xml_2014
       VALUES LESS THAN (TO_DATE(' 2015-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA
       LOB ("DLIEU_TKHAI") STORE AS (TABLESPACE  QLT_RECV_DATA),
       PARTITION rcv_chuyen_dlieu_xml_2015
       VALUES LESS THAN (TO_DATE(' 2016-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA
       LOB ("DLIEU_TKHAI") STORE AS (TABLESPACE  QLT_RECV_DATA),
       PARTITION rcv_chuyen_dlieu_xml_2016
       VALUES LESS THAN (TO_DATE(' 2017-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA
       LOB ("DLIEU_TKHAI") STORE AS (TABLESPACE  QLT_RECV_DATA),
       PARTITION rcv_chuyen_dlieu_xml_2017
       VALUES LESS THAN (TO_DATE(' 2018-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA
       LOB ("DLIEU_TKHAI") STORE AS (TABLESPACE  QLT_RECV_DATA),
       PARTITION rcv_chuyen_dlieu_xml_2018
       VALUES LESS THAN (TO_DATE(' 2019-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA
       LOB ("DLIEU_TKHAI") STORE AS (TABLESPACE  QLT_RECV_DATA),
       PARTITION rcv_chuyen_dlieu_xml_2019
       VALUES LESS THAN (TO_DATE(' 2020-01-01 00:00:00', 'SYYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))
       TABLESPACE  QLT_RECV_DATA
       LOB ("DLIEU_TKHAI") STORE AS (TABLESPACE  QLT_RECV_DATA),
       PARTITION rcv_chuyen_dlieu_xml_20
       VALUES LESS THAN (MAXVALUE)
       TABLESPACE  QLT_RECV_DATA
       LOB ("DLIEU_TKHAI") STORE AS (TABLESPACE  QLT_RECV_DATA)
      )
/
ALTER TABLE RCV_IHTKK_TKHAI
ADD CONSTRAINT RCV_IHTKK_TKHAI_pk PRIMARY KEY (id)
/
CREATE INDEX RCV_IHTKK_TKHAI_id_ind ON RCV_IHTKK_TKHAI
(
   tkhai_id               ASC
)
TABLESPACE  qlt_recv_data

/

CREATE INDEX RCV_IHTKK_TKHAI_ngay_gui_ind ON RCV_IHTKK_TKHAI
(
  ngay_gui                          ASC
)
TABLESPACE  qlt_recv_data
/
CREATE INDEX RCV_IHTKK_TKHAI_ngay_nhan_ind ON RCV_IHTKK_TKHAI
(
  ngay_nhan                          ASC
)
TABLESPACE  qlt_recv_data
/
CREATE INDEX RCV_IHTKK_TKHAI_ngay_nop_ind ON RCV_IHTKK_TKHAI
(
  ngay_nop                          ASC
)
TABLESPACE  qlt_recv_data
/
CREATE INDEX RCV_IHTKK_TKHAI_tin_ind ON RCV_IHTKK_TKHAI
(
  tin                         ASC
)
TABLESPACE  qlt_recv_data
/
CREATE INDEX RCV_IHTKK_TKHAI_tin_dly_ind ON RCV_IHTKK_TKHAI
(
  tin_dly                          ASC
)
TABLESPACE  qlt_recv_data
/
CREATE INDEX RCV_IHTKK_TKHAI_da_nhan_ind ON RCV_IHTKK_TKHAI
(
  da_nhan                       ASC
)
TABLESPACE  qlt_recv_data
/
CREATE INDEX RCV_IHTKK_TKHAI_group1_ind ON RCV_IHTKK_TKHAI
(
  tin                          ASC,
  da_nhan                       ASC
)
TABLESPACE  qlt_recv_data

/
CREATE INDEX RCV_IHTKK_TKHAI_msg_id_ind ON RCV_IHTKK_TKHAI
(
  msg_id                          ASC
)
TABLESPACE  qlt_recv_data

/
CREATE INDEX RCV_IHTKK_TKHAI_tthai_tms_ind ON RCV_IHTKK_TKHAI
(
  tthai_tms                          ASC
)
TABLESPACE  qlt_recv_data
/
 ALTER TABLE rcv_ihtkk_tkhai ADD CONSTRAINT dsach_tkhai_uk UNIQUE (tkhai_id, tin, ma_tkhai,ngay_nop ,ky_kkhai,
  lan_nop,hthuc_nop)
/
