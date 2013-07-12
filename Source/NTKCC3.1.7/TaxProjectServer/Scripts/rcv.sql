-- Tao bang

-- Bang RCV_MAP_TKHAI
CREATE TABLE rcv_map_tkhai
    (ma_tkhai                       VARCHAR2(2) NOT NULL,
    ma_tkhai_qlt                   VARCHAR2(2) NOT NULL,
    ghi_chu                        VARCHAR2(100))
  PCTFREE     10
  PCTUSED     40
  INITRANS    1
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

CREATE PUBLIC SYNONYM rcv_map_tkhai
  FOR rcv_map_tkhai
/

GRANT DELETE ON rcv_map_tkhai TO qlt
/
GRANT INSERT ON rcv_map_tkhai TO qlt
/
GRANT SELECT ON rcv_map_tkhai TO qlt
/
GRANT UPDATE ON rcv_map_tkhai TO qlt
/
GRANT SELECT ON rcv_map_tkhai TO qlt_read
/

ALTER TABLE rcv_map_tkhai
ADD CONSTRAINT rcv_map_tkhai_pk PRIMARY KEY (ma_tkhai)
USING INDEX
  PCTFREE     10
  INITRANS    2
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

-- Bang RCV_DM_TKHAI
CREATE TABLE rcv_dm_tkhai
    (ma                             VARCHAR2(2) NOT NULL,
    ten                            VARCHAR2(60) NOT NULL)
  PCTFREE     5
  PCTUSED     90
  INITRANS    1
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

CREATE PUBLIC SYNONYM rcv_dm_tkhai
  FOR rcv_dm_tkhai
/

GRANT DELETE ON rcv_dm_tkhai TO qlt
/
GRANT INSERT ON rcv_dm_tkhai TO qlt
/
GRANT SELECT ON rcv_dm_tkhai TO qlt
/
GRANT UPDATE ON rcv_dm_tkhai TO qlt
/
GRANT SELECT ON rcv_dm_tkhai TO qlt_read
/

ALTER TABLE rcv_dm_tkhai
ADD CONSTRAINT rcv_dtk_pk PRIMARY KEY (ma)
USING INDEX
  PCTFREE     5
  INITRANS    2
  MAXTRANS    255
  TABLESPACE  qlt_indexes
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

-- Bang RCV_GDIEN_TKHAI
CREATE TABLE rcv_gdien_tkhai
    (id                             NUMBER(10,0) NOT NULL,
    ten_ctieu                      VARCHAR2(200) NOT NULL,
    cot_01                         VARCHAR2(10),
    cot_02                         VARCHAR2(10),
    cot_03                         VARCHAR2(10),
    cot_04                         VARCHAR2(10),
    cot_05                         VARCHAR2(10),
    cot_06                         VARCHAR2(10),
    cot_07                         VARCHAR2(10),
    cot_08                         VARCHAR2(10),
    cot_09                         VARCHAR2(10),
    cot_10                         VARCHAR2(10),
    so_tt                          NUMBER(3,0),
    loai_dlieu                     VARCHAR2(4),
    ma_ctieu                       VARCHAR2(3))
  PCTFREE     5
  PCTUSED     90
  INITRANS    1
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

CREATE PUBLIC SYNONYM rcv_gdien_tkhai
  FOR rcv_gdien_tkhai
/

GRANT DELETE ON rcv_gdien_tkhai TO qlt
/
GRANT INSERT ON rcv_gdien_tkhai TO qlt
/
GRANT SELECT ON rcv_gdien_tkhai TO qlt
/
GRANT UPDATE ON rcv_gdien_tkhai TO qlt
/
GRANT SELECT ON rcv_gdien_tkhai TO qlt_read
/

ALTER TABLE rcv_gdien_tkhai
ADD CONSTRAINT rcv_gtk_pk PRIMARY KEY (id)
USING INDEX
  PCTFREE     5
  INITRANS    2
  MAXTRANS    255
  TABLESPACE  qlt_indexes
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

-- Bang RCV_MAP_CTIEU
CREATE TABLE rcv_map_ctieu
    (loai_dlieu                     VARCHAR2(4) NOT NULL,
    ky_hieu                        VARCHAR2(10) NOT NULL,
    kieu_dlieu                     VARCHAR2(1) NOT NULL,
    gdn_id                         NUMBER(10,0) NOT NULL,
    ky_hieu_ctieu                  VARCHAR2(4))
  PCTFREE     5
  PCTUSED     90
  INITRANS    1
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

CREATE PUBLIC SYNONYM rcv_map_ctieu
  FOR rcv_map_ctieu
/

GRANT DELETE ON rcv_map_ctieu TO qlt
/
GRANT INSERT ON rcv_map_ctieu TO qlt
/
GRANT SELECT ON rcv_map_ctieu TO qlt
/
GRANT UPDATE ON rcv_map_ctieu TO qlt
/
GRANT SELECT ON rcv_map_ctieu TO qlt_read
/

ALTER TABLE rcv_map_ctieu
ADD CONSTRAINT rcv_mctieu_uk UNIQUE (loai_dlieu, ky_hieu)
USING INDEX
  PCTFREE     10
  INITRANS    2
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

ALTER TABLE rcv_map_ctieu
ADD CONSTRAINT rcv_mctieu_fk FOREIGN KEY (gdn_id)
REFERENCES rcv_gdien_tkhai (id)
/

-- Bang RCV_TKHAI_HDR
CREATE TABLE rcv_tkhai_hdr
    (id                             NUMBER(10,0) NOT NULL,
    tin                            VARCHAR2(14) NOT NULL,
    ten_dtnt                       VARCHAR2(60) NOT NULL,
    dia_chi                        VARCHAR2(60),
    loai_tkhai                     VARCHAR2(2) NOT NULL,
    ngay_nop                       DATE NOT NULL,
    kylb_tu_ngay                   DATE NOT NULL,
    kylb_den_ngay                  DATE NOT NULL,
    kykk_tu_ngay                   DATE NOT NULL,
    kykk_den_ngay                  DATE NOT NULL,
    ngay_cap_nhat                  DATE NOT NULL,
    nguoi_cap_nhat                 VARCHAR2(60) NOT NULL,
    co_loi_ddanh                   CHAR(1),
    so_hieu_tep                    VARCHAR2(20),
    so_tt_tk                       NUMBER(10,0),
    da_nhan                        CHAR(1),
    ghi_chu_loi                    VARCHAR2(100),
    co_gtrinh_02a                  CHAR(1),
    co_gtrinh_02b                  CHAR(1),
    co_gtrinh_02c                  CHAR(1))
  PCTFREE     10
  PCTUSED     40
  INITRANS    1
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

CREATE PUBLIC SYNONYM rcv_tkhai_hdr
  FOR rcv_tkhai_hdr
/

GRANT DELETE ON rcv_tkhai_hdr TO qlt
/
GRANT INSERT ON rcv_tkhai_hdr TO qlt
/
GRANT SELECT ON rcv_tkhai_hdr TO qlt
/
GRANT UPDATE ON rcv_tkhai_hdr TO qlt
/
GRANT SELECT ON rcv_tkhai_hdr TO qlt_read
/

ALTER TABLE rcv_tkhai_hdr
ADD CONSTRAINT rcv_tkh_pk PRIMARY KEY (id)
USING INDEX
  PCTFREE     10
  INITRANS    2
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

ALTER TABLE rcv_tkhai_hdr
ADD CONSTRAINT rcv_tkh_fk FOREIGN KEY (loai_tkhai)
REFERENCES rcv_dm_tkhai (ma)
/

-- Bang RCV_TKHAI_DTL
CREATE TABLE rcv_tkhai_dtl
    (id                             NUMBER(10,0) NOT NULL,
    hdr_id                         NUMBER(10,0) NOT NULL,
    loai_dlieu                     VARCHAR2(4) NOT NULL,
    ky_hieu                        VARCHAR2(10) NOT NULL,
    gia_tri                        VARCHAR2(1000),
    row_id                         NUMBER(10,0))
  PCTFREE     10
  PCTUSED     40
  INITRANS    1
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

CREATE PUBLIC SYNONYM rcv_tkhai_dtl
  FOR rcv_tkhai_dtl
/

GRANT DELETE ON rcv_tkhai_dtl TO qlt
/
GRANT INSERT ON rcv_tkhai_dtl TO qlt
/
GRANT SELECT ON rcv_tkhai_dtl TO qlt
/
GRANT UPDATE ON rcv_tkhai_dtl TO qlt
/
GRANT SELECT ON rcv_tkhai_dtl TO qlt_read
/

ALTER TABLE rcv_tkhai_dtl
ADD CONSTRAINT rcv_tkd_pk PRIMARY KEY (id)
USING INDEX
  PCTFREE     10
  INITRANS    2
  MAXTRANS    255
  TABLESPACE  qlt_dmuc
  STORAGE   (
    INITIAL     65536
    MINEXTENTS  1
    MAXEXTENTS  2147483645
  )
/

ALTER TABLE rcv_tkhai_dtl
ADD CONSTRAINT rcv_tkd_hdr_fk FOREIGN KEY (hdr_id)
REFERENCES rcv_tkhai_hdr (id) ON DELETE CASCADE
/
ALTER TABLE rcv_tkhai_dtl
ADD CONSTRAINT rcv_tkd_mctieu_fk FOREIGN KEY (loai_dlieu, ky_hieu)
REFERENCES rcv_map_ctieu (loai_dlieu,ky_hieu)
/

-- Tao View

-- View RCV_V_TKHAI_GTGT_KT
CREATE OR REPLACE VIEW rcv_v_tkhai_gtgt_kt (
   hdr_id,
   ctk_id,
   so_tt,
   ten_ctieu,
   doanhso_dtnt,
   sothue_dtnt,
   kieu_dlieu_ds,
   kieu_dlieu_st,
   ky_hieu_ctieu_ds,
   ky_hieu_ctieu_st )
AS
SELECT dtl.hdr_id
     , dtl.ctk_id
     , MAX(dtl.so_tt) so_tt
     , MAX(gd.ten_ctieu) ten_ctieu
     , MAX(dtl.doanhso_dtnt) doanhso_dtnt
     , MAX(dtl.sothue_dtnt) sothue_dtnt
     , MAX(dtl.kieu_dlieu_ds) kieu_dlieu_ds
     , MAX(dtl.kieu_dlieu_st) kieu_dlieu_st
     , MAX(dtl.ky_hieu_ctieu_ds) ky_hieu_ctieu_ds
     , MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id hdr_id,
         gdien.id id,
         gdien.so_tt so_tt,
         tkd.row_id row_id,
         gdien.ma_ctieu ctk_id,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) doanhso_dtnt,
    	 DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) sothue_dtnt,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ds,
    	 DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_st,
         DECODE(gdien.cot_01, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL) ky_hieu_ctieu_ds,
         DECODE(gdien.cot_02, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL) ky_hieu_ctieu_st
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0101')
	
) dtl
WHERE (gd.loai_dlieu = '0101')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         dtl.ctk_id
/

CREATE PUBLIC SYNONYM rcv_v_tkhai_gtgt_kt
  FOR rcv_v_tkhai_gtgt_kt
/

GRANT SELECT ON rcv_v_tkhai_gtgt_kt TO qlt
/
GRANT SELECT ON rcv_v_tkhai_gtgt_kt TO qlt_read
/

-- View RCV_V_TKHAI_GTGT_KT_PLUC2A
CREATE OR REPLACE VIEW rcv_v_tkhai_gtgt_kt_pluc2a (
   hdr_id,
   ctk_id,
   dien_giai,
   ky_hieu,
   gia_tri_ky_kkhai,
   gia_tri_slieu_kkhai,
   gia_tri_slieu_dchinh,
   gia_tri_hhdv,
   gia_tri_thue_gtgt,
   gia_tri_lydo_dchinh,
   kieu_dlieu_dien_giai,
   kieu_dlieu_ma_ctieu,
   kieu_dlieu_kykk,
   kieu_dlieu_slieu_kkhai,
   kieu_dlieu_slieu_dchinh,
   kieu_dlieu_hhdv,
   kieu_dlieu_thue_gtgt,
   kieu_dlieu_lydo_dchinh )
AS
SELECT dtl.hdr_id
     , gd.ma_ctieu ctk_id
     , MAX(dtl.gia_tri_dien_giai) dien_giai
     , MAX(dtl.gia_tri_ma_ctieu) ky_hieu
     , MAX(dtl.gia_tri_ky_kkhai) gia_tri_ky_kkhai
     , MAX(dtl.gia_tri_slieu_kkhai) gia_tri_slieu_kkhai
     , MAX(dtl.gia_tri_slieu_dchinh) gia_tri_slieu_dchinh
     , MAX(dtl.gia_tri_hhdv) gia_tri_hhdv
     , MAX(dtl.gia_tri_thue_gtgt) gia_tri_thue_gtgt
     , MAX(dtl.gia_tri_lydo_dchinh) gia_tri_lydo_dchinh
     , MAX(dtl.kieu_dlieu_dien_giai) kieu_dlieu_dien_giai
     , MAX(dtl.kieu_dlieu_ma_ctieu) kieu_dlieu_ma_ctieu
     , MAX(dtl.kieu_dlieu_kykk) kieu_dlieu_kykk
     , MAX(dtl.kieu_dlieu_slieu_kkhai) kieu_dlieu_slieu_kkhai
     , MAX(dtl.kieu_dlieu_slieu_dchinh) kieu_dlieu_slieu_dchinh
     , MAX(dtl.kieu_dlieu_hhdv) kieu_dlieu_hhdv
     , MAX(dtl.kieu_dlieu_thue_gtgt) kieu_dlieu_thue_gtgt
     , MAX(dtl.kieu_dlieu_lydo_dchinh) kieu_dlieu_lydo_dchinh
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         tkd.row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_dien_giai,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_ma_ctieu,
    	 DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_ky_kkhai,
    	 DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_slieu_kkhai,
    	 DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_slieu_dchinh,
    	 DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_hhdv,
    	 DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_thue_gtgt,
    	 DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_lydo_dchinh,
    	 DECODE(gdien.cot_08, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_dien_giai,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ma_ctieu,
    	 DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_kykk,
    	 DECODE(gdien.cot_03, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_slieu_kkhai,
    	 DECODE(gdien.cot_04, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_slieu_dchinh,
    	 DECODE(gdien.cot_05, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_hhdv,
    	 DECODE(gdien.cot_06, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_thue_gtgt,
    	 DECODE(gdien.cot_07, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_lydo_dchinh
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0102')
	
) dtl
WHERE (gd.loai_dlieu = '0102')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         gd.ma_ctieu
/

CREATE PUBLIC SYNONYM rcv_v_tkhai_gtgt_kt_pluc2a
  FOR rcv_v_tkhai_gtgt_kt_pluc2a
/

GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2a TO qlt
/
GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2a TO qlt_read
/

-- View RCV_V_TKHAI_GTGT_KT_PLUC2B
CREATE OR REPLACE VIEW rcv_v_tkhai_gtgt_kt_pluc2b (
   hdr_id,
   ctg_id,
   ten_ctieu,
   gia_tri_ctieu,
   kieu_dlieu_ctieu )
AS
SELECT dtl.hdr_id
     , MAX(dtl.ctg_id) ctg_id
     , gd.ten_ctieu
     , MAX(dtl.gia_tri_ctieu) gia_tri_ctieu
     , MAX(dtl.kieu_dlieu_ctieu) kieu_dlieu_ctieu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_ctieu,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ctieu,
    	 DECODE(gdien.cot_01, tkd.ky_hieu, gdien.ma_ctieu, NULL) ctg_id,
         gdien.id,
         tkd.row_id
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0103')
    AND (gdien.ma_ctieu IS NOT NULL)
) dtl
WHERE (gd.loai_dlieu = '0103')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         gd.ten_ctieu
/

CREATE PUBLIC SYNONYM rcv_v_tkhai_gtgt_kt_pluc2b
  FOR rcv_v_tkhai_gtgt_kt_pluc2b
/

GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2b TO qlt
/
GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2b TO qlt_read
/

-- View RCV_V_TKHAI_GTGT_KT_PLUC2C
CREATE OR REPLACE VIEW rcv_v_tkhai_gtgt_kt_pluc2c (
   hdr_id,
   ctg_id,
   ten_ctieu,
   gia_tri_ctieu,
   kieu_dlieu_ctieu )
AS
SELECT dtl.hdr_id
     , MAX(dtl.ctg_id) ctg_id
     , gd.ten_ctieu
     , MAX(dtl.gia_tri_ctieu) gia_tri_ctieu
     , MAX(dtl.kieu_dlieu_ctieu) kieu_dlieu_ctieu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_ctieu,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ctieu,
    	 DECODE(gdien.cot_01, tkd.ky_hieu, gdien.ma_ctieu, NULL) ctg_id,
         gdien.id,
         tkd.row_id
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0104')
    AND (gdien.ma_ctieu IS NOT NULL)
) dtl
WHERE (gd.loai_dlieu = '0104')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         gd.ten_ctieu
/

CREATE PUBLIC SYNONYM rcv_v_tkhai_gtgt_kt_pluc2c
  FOR rcv_v_tkhai_gtgt_kt_pluc2c
/

GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2c TO qlt
/
GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2c TO qlt_read
/


-- Insert du lieu

-- Du lieu cua bang RCV_DM_TKHAI
INSERT INTO rcv_dm_tkhai
VALUES
('01','Tê khai thuÕ gi¸ trÞ gia t¨ng khÊu trõ')
/
INSERT INTO rcv_dm_tkhai
VALUES
('02','B¶n x¸c ®Þnh sè thuÕ TNDN nép theo quý')
/

-- Du lieu cua bang RCV_MAP_TKHAI
INSERT INTO rcv_map_tkhai
VALUES
('01','14','Tê khai thuÕ gi¸ trÞ gia t¨ng khÊu trõ')
/
INSERT INTO rcv_map_tkhai
VALUES
('02','26','B¶n x¸c ®Þnh sè thuÕ TNDN nép theo quý')
/

-- Du lieu cua bang RCV_GDIEN_TKHAI
INSERT INTO rcv_gdien_tkhai
VALUES
(1,'Kh«ng cã ho¹t ®éng mua, b¸n ph¸t sinh trong kú','1',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'0101','252')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(2,'ThuÕ gi¸ trÞ gi¸ t¨ng cßn ®­îc khÊu trõ kú tr­íc chuyÓn sang',NULL,'2',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,2,'0101','253')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(3,'Kª khai thuÕ GTGT ph¶i nép ng©n s¸ch Nhµ n­íc','3',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,3,'0101','254')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(4,'Hµng hãa dÞch vô (HHDV) mua vµo','4',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,4,'0101','255')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(5,'Hµng ho¸ dÞch vô mua vµo trong kú ([12]=[14]+[16]; [13]=[15]+[17])','5','6',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,5,'0101','256')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(6,'  Hµng hãa, dÞch vô mua vµo trong n­íc','7','8',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,6,'0101','257')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(7,'  Hµng hãa, dÞch vô nhËp khÈu','9','10',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,7,'0101','258')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(8,'§iÒu chØnh thuÕ GTGT cña HHDV mua vµo c¸c kú tr­íc','11',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,8,'0101','259')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(9,'  §iÒu chØnh t¨ng','12','13',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,9,'0101','260')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(10,'  §iÒu chØnh gi¶m','14','15',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,10,'0101','261')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(11,'Tæng sè thuÕ GTGT cña HHDV mua vµo ([22]=[13]+[19]-[21])',NULL,'16',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,11,'0101','262')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(12,'Tæng sè thuÕ GTGT ®­îc khÊu trõ kú nµy',NULL,'17',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,12,'0101','263')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(13,'Hµng hãa, dÞch vô b¸n ra','18',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,13,'0101','264')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(14,'Hµng hãa, dÞch vô b¸n ra trong kú ([24]=[26]+[27]; [25]=[28])','19','20',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,14,'0101','265')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(15,'Hµng hãa, dÞch vô b¸n ra kh«ng chÞu thuÕ GTGT','21',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,15,'0101','266')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(16,'Hµng hãa, dÞch vô b¸n ra chÞu thuÕ GTGT ([27]=[29]+[30]+[32]; [28]=[31]+[33])','22','23',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,16,'0101','267')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(17,'Hµng hãa, dÞch vô b¸n ra chÞu thuÕ suÊt 0%','24',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,17,'0101','268')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(18,'Hµng hãa, dÞch vô b¸n ra chÞu thuÕ suÊt 5%','25','26',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,18,'0101','269')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(19,'Hµng hãa, dÞch vô b¸n ra chÞu thuÕ suÊt 10%','27','28',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,19,'0101','270')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(20,'§iÒu chØnh thuÕ GTGT cña HHDV b¸n ra c¸c kú tr­íc','29',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,20,'0101','271')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(21,'  §iÒu chØnh t¨ng','30','31',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,21,'0101','272')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(22,'  §iÒu chØnh gi¶m','32','33',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,22,'0101','273')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(23,'Tæng doanh thu vµ thuÕ GTGT cña HHDV b¸n ra ([38]=[24]+[34]-[36]; [39]=[25]+[35]-[37])','34','35',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,23,'0101','274')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(24,'X¸c ®Þnh nghÜa vô thuÕ GTGT ph¶i nép trong kú','36',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,24,'0101','275')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(25,'ThuÕ GTGT ph¶i nép trong kú ([40]=[39]-[23]-[11])',NULL,'37',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,25,'0101','276')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(26,'ThuÕ GTGT ch­a khÊu trõ hÕt kú nµy ([41]=[39]-[23]-[11])',NULL,'38',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,26,'0101','277')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(27,'ThuÕ GTGT ®Ò nghÞ hoµn kú nµy',NULL,'39',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,27,'0101','278')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(28,'ThuÕ GTGT cßn ®­îc khÊu trõ chuyÓn kú sau ([43]=[41]-[42])',NULL,'40',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,28,'0101','279')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(29,'Hµng hãa, dÞch vô mua vµo trong n­íc','1','2','3','4','5','6','7','8',NULL,NULL,1,'0102','257')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(36,'ThuÕ GTGT cña HHDV mua vµo dïng cho ho¹t ®éng kh«ng chÞu thuÕ','2',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,2,'0103',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(37,'ThuÕ GTGT cña HHDV mua vµo dïng chung cho ho¹t ®éng QLKD cña c¬ së kinh doanh','3',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,3,'0103',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(38,'Tæng doanh thu hµng hãa, dÞch vô b¸n ra trong kú','4',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,4,'0103','1')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(39,'Doanh thu HHDV b¸n ra chÞu thuÕ trong kú','5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,5,'0103','2')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(40,'Tû lÖ doanh thu HHDV b¸n ra chÞu thuÕ trªn tæng doanh thu cña kú kª khai (3) = (2)/(1)','6',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,6,'0103','3')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(35,'ThuÕ GTGT cña HHDV mua vµo dïng cho ho¹t ®éng chÞu thuÕ','1',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'0103',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(41,'ThuÕ GTGT cña HHDV mua vµo trong kú','7',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,7,'0103','4')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(42,'ThuÕ GTGT cña HHDV mua vµo ®­îc khÊu trõ trong kú','8',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,8,'0103','5')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(43,'Tæng thuÕ GTGT cña HHDV mua vµo dïng cho ho¹t ®éng chÞu thuÕ','1',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'0104',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(44,'Tæng thuÕ GTGT cña HHDV mua vµo dïng cho ho¹t ®éng kh«ng chÞu thuÕ','2',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,2,'0104',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(45,'Tæng thuÕ GTGT cña HHDV mua vµo dïng chung cho ho¹t ®éng QLKD cña c¬ së kinh doanh','3',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,3,'0104',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(46,'Tæng doanh thu hµng hãa, dÞch vô b¸n ra trong n¨m','4',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,4,'0104','6')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(47,'Doanh thu hµng hãa, dÞch vô b¸n ra chÞu thuÕ','5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,5,'0104','7')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(48,'Tû lÖ doanh thu HHDV b¸n ra chÞu thuÕ trªn tæng doanh thu cña n¨m (3) = (2)/(1)','6',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,6,'0104','8')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(49,'Tæng thuÕ GTGT cña HHDV mua vµo trong n¨m','7',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,7,'0104','9')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(50,'ThuÕ GTGT ®Çu vµo ®­îc khÊu trõ trong n¨m (5) = (4) x (3)','8',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,8,'0104','10')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(51,'ThuÕ GTGT ®Çu vµo ®· kª khai khÊu trõ 12 th¸ng','9',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,9,'0104','11')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(52,'§iÒu chØnh t¨ng (+), gi¶m (-) thuÕ GTGT ®Çu vµo ®­îc khÊu trõ trong n¨m (7) = (5) - (6)','10',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,10,'0104','12')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(53,'ThuÕ GTGT cña HHDV mua vµo','9',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,9,'0103',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(54,'Tæng sè thuÕ GTGT cña HHDV mua vµo','11',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,11,'0104',NULL)
/

-- Du lieu cua bang RCV_MAP_CTIEU
INSERT INTO rcv_map_ctieu
VALUES
('0101','1','N',1,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','3','N',3,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','4','N',4,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','5','N',5,'[12]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','7','N',6,'[14]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','9','N',7,'[16]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','11','N',8,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','12','N',9,'[18]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','14','N',10,'[20]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','18','N',13,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','19','N',14,'[24]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','21','N',15,'[26]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','22','N',16,'[27]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','24','N',17,'[29]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','25','N',18,'[30]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','27','N',19,'[32]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','29','N',20,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','30','N',21,'[34]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','32','N',22,'[36]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','34','N',23,'[38]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','36','N',24,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','2','N',2,'[11]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','6','N',5,'[13]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','8','N',6,'[15]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','10','N',7,'[17]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','13','N',9,'[19]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','15','N',10,'[21]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','16','N',11,'[22]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','17','N',12,'[23]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','20','N',14,'[25]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','23','N',16,'[28]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','26','N',18,'[31]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','28','N',19,'[33]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','31','N',21,'[35]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','33','N',22,'[37]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','35','N',23,'[39]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','37','N',25,'[40]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','38','N',26,'[41]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','39','N',27,'[42]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0101','40','N',28,'[43]')
/
INSERT INTO rcv_map_ctieu
VALUES
('0102','1','N',29,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0102','8','C',29,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','11','N',54,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0102','2','D',29,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0102','3','N',29,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0102','4','N',29,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0102','5','N',29,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0102','6','N',29,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0102','7','C',29,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0103','2','N',36,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0103','3','N',37,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0103','4','N',38,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0103','5','N',39,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0103','6','N',40,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0103','1','N',35,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0103','7','N',41,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0103','8','N',42,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','1','N',43,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','2','N',44,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','3','N',45,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','4','N',46,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','5','N',47,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','6','N',48,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','7','N',49,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','8','N',50,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','9','N',51,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0104','10','N',52,NULL)
/
INSERT INTO rcv_map_ctieu
VALUES
('0103','9','N',53,NULL)
/

-- Tao sequence

-- Sequence RCV_XLTK_HDR_SEQ
CREATE SEQUENCE rcv_xltk_hdr_seq
  INCREMENT BY 1
  START WITH 123937
  MINVALUE 1
  MAXVALUE 999999999999999999999999999
  NOCYCLE
  NOORDER
  CACHE 20
/
CREATE PUBLIC SYNONYM rcv_xltk_hdr_seq
  FOR rcv_xltk_hdr_seq
/
GRANT SELECT ON rcv_xltk_hdr_seq TO qlt
/
GRANT SELECT ON rcv_xltk_hdr_seq TO qlt_read
/

-- Sequence RCV_XLTK_DTL_SEQ
CREATE SEQUENCE rcv_xltk_dtl_seq
  INCREMENT BY 1
  START WITH 124117
  MINVALUE 1
  MAXVALUE 999999999999999999999999999
  NOCYCLE
  NOORDER
  CACHE 20
/
CREATE PUBLIC SYNONYM rcv_xltk_dtl_seq
  FOR rcv_xltk_dtl_seq
/
GRANT SELECT ON rcv_xltk_dtl_seq TO qlt
/
GRANT SELECT ON rcv_xltk_dtl_seq TO qlt_read
/

COMMIT
/
