-- To khai 01/GTGT
CREATE VIEW RCV_V_TKHAI_GTGT_KT13 AS
SELECT  DTL.HDR_ID
        , DTL.CTK_ID
        , Max(DTL.SO_TT)            SO_TT
        , Max(DTL.TEN_CTIEU)         TEN_CTIEU
        , Max(DTL.DOANHSO_DTNT)     DOANHSO_DTNT
        , Max(DTL.SOTHUE_DTNT)      SOTHUE_DTNT
        , Max(DTL.KY_HIEU_CTIEU_DS) KY_HIEU_CTIEU_DS
        , Max(DTL.KY_HIEU_CTIEU_ST) KY_HIEU_CTIEU_ST
FROM    (
        SELECT  TKD.HDR_ID
                , GDIEN.ID
                , GDIEN.SO_TT
                , TKD.ROW_ID
                , GDIEN.MA_CTIEU    CTK_ID
                , GDIEN.TEN_CTIEU   TEN_CTIEU
                , Replace(Decode(GDIEN.COT_01, TKD.KY_HIEU, TKD.GIA_TRI, Null),'%','')  DOANHSO_DTNT
                , Replace(Decode(GDIEN.COT_02, TKD.KY_HIEU, TKD.GIA_TRI, Null),'%','')  SOTHUE_DTNT
                , Decode(GDIEN.COT_01, TKD.KY_HIEU, '['||GDIEN.KY_HIEU_CTIEU||']', Null)   KY_HIEU_CTIEU_DS
                , Decode(GDIEN.COT_02, TKD.KY_HIEU, '['||GDIEN.KY_HIEU_CTIEU||']', Null)   KY_HIEU_CTIEU_ST
        FROM    RCV_TKHAI_DTL   TKD,
                (
                Select GD.*, CT.KY_HIEU, CT.KY_HIEU_CTIEU
                From RCV_GDIEN_TKHAI GD, RCV_MAP_CTIEU   CT
                where CT.GDN_ID (+) = GD.ID
                    And GD.LOAI_DLIEU = '01_GTGT13'
                ) GDIEN
        WHERE   GDIEN.KY_HIEU = TKD.KY_HIEU (+)
                And TKD.LOAI_DLIEU (+)= '01_GTGT13'
                --And HDR_ID (+) = SYS_CONTEXT ('PARAMS', 'v_Hdr_ID')
        ) DTL
GROUP BY DTL.HDR_ID, DTL.CTK_ID;


-- Phu luc 01-1/GTGT
CREATE  VIEW RCV_V_PLUC_TKHAI_GTGT_KT01_13 AS
SELECT
    "HDR_ID",
    "ROW_ID",
    "SO_TT",
    "NHOM_CTIEU",
    "KY_HIEU_MAU_HDON",
    "KY_HIEU_HDON",
    "SO_HOA_DON",
    "NGAY_HOA_DON",
    "TIN",
    "TEN_DTNT",
    "TEN_HANG",
    "DOANH_SO",
    "THUE_XUAT",
    "SO_THUE",
    "GHI_CHU"
FROM
    (
        SELECT
            dtl.hdr_id

            ,
            dtl.row_id                                row_id ,
            dtl.so_tt                                 so_tt ,
            DECODE (dtl.so_tt,1,1,3,2,5,3,7,4,9,5, 0) nhom_ctieu ,
            MAX(dtl.ky_hieu_mau_hdon) ky_hieu_mau_hdon,
            MAX(dtl.ky_hieu_hdon)                     ky_hieu_hdon ,
            MAX(dtl.so_hoa_don)                       so_hoa_don ,
            MAX(dtl.ngay_hoa_don)                     ngay_hoa_don ,
            MAX(dtl.tin)                              tin ,
            MAX(dtl.ten_dtnt)                         ten_dtnt ,
            MAX(dtl.ten_hang)                         ten_hang ,
            MAX(dtl.doanh_so)                         doanh_so ,
            MAX(dtl.thue_xuat)                        thue_xuat ,
            MAX(dtl.so_thue)                          so_thue ,
            MAX(dtl.ghi_chu)                          ghi_chu
        FROM
            rcv_gdien_tkhai gd,
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.id,
                    gdien.so_tt so_tt,
                    substr(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL),1,10) ky_hieu_mau_hdon,
                    SUBSTR(DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL),1,10) ky_hieu_hdon,
                    SUBSTR(DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL),1,10) so_hoa_don,
                    DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)              ngay_hoa_don,
                    DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL)              tin,
                    DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)              ten_dtnt,
                    DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL)              ten_hang,
                    DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL)              doanh_so,
                    REPLACE(REPLACE(DECODE(gdien.cot_14, tkd.ky_hieu, tkd.gia_tri, NULL),'%',''),
                    ',','.')                                             thue_xuat,
                    DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue,
                    DECODE(gdien.cot_13, tkd.ky_hieu, tkd.gia_tri, NULL) ghi_chu
                FROM
                    rcv_tkhai_dtl tkd,
                    rcv_gdien_tkhai gdien,
                    rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.id)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '01_01_GTGT13') ) dtl
        WHERE
            (
                gd.loai_dlieu = '01_01_GTGT13')
        AND (
                dtl.id = gd.id)
        GROUP BY
            dtl.hdr_id,

            dtl.so_tt,
            dtl.row_id )
WHERE
    so_hoa_don IS NOT NULL
AND ngay_hoa_don IS NOT NULL;

-- Phu luc 01-2/GTGT
CREATE VIEW RCV_V_PLUC_TKHAI_GTGT_KT02_13 AS
Select "HDR_ID","SO_TT","ROW_ID","NHOM_CTIEU","KY_HIEU_MAU_HDON","KY_HIEU_HDON","SO_HOA_DON","NGAY_HOA_DON","TIN","TEN_DTNT","TEN_HANG","DOANH_SO","THUE_XUAT","SO_THUE","GHI_CHU"
From (SELECT dtl.hdr_id
     , dtl.so_tt so_tt
     , dtl.row_id row_id
     , DECODE (dtl.so_tt,1,1,3,2,5,3,7,4,9,5, 0) nhom_ctieu
     ,MAX(dtl.ky_hieu_mau_hdon) ky_hieu_mau_hdon
     , MAX(dtl.ky_hieu_hdon) ky_hieu_hdon
     , MAX(dtl.so_hoa_don) so_hoa_don
     , MAX(dtl.ngay_hoa_don) ngay_hoa_don
     , MAX(dtl.tin) tin
     , MAX(dtl.ten_dtnt) ten_dtnt
     , MAX(dtl.ten_hang) ten_hang
     , MAX(dtl.doanh_so) doanh_so
     , MAX(dtl.thue_xuat) thue_xuat
     , MAX(dtl.so_thue) so_thue
     , MAX(dtl.ghi_chu) ghi_chu
FROM rcv_gdien_tkhai gd,
(
SELECT   tkd.hdr_id,
         tkd.row_id row_id,
         gdien.id,
         gdien.so_tt so_tt,
          substr(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL),1,10) ky_hieu_mau_hdon,
         substr(DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL),1,10) ky_hieu_hdon,
         substr(DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL),1,10) so_hoa_don,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) ngay_hoa_don,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) tin,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) ten_dtnt,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) ten_hang,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) doanh_so,
         replace(replace(DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL),'%',''),',','.') thue_xuat,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue,
         DECODE(gdien.cot_13, tkd.ky_hieu, tkd.gia_tri, NULL) ghi_chu
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_02_GTGT13')
) dtl
WHERE (gd.loai_dlieu = '01_02_GTGT13')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.so_tt,
         dtl.row_id
        ) where so_hoa_don is not null and ngay_hoa_don is not null
/* bo dk ky_hieu_hdon is not null and */;

-- Phu luc 01-3/GTGT


-- Phu luc 01-4A/GTGT
CREATE OR REPLACE VIEW RCV_V_TKHAI_GTGT_KT_PLUC4A_13 AS
SELECT dtl.hdr_id
     , dtl.so_tt
     , MAX(dtl.ctg_id) ctg_id
     , gd.ten_ctieu
     , MAX(dtl.gia_tri_ctieu) gia_tri_ctieu
     , MAX(dtl.kieu_dlieu_ctieu) kieu_dlieu_ctieu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.so_tt so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu,DECODE(tkd.ky_hieu,'7',ROUND(to_number(REPLACE(tkd.gia_tri,',','.')),2),'9',ROUND(to_number(REPLACE(tkd.gia_tri,',','.')),0),tkd.gia_tri), NULL) gia_tri_ctieu,
         DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ctieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, gdien.ma_ctieu, NULL) ctg_id,
         gdien.id,
         tkd.row_id
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_4A_GTGT13')
    AND (gdien.ma_ctieu IS NOT NULL)
) dtl
WHERE (gd.loai_dlieu = '01_4A_GTGT13')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.so_tt,
         dtl.row_id,
         gd.ten_ctieu;

-- Phu luc 01-4B/GTGT
CREATE OR REPLACE VIEW RCV_V_TKHAI_GTGT_KT_PLUC4B_13 AS
SELECT dtl.hdr_id
     , dtl.so_tt
     , MAX(dtl.ctg_id) ctg_id
     , gd.ten_ctieu
     , MAX(dtl.gia_tri_ctieu) gia_tri_ctieu
     , MAX(dtl.kieu_dlieu_ctieu) kieu_dlieu_ctieu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.so_tt,
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
    AND (tkd.loai_dlieu = '01_4B_GTGT13')
    AND (gdien.ma_ctieu IS NOT NULL)
) dtl
WHERE (gd.loai_dlieu = '01_4B_GTGT13')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.so_tt,
         dtl.row_id,
         gd.ten_ctieu;

-- Phu luc 01-5/GTGT
CREATE OR REPLACE VIEW RCV_V_PLUC_TKHAI_GTGT_KT05_13 AS
SELECT dtl.hdr_id
     , dtl.row_id  so_tt
     , MAX(dtl.gia_tri_ky_kkhai)        so_ctu
     , MAX(dtl.gia_tri_slieu_kkhai)     ngay_nop
     , MAX(dtl.noi_nop_tien)    noi_nop_tien
     , MAX(dtl.co_quan_thue)            co_quan_thue
     , MAX(dtl.so_tien)       so_tien
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
       DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) noi_nop_tien,
       DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) co_quan_thue,
       DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_05_GTGT13')
) dtl
WHERE (gd.loai_dlieu = '01_05_GTGT13')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         gd.ma_ctieu;

-- Phu luc 01-6/GTGT
CREATE OR REPLACE VIEW RCV_V_PLUC_TKHAI_GTGT_KT06_13 AS
SELECT dtl.hdr_id
     , dtl.row_id  so_tt
     , MAX(dtl.ten_cs) ten_cs
     , MAX(dtl.ma_so_thue) ma_so_thue
     , MAX(dtl.hang_hoa_chiu_thue_5) hang_hoa_chiu_thue_5
     , MAX(dtl.hang_hoa_chiu_thue_10) hang_hoa_chiu_thue_10
     , MAX(dtl.Tong) Tong
     , MAX(dtl.so_thue_pn1) so_thue_pn1
     , MAX(dtl.so_thue_pn2) so_thue_pn2
     , MAX(dtl.CQT) CQT
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         tkd.row_id,
         gdien.id,
         gdien.so_tt,
       DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) ten_cs,       
       DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) ma_so_thue,
       DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) hang_hoa_chiu_thue_5,
       DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  hang_hoa_chiu_thue_10,
       DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) Tong,
       DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue_pn1,
       DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue_pn2,
       DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) CQT

  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_06_GTGT13')
    and gdien.loai_dlieu = '01_06_GTGT13'
) dtl
WHERE (gd.loai_dlieu = '01_06_GTGT13')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         gd.ma_ctieu;


-- Phu luc 01-7

CREATE OR REPLACE VIEW RCV_V_PLUC_TKHAI_GTGT_KT0713 AS
SELECT dtl.hdr_id
     , dtl.row_id so_tt
     , dtl.so_tt nhom
     , MAX(dtl.Loai_xe) Loai_xe
     , MAX(dtl.Don_vi_tinh) Don_vi_tinh     
     , MAX(dtl.So_xe_ban_td) So_xe_ban_td
     , MAX(dtl.Gia_tren_hd) Gia_tren_hd
     , MAX(dtl.Ghi_chu) Ghi_chu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         tkd.row_id row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) Loai_xe,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) Don_vi_tinh,         
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) So_xe_ban_td,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) Gia_tren_hd,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) Ghi_chu
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_07_GTGT13')
) dtl
WHERE ( gd.loai_dlieu = '01_07_GTGT13')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
          dtl.so_tt,
         dtl.row_id;

--To khai 02/GTGT
CREATE OR REPLACE VIEW RCV_V_TKHAI_GTGT_DTU_13 AS
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
         DECODE(gdien.cot_01, tkd.ky_hieu, '['||ctieu.ky_hieu_ctieu||']', NULL) ky_hieu_ctieu_ds,
         DECODE(gdien.cot_02, tkd.ky_hieu, '['||ctieu.ky_hieu_ctieu||']', NULL) ky_hieu_ctieu_st
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '02_GTGT13')
) dtl
WHERE (gd.loai_dlieu = '02_GTGT13')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
       --  dtl.row_id,
         dtl.ctk_id;

-- Phu luc 02-1/GTGT
CREATE OR REPLACE VIEW RCV_V_PLUC_GTGT_KT0201_13 AS
(/* Formatted on 2011/08/11 13:39 (Formatter Plus v4.8.7) */
SELECT   dtl.hdr_id,dtl.so_tt-2  so_tt, dtl.row_id row_id,
         '4' nhom_ctieu,
         MAX (dtl.ky_hieu_hdon) ky_hieu_hdon, MAX (dtl.so_hoa_don) so_hoa_don,
         MAX (dtl.ngay_hoa_don) ngay_hoa_don, MAX (dtl.tin) tin,
         MAX (dtl.ten_dtnt) ten_dtnt, MAX (dtl.ten_hang) ten_hang,
         MAX (dtl.doanh_so) doanh_so, MAX (dtl.thue_xuat) thue_xuat,
         MAX (dtl.so_thue) so_thue, MAX (dtl.ghi_chu) ghi_chu
    FROM (SELECT tkd.hdr_id, tkd.row_id row_id, gdien.ID,
                 gdien.so_tt so_tt,
                 DECODE (gdien.cot_02,
                         tkd.ky_hieu, tkd.gia_tri,
                         NULL
                        ) ky_hieu_hdon,
                 DECODE (gdien.cot_03,
                         tkd.ky_hieu, tkd.gia_tri,
                         NULL
                        ) so_hoa_don,
                 DECODE (gdien.cot_04,
                         tkd.ky_hieu, tkd.gia_tri,
                         NULL
                        ) ngay_hoa_don,
                 DECODE (gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) tin,
                 DECODE (gdien.cot_05,
                         tkd.ky_hieu, tkd.gia_tri,
                         NULL
                        ) ten_dtnt,
                 DECODE (gdien.cot_07,
                         tkd.ky_hieu, tkd.gia_tri,
                         NULL
                        ) ten_hang,
                 DECODE (gdien.cot_08,
                         tkd.ky_hieu, tkd.gia_tri,
                         NULL
                        ) doanh_so,
                 REPLACE (REPLACE (DECODE (gdien.cot_09,
                                           tkd.ky_hieu, tkd.gia_tri,
                                           NULL
                                          ),
                                   '%',
                                   ''
                                  ),
                          ',',
                          '.'
                         ) thue_xuat,
                 DECODE (gdien.cot_10,
                         tkd.ky_hieu, tkd.gia_tri,
                         NULL
                        ) so_thue,
                 DECODE (gdien.cot_11,
                         tkd.ky_hieu, tkd.gia_tri,
                         NULL
                        ) ghi_chu
            FROM rcv_tkhai_dtl tkd, rcv_gdien_tkhai gdien,
                 rcv_map_ctieu ctieu
           WHERE (ctieu.gdn_id = gdien.ID)
             AND (ctieu.ky_hieu = tkd.ky_hieu)
             AND (tkd.loai_dlieu IN
                     ('02_01_GTGT13'
                     ,'02_01_1_GTGT13',
                      '02_01_2_GTGT13',
                      '02_01_3_GTGT13',
                      '02_01_4_GTGT13',
                      '02_01_5_GTGT13',
                      '02_01_6_GTGT13',
                      '02_01_7_GTGT13',
                      '02_01_8_GTGT13',
                      '02_01_9_GTGT13'
                     )
                 )
             AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
GROUP BY dtl.hdr_id,dtl.so_tt, dtl.row_id
);


-- To khai 03/GTGT

CREATE OR REPLACE VIEW RCV_V_TKHAI_03GTGT_13 AS
SELECT
    dtl.hdr_id,
    dtl.ctk_id,
    MAX(dtl.so_tt)            so_tt ,
    MAX(dtl.so_dtnt)          so_dtnt ,
    MAX(dtl.kieu_dlieu_ds)    kieu_dlieu_ds ,
    MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
FROM
    rcv_gdien_tkhai gd,
    (
        SELECT
            tkd.hdr_id                                                             hdr_id,
            gdien.id                                                               id,
            tkd.loai_dlieu                                                         loai_dlieu,
            gdien.so_tt                                                            so_tt,
            tkd.row_id                                                             row_id,
            gdien.ma_ctieu                                                         ctk_id,
            REPLACE (DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL),',','.') so_dtnt,
            DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL)            kieu_dlieu_ds,
            DECODE(gdien.cot_01, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL)           ky_hieu_ctieu_st
        FROM
            rcv_tkhai_dtl tkd,
            rcv_gdien_tkhai gdien,
            rcv_map_ctieu ctieu
        WHERE
            (
                ctieu.gdn_id = gdien.id)
        AND (
                ctieu.ky_hieu = tkd.ky_hieu)
        AND (
                tkd.loai_dlieu = '03_GTGT13'
            ) ) dtl
WHERE
    (
        gd.loai_dlieu = dtl.loai_dlieu)
AND (
        dtl.id = gd.id)
GROUP BY
    dtl.hdr_id,
    dtl.ctk_id;


--To khai 04/GTGT
CREATE OR REPLACE VIEW RCV_V_TKHAI_04GTGT_13 AS
SELECT
    dtl.hdr_id,
    dtl.ctk_id,
    MAX(dtl.so_tt)            so_tt ,
    MAX(dtl.so_dtnt)          so_dtnt ,
    MAX(dtl.kieu_dlieu_ds)    kieu_dlieu_ds ,
    MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
FROM
    rcv_gdien_tkhai gd,
    (
        SELECT
            tkd.hdr_id                                                             hdr_id,
            gdien.id                                                               id,
            tkd.loai_dlieu                                                         loai_dlieu,
            gdien.so_tt                                                            so_tt,
            tkd.row_id                                                             row_id,
            gdien.ma_ctieu                                                         ctk_id,
            REPLACE (DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL),',','.') so_dtnt,
            DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL)            kieu_dlieu_ds,
            DECODE(gdien.cot_01, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL)           ky_hieu_ctieu_st
        FROM
            rcv_tkhai_dtl tkd,
            rcv_gdien_tkhai gdien,
            rcv_map_ctieu ctieu
        WHERE
            (
                ctieu.gdn_id = gdien.id)
        AND (
                ctieu.ky_hieu = tkd.ky_hieu)
        AND (
                tkd.loai_dlieu = '04_GTGT13' ) ) dtl
WHERE
    (
        gd.loai_dlieu = dtl.loai_dlieu)
AND (
        dtl.id = gd.id)
GROUP BY
    dtl.hdr_id,
    dtl.ctk_id;


-- Phu luc 04-1/GTGT
CREATE  VIEW RCV_V_PLUC_TK04_01_GTGT_13 AS
SELECT
    "HDR_ID",
    "ROW_ID",
    "SO_TT",
    "NHOM_CTIEU",
    "KY_HIEU_MAU_HDON",
    "KY_HIEU_HDON",
    "SO_HOA_DON",
    "NGAY_HOA_DON",
    "TIN",
    "TEN_DTNT",
    "TEN_HANG",
    "DOANH_SO",
    "GHI_CHU"
FROM
    (
        SELECT
            dtl.hdr_id

            ,
            dtl.row_id                                row_id ,
            dtl.so_tt                                 so_tt ,
            DECODE (dtl.so_tt,1,1,3,2,5,3,7,4,9,5, 0) nhom_ctieu ,
            MAX(dtl.ky_hieu_mau_hdon) ky_hieu_mau_hdon,
            MAX(dtl.ky_hieu_hdon)                     ky_hieu_hdon ,
            MAX(dtl.so_hoa_don)                       so_hoa_don ,
            MAX(dtl.ngay_hoa_don)                     ngay_hoa_don ,
            MAX(dtl.tin)                              tin ,
            MAX(dtl.ten_dtnt)                         ten_dtnt ,
            MAX(dtl.ten_hang)                         ten_hang ,
            MAX(dtl.doanh_so)                         doanh_so ,
            MAX(dtl.ghi_chu)                          ghi_chu
        FROM
            rcv_gdien_tkhai gd,
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.id,
                    gdien.so_tt so_tt,
                    substr(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL),1,10) ky_hieu_mau_hdon,
                    SUBSTR(DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL),1,10) ky_hieu_hdon,
                    SUBSTR(DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL),1,10) so_hoa_don,
                    DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)              ngay_hoa_don,
                    DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL)              tin,
                    DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)              ten_dtnt,
                    DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL)              ten_hang,
                    DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL)              doanh_so,                   
                    DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) ghi_chu
                FROM
                    rcv_tkhai_dtl tkd,
                    rcv_gdien_tkhai gdien,
                    rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.id)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '04_01_GTGT13') ) dtl
        WHERE
            (
                gd.loai_dlieu = '04_01_GTGT13')
        AND (
                dtl.id = gd.id)
        GROUP BY
            dtl.hdr_id,

            dtl.so_tt,
            dtl.row_id )
WHERE
    so_hoa_don IS NOT NULL
AND ngay_hoa_don IS NOT NULL;


-- To khai 05/GTGT
-- khong thay doi

-- To khai 01A/TNDN, 01B/TNDN
CREATE OR REPLACE VIEW RCV_V_TKHAI_TNDN_01_13 AS
SELECT
    dtl.hdr_id,
    dtl.ctk_id,
    MAX(dtl.so_tt)            so_tt ,
    MAX(dtl.so_dtnt)          so_dtnt ,
    MAX(dtl.kieu_dlieu_ds)    kieu_dlieu_ds ,
    MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
FROM
    rcv_gdien_tkhai gd,
    (
        SELECT
            tkd.hdr_id                                                             hdr_id,
            gdien.id                                                               id,
            tkd.loai_dlieu                                                         loai_dlieu,
            gdien.so_tt                                                            so_tt,
            tkd.row_id                                                             row_id,
            gdien.ma_ctieu                                                         ctk_id,
            REPLACE (DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL),',','.') so_dtnt,
            DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL)            kieu_dlieu_ds,
            DECODE(gdien.cot_01, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL)           ky_hieu_ctieu_st
        FROM
            rcv_tkhai_dtl tkd,
            rcv_gdien_tkhai gdien,
            rcv_map_ctieu ctieu
        WHERE
            (
                ctieu.gdn_id = gdien.id)
        AND (
                ctieu.ky_hieu = tkd.ky_hieu)
        AND (
                tkd.loai_dlieu = '01A_TNDN13'
            OR  tkd.loai_dlieu = '01B_TNDN13') ) dtl
WHERE
    (
        gd.loai_dlieu = dtl.loai_dlieu)
AND (
        dtl.id = gd.id)
GROUP BY
    dtl.hdr_id,
    dtl.ctk_id;


-- Phu luc 01-1/TNDN kem theo to khai 01A/TNDN, 01B/TNDN
CREATE OR REPLACE VIEW RCV_V_PLUC_TKHAI_TNDN_01_13
(hdr_id, row_id, so_tt, ten_dn, mst, co_quan_quan_ly, ty_le, so_thue_phan_bo)
AS
SELECT
    dtl.hdr_id ,
    MAX(dtl.row_id)               row_id ,
    MAX(dtl.so_tt)                so_tt ,
    MAX(dtl.ten_dn)               ten_dn ,
    MAX(dtl.mst)                  mst ,
    MAX(dtl.co_quan_thue_quan_ly) co_quan_thue_quan_ly ,
    MAX(dtl.ty_le)                ty_le ,
    MAX(dtl.so_thue_phan_bo)      so_thue_phan_bo
FROM
    rcv_gdien_tkhai gd,
    (
        SELECT
            tkd.hdr_id,
            tkd.loai_dlieu,
            gdien.id,
            tkd.row_id,
            gdien.so_tt                                          so_tt,
            DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_dn,
            DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) mst,
            DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) co_quan_thue_quan_ly,
            DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) ty_le,
            DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue_phan_bo
        FROM
            rcv_tkhai_dtl tkd,
            rcv_gdien_tkhai gdien,
            rcv_map_ctieu ctieu
        WHERE
            (
                ctieu.gdn_id = gdien.id)
        AND (
                ctieu.ky_hieu = tkd.ky_hieu)
        AND (
                tkd.loai_dlieu = '01A_01_TNDN13'
            OR  tkd.loai_dlieu = '01B_01_TNDN13') ) dtl
WHERE
    (
        gd.loai_dlieu = dtl.loai_dlieu)
AND (
        dtl.id = gd.id)
GROUP BY
    dtl.hdr_id,
    dtl.row_id,
    dtl.so_tt;


-- To khai 02/TNDN
-- view to khai khong thay doi
-- Phu luc 02-1/TNDN
CREATE OR REPLACE VIEW RCV_V_PLUC_TKHAI_TNDN_02_13
(hdr_id, row_id, so_tt, ten_ben_cn, ma_so_thue, dia_chi, HD_chuyen_nhuong)
AS
SELECT
    dtl.hdr_id ,
    MAX(dtl.row_id)               row_id ,
    MAX(dtl.so_tt)                so_tt ,   
    MAX(dtl.ten_ben_cn)                  ten_ben_cn ,
    MAX(dtl.ma_so_thue) ma_so_thue ,
    MAX(dtl.dia_chi)                dia_chi ,
    MAX(dtl.HD_chuyen_nhuong)      HD_chuyen_nhuong
FROM
    rcv_gdien_tkhai gd,
    (
        SELECT
            tkd.hdr_id,
            tkd.loai_dlieu,
            gdien.id,
            tkd.row_id,
            gdien.so_tt                                          so_tt,
            DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) ten_ben_cn,
            DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL)ma_so_thue,
            DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) dia_chi,
            DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) HD_chuyen_nhuong
        FROM
            rcv_tkhai_dtl tkd,
            rcv_gdien_tkhai gdien,
            rcv_map_ctieu ctieu
        WHERE
            (
                ctieu.gdn_id = gdien.id)
        AND (
                ctieu.ky_hieu = tkd.ky_hieu)
        AND (
                tkd.loai_dlieu = '02_01_TNDN13'
            OR  tkd.loai_dlieu = '02_01_TNDN13') ) dtl
WHERE
    (
        gd.loai_dlieu = dtl.loai_dlieu)
AND (
        dtl.id = gd.id)
GROUP BY
    dtl.hdr_id,
    dtl.row_id,
    dtl.so_tt;
    
   
 --To khai 01/TBVMT  
 CREATE OR REPLACE VIEW RCV_V_TKHAI_TBVMT_01 AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.don_vi_tinh) don_vi_tinh
     , MAX(dtl.so_luong) so_luong
     , MAX(dtl.muc_thue) muc_thue
     , MAX(dtl.Thue_BVMT) Thue_BVMT
     , MAX(dtl.ten_hang_hoa) ten_hang_hoa    
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) don_vi_tinh,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) so_luong,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) muc_thue,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) Thue_BVMT,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) ten_hang_hoa
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_TBVMT13')
    and gdien.loai_dlieu =  '01_TBVMT13'
   and ctieu.loai_dlieu =  '01_TBVMT13'
) dtl
WHERE (gd.loai_dlieu = '01_TBVMT13')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
          dtl.so_tt;


-- 






-- Phu luc 01/KHBS cho to khai 01/GTGT
CREATE OR REPLACE VIEW RCV_V_PLUC_KHBS_GTGT_KT13 AS
SELECT dtl.hdr_id
     , dtl.so_tt so_tt
     , dtl.row_id row_id
     , MAX(dtl.ten_ctieu_dc) ten_ctieu_dc
     , MAX(dtl.ma_ctieu) ma_ctieu
     , MAX(dtl.so_da_kk) so_da_kk
     , MAX(dtl.so_dieu_chinh) so_dieu_chinh
     , MAX(dtl.so_chenh_lech) so_chenh_lech
FROM rcv_gdien_tkhai gd,
(
SELECT   tkd.hdr_id,
         tkd.row_id row_id,
         gdien.id,
         gdien.so_tt so_tt,
    	   DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_ctieu_dc,
    	   DECODE(gdien.cot_02, tkd.ky_hieu, '['||tkd.gia_tri||']', NULL) ma_ctieu,
    	   DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) so_da_kk,
    	   DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_dieu_chinh,
    	   DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) so_chenh_lech
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = 'KHBS_01_GTGT13')
) dtl
WHERE ( gd.loai_dlieu = 'KHBS_01_GTGT13')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.so_tt,
         dtl.row_id;


-- Phu luc 01/KHBS cho cac to khai con lai theo TT156
CREATE OR REPLACE VIEW RCV_V_PLUC_KHBS13 AS
SELECT   dtl.hdr_id, dtl.so_tt so_tt, dtl.row_id row_id,
            MAX (dtl.ten_ctieu_dc) ten_ctieu_dc, MAX (dtl.ma_ctieu) ma_ctieu,
            MAX (dtl.so_da_kk) so_da_kk, MAX (dtl.so_dieu_chinh)
                                                                so_dieu_chinh,
            MAX (dtl.so_chenh_lech) so_chenh_lech
       FROM rcv_gdien_tkhai gd,
            (SELECT tkd.hdr_id, tkd.row_id row_id, gdien.ID,
                    gdien.so_tt so_tt,
                    DECODE (gdien.cot_01,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) ten_ctieu_dc,
                    DECODE (gdien.cot_02,
                            tkd.ky_hieu, '[' || tkd.gia_tri || ']',
                            NULL
                           ) ma_ctieu,
                    DECODE (gdien.cot_03,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) so_da_kk,
                    DECODE (gdien.cot_04,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) so_dieu_chinh,
                    DECODE (gdien.cot_05,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) so_chenh_lech
               FROM rcv_tkhai_dtl tkd,
                    rcv_gdien_tkhai gdien,
                    rcv_map_ctieu ctieu
              WHERE (ctieu.gdn_id = gdien.ID)
                AND (ctieu.ky_hieu = tkd.ky_hieu)
                AND (   tkd.loai_dlieu = 'KHBS_01A_TNDN13'
                     OR tkd.loai_dlieu = 'KHBS_01B_TNDN13'
                     OR tkd.loai_dlieu = 'KHBS_03_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_02_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_01_TAIN13'
                     OR tkd.loai_dlieu = 'KHBS_01_TTDB13'
                     OR tkd.loai_dlieu = 'KHBS_03_NTNN13'
                     OR tkd.loai_dlieu = 'KHBS_05_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_01_BVMT13'                    
                     OR tkd.loai_dlieu = 'KHBS_04_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_02_TNDN13'
                     OR tkd.loai_dlieu = 'KHBS_01_PHXD13'
                      OR tkd.loai_dlieu = 'KHBS_01_NTNN13'
                       OR tkd.loai_dlieu = 'KHBS_01_TBVMT13'
                    )) dtl
      WHERE (   gd.loai_dlieu = 'KHBS_01A_TNDN13'
             OR gd.loai_dlieu = 'KHBS_01B_TNDN13'
             OR gd.loai_dlieu = 'KHBS_03_GTGT13'
             OR gd.loai_dlieu = 'KHBS_02_GTGT13'
             OR gd.loai_dlieu = 'KHBS_01_TAIN13'
             OR gd.loai_dlieu = 'KHBS_01_TTDB13'
             OR gd.loai_dlieu = 'KHBS_03_NTNN13'            
             OR gd.loai_dlieu = 'KHBS_05_GTGT13'
             OR gd.loai_dlieu = 'KHBS_01_BVMT13'
             OR gd.loai_dlieu = 'KHBS_04_GTGT13'
             OR gd.loai_dlieu = 'KHBS_02_TNDN13'
             OR gd.loai_dlieu = 'KHBS_01_PHXD13'
			 OR gd.loai_dlieu = 'KHBS_01_NTNN13'
       OR gd.loai_dlieu = 'KHBS_01_TBVMT13'
            )
        AND (dtl.ID = gd.ID)
   GROUP BY dtl.hdr_id, dtl.so_tt, dtl.row_id;





