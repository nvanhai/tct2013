/****************************************************************************************/
/*=============================  */
/*01_GTGT: PL 01-1*/
CREATE OR REPLACE VIEW RCV_V_PLUC_TKHAI_GTGT_KT01_13 AS
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

/*=============================  
'01_GTGT: PL 01-2*/
CREATE OR REPLACE VIEW RCV_V_PLUC_TKHAI_GTGT_KT02_13 AS
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

/*=============================  
'01_GTGT: PL 01-3
'=============================  
'01_GTGT: PL 01-5*/
CREATE OR REPLACE VIEW RCV_V_PLUC_TKHAI_GTGT_KT05_13
(hdr_id, so_tt, so_ctu, ngay_nop, noi_nop_tien,co_quan_thue,so_tien)
AS
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

/****************************************************************************************
'=============================  
'03_GTGT13*/
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

/****************************************************************************************
'=============================    
04_GTGT13*/
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

/****************************************************************************************    
'=============================
'05_GTGT11*/
CREATE OR REPLACE VIEW
    RCV_V_TKHAI_GTGT_TT11
    (
        HDR_ID,
        ROW_ID,
        SO_TT,
        MA_CTIEU,
        HANG_HOA_DICH_VU_CHIU_THUE
    ) AS
SELECT
    dtl.hdr_id ,
    dtl.row_id ,
    dtl.so_tt ,
    dtl.ma_ctieu ,
    MAX(dtl.hang_hoa_dich_vu_chiu_thue) hang_hoa_dich_vu_chiu_thue
FROM
    rcv_gdien_tkhai gd,
    (
        SELECT
            tkd.hdr_id,
            NVL(tkd.row_id,0) row_id,
            gdien.id,
            gdien.so_tt,
            gdien.ma_ctieu,
            DECODE(gdien.cot_01,tkd.ky_hieu, tkd.gia_tri, NULL) hang_hoa_dich_vu_chiu_thue
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
                tkd.loai_dlieu = '05_GTGT11')
        AND gdien.loai_dlieu = '05_GTGT11'
        AND ctieu.loai_dlieu = '05_GTGT11') dtl
WHERE
    (
        gd.loai_dlieu = '05_GTGT11')
AND (
        dtl.id = gd.id)
GROUP BY
    dtl.hdr_id,
    dtl.row_id,
    dtl.so_tt,
    dtl.ma_ctieu ;

/****************************************************************************************    
'=============================    

'01A/TNDN13,01B/TNDN13 (dung chung view)*/
CREATE OR REPLACE VIEW
    RCV_V_TKHAI_TNDN_01A_13
    (
        HDR_ID,
        CTK_ID,
        SO_TT,
        SO_DTNT,
        KIEU_DLIEU_DS,
        KY_HIEU_CTIEU_ST
    ) AS
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
    
/*============================= 
'PL 01-1 cho O1A/TNDN,01B/TNDN*/
CREATE OR REPLACE VIEW
    RCV_V_PLUC_TKHAI_TNDN_01_13
    (
        HDR_ID,
        ROW_ID,
        SO_TT,
        TEN_DN,
        MST,
        CO_QUAN_QUAN_LY,
        TY_LE,
        SO_THUE_PHAN_BO
    ) AS
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
/*============================= 
'****************************************************************************************/