-- 04/TNDN
--TT151
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_04_TNDN AS
SELECT   dtl.hdr_id
		,dtl.so_tt  so_tt
		,dtl.row_id row_id,
         MAX (dtl.doanh_thu_dv) doanh_thu_dv,
         MAX (dtl.ty_le_dv) ty_le_dv,
         MAX (dtl.so_thue_dv) so_thue_dv,
         MAX (dtl.doanh_thu_hh) doanh_thu_hh,
         MAX (dtl.ty_le_hh) ty_le_hh,
         MAX (dtl.so_thue_hh) so_thue_hh,                
         MAX (dtl.doanh_thu_khac) doanh_thu_khac,
         MAX (dtl.ty_le_khac) ty_le_khac,
         MAX (dtl.so_thue_khac) so_thue_khac,
         MAX (dtl.tong_so_thue) tong_so_thue
    FROM (SELECT tkd.hdr_id, tkd.row_id row_id, gdien.ID,
                 gdien.so_tt so_tt,
                 DECODE (gdien.cot_01,tkd.ky_hieu, tkd.gia_tri,NULL) doanh_thu_dv,
                 DECODE (gdien.cot_02,tkd.ky_hieu, tkd.gia_tri,NULL) ty_le_dv,
                 DECODE (gdien.cot_03,tkd.ky_hieu, tkd.gia_tri,NULL) so_thue_dv,
                 DECODE (gdien.cot_04,tkd.ky_hieu, tkd.gia_tri,NULL) doanh_thu_hh,
                 DECODE (gdien.cot_05,tkd.ky_hieu, tkd.gia_tri,NULL) ty_le_hh,
                 DECODE (gdien.cot_06,tkd.ky_hieu, tkd.gia_tri,NULL) so_thue_hh,
                 DECODE (gdien.cot_07,tkd.ky_hieu, tkd.gia_tri,NULL) doanh_thu_khac,
                 DECODE (gdien.cot_08,tkd.ky_hieu, tkd.gia_tri,NULL) ty_le_khac,
                 DECODE (gdien.cot_09,tkd.ky_hieu, tkd.gia_tri,NULL) so_thue_khac,
                 DECODE (gdien.cot_10,tkd.ky_hieu, tkd.gia_tri,NULL) tong_so_thue
            FROM qlt_ntk.rcv_tkhai_dtl tkd, qlt_ntk.rcv_gdien_tkhai gdien,
                 qlt_ntk.rcv_map_ctieu ctieu
           WHERE (ctieu.gdn_id = gdien.ID)
             AND (ctieu.ky_hieu = tkd.ky_hieu)
             AND (tkd.loai_dlieu ='04_TNDN')                     
             AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
GROUP BY dtl.hdr_id,dtl.so_tt, dtl.row_id;

         
-- 06/TNDN         
--TT151
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_06_TNDN AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.so_tien) so_tien
     , MAX('['||dtl.ky_hieu||']') ky_hieu
     , MAX(dtl.ma_ctieu) ma_ctieu
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien,
         ctieu.ky_hieu,
         gdien.ma_ctieu
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
  AND (tkd.loai_dlieu = '06_TNDN')
  AND gdien.loai_dlieu =  '06_TNDN'
  AND ctieu.loai_dlieu =  '06_TNDN'
  AND gdien.ma_ctieu Is Not Null
) dtl
WHERE (gd.loai_dlieu = '06_TNDN')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         dtl.so_tt;  
         
commit;	


-- 01/GTGT
-- TT26
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_GTGT_KT15 AS
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
                From QLT_NTK.RCV_GDIEN_TKHAI GD, QLT_NTK.RCV_MAP_CTIEU   CT
                where CT.GDN_ID (+) = GD.ID
                    And GD.LOAI_DLIEU = '01_GTGT15'
                ) GDIEN
        WHERE   GDIEN.KY_HIEU = TKD.KY_HIEU (+)
                And TKD.LOAI_DLIEU (+)= '01_GTGT15'
                --And HDR_ID (+) = SYS_CONTEXT ('PARAMS', 'v_Hdr_ID')
        ) DTL
GROUP BY DTL.HDR_ID, DTL.CTK_ID;

-- Pl 01-7/GTGT
-- TT26
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_GTGT_KT0107_15 AS
SELECT   dtl.hdr_id
		,dtl.so_tt  so_tt
		,dtl.row_id row_id,
         MAX (dtl.ten_cong_trinh) ten_cong_trinh,
         MAX (dtl.doanh_thu) doanh_thu,
         MAX (dtl.ma_CQT) ma_CQT,
         MAX (dtl.ten_CQT) ten_CQT,
         MAX (dtl.tyLe_Pb) tyLe_Pb,
         MAX (dtl.so_thue) so_thue    
    FROM (SELECT tkd.hdr_id, tkd.row_id row_id, gdien.ID,
                 gdien.so_tt so_tt,
                 DECODE (gdien.cot_01,tkd.ky_hieu, tkd.gia_tri,NULL) ten_cong_trinh,
                 DECODE (gdien.cot_02,tkd.ky_hieu, tkd.gia_tri,NULL) doanh_thu,
                 DECODE (gdien.cot_03,tkd.ky_hieu, tkd.gia_tri,NULL) ma_CQT,
                 DECODE (gdien.cot_04,tkd.ky_hieu, tkd.gia_tri,NULL) ten_CQT,
                 DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) tyLe_Pb,
                 DECODE (gdien.cot_06,tkd.ky_hieu, tkd.gia_tri,NULL) so_thue
            FROM QLT_NTK.rcv_tkhai_dtl tkd, QLT_NTK.rcv_gdien_tkhai gdien,
                 QLT_NTK.rcv_map_ctieu ctieu
           WHERE (ctieu.gdn_id = gdien.ID)
             AND (ctieu.ky_hieu = tkd.ky_hieu)
             AND (tkd.loai_dlieu ='01_7_GTGT15')                     
             AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
GROUP BY dtl.hdr_id,dtl.so_tt, dtl.row_id;



-- KHBS 
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_KHBS14 AS
SELECT   dtl.hdr_id, dtl.so_tt so_tt, dtl.row_id row_id,
            MAX (dtl.ten_ctieu_dc) ten_ctieu_dc, MAX (dtl.ma_ctieu) ma_ctieu,
            MAX (dtl.so_da_kk) so_da_kk, MAX (dtl.so_dieu_chinh)
                                                                so_dieu_chinh,
            MAX (dtl.so_chenh_lech) so_chenh_lech
       FROM QLT_NTK.rcv_gdien_tkhai gd,
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
               FROM QLT_NTK.RCV_TKHAI_DTL tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
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
                     OR tkd.loai_dlieu = 'KHBS_01A_TNDN_DK'
                     OR tkd.loai_dlieu = 'KHBS_01B_TNDN_DK'
                     OR tkd.loai_dlieu = 'KHBS_01_TD_GTGT'
                     OR tkd.loai_dlieu = 'KHBS_03_TD_TAIN'
					 --TT 151
                     OR tkd.loai_dlieu = 'KHBS_02_TNDN14'
                     --Bo sung to QT 2014
                     OR tkd.loai_dlieu = 'KHBS_03_TNDN14'
                     OR tkd.loai_dlieu = 'KHBS_02_TAIN14'
                     OR tkd.loai_dlieu = 'KHBS_02_BVMT14'
                     OR tkd.loai_dlieu = 'KHBS_02_PHLP'
                     OR tkd.loai_dlieu = 'KHBS_03A_TD_TAIN'
                     OR tkd.loai_dlieu = 'KHBS_01_PHLP'
                     OR tkd.loai_dlieu = 'KHBS_02_NTNN14'
                     OR tkd.loai_dlieu = 'KHBS_04_NTNN14'
                     OR tkd.loai_dlieu = 'KHBS_02_TNDN_DK'
                     OR tkd.loai_dlieu = 'KHBS_02_TAIN_DK'
                     --End QT
					 OR tkd.loai_dlieu = 'KHBS_01_TAIN_DK'
					 OR tkd.loai_dlieu = 'KHBS_04_TNDN'
					 OR tkd.loai_dlieu = 'KHBS_06_TNDN'
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
             OR gd.loai_dlieu = 'KHBS_01A_TNDN_DK'
             OR gd.loai_dlieu = 'KHBS_01B_TNDN_DK'
             OR gd.loai_dlieu = 'KHBS_01_TD_GTGT'
             OR gd.loai_dlieu = 'KHBS_03_TD_TAIN'
			 --TT 151
			 OR gd.loai_dlieu = 'KHBS_02_TNDN14'
             --Bo sung to QT 2014
             OR gd.loai_dlieu = 'KHBS_03_TNDN14'
             OR gd.loai_dlieu = 'KHBS_02_TAIN14'
             OR gd.loai_dlieu = 'KHBS_02_BVMT14'
             OR gd.loai_dlieu = 'KHBS_02_PHLP'
             OR gd.loai_dlieu = 'KHBS_03A_TD_TAIN'
             OR gd.loai_dlieu = 'KHBS_01_PHLP'
             OR gd.loai_dlieu = 'KHBS_02_NTNN14'
             OR gd.loai_dlieu = 'KHBS_04_NTNN14'
             OR gd.loai_dlieu = 'KHBS_02_TNDN_DK'
             OR gd.loai_dlieu = 'KHBS_02_TAIN_DK'
             --End QT
			 OR gd.loai_dlieu = 'KHBS_01_TAIN_DK'
			 OR gd.loai_dlieu = 'KHBS_04_TNDN'
			 OR gd.loai_dlieu = 'KHBS_06_TNDN'
            )
        AND (dtl.ID = gd.ID)
   GROUP BY dtl.hdr_id, dtl.so_tt, dtl.row_id;
