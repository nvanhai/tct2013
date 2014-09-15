--03/TNDN
CREATE OR REPLACE VIEW RCV_V_TKHAI_03_TNDN AS
SELECT dtl.hdr_id, dtl.ctk_id, MAX(dtl.so_tt) so_tt
, MAX(dtl.so_dtnt) so_dtnt
, MAX(dtl.kieu_dlieu_ds) kieu_dlieu_ds
, MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
FROM QLT_NTK.rcv_gdien_tkhai gd,
(SELECT tkd.hdr_id hdr_id,
gdien.id id,
tkd.loai_dlieu loai_dlieu,
gdien.so_tt so_tt,
tkd.row_id row_id,
gdien.ma_ctieu ctk_id,
replace (DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL),',','.') so_dtnt,
DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ds,
DECODE(gdien.cot_01, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL) ky_hieu_ctieu_st
FROM QLT_NTK.rcv_tkhai_dtl tkd,
QLT_NTK.rcv_gdien_tkhai gdien,
QLT_NTK.rcv_map_ctieu ctieu
WHERE (ctieu.gdn_id = gdien.id)
AND (ctieu.ky_hieu = tkd.ky_hieu)
AND (tkd.loai_dlieu = '03_TNDN11')
) dtl
WHERE (gd.loai_dlieu = dtl.loai_dlieu)
--(gd.loai_dlieu = '03_TNDN11' OR gd.loai_dlieu = '01B_TNDN')
AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
dtl.ctk_id
;

--02/TAIN
CREATE OR REPLACE VIEW RCV_V_TKHAI_02_TAIN AS
SELECT hdr_id
      ,row_id
      ,so_tt
      ,ddiem_kthac
      ,btn_id
      ,don_vi_tinh
      ,san_luong
      ,gia_don_vi
      ,tsuat_dtnt
      ,gia_tt_don_vi
      ,decode(gia_don_vi,0,san_luong*gia_tt_don_vi,san_luong*gia_don_vi*(tsuat_dtnt/100)) thue_phai_nop_dtnt
      ,thue_da_ke_khai
      ,chenh_lech
FROM
(
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.ddiem_kthac) ddiem_kthac
     , MAX(dtl.btn_id) btn_id
     , MAX(dtl.don_vi_tinh) don_vi_tinh
     , MAX(dtl.san_luong) san_luong
     , MAX(dtl.gia_don_vi) gia_don_vi
     , MAX(dtl.tsuat_dtnt) tsuat_dtnt
     , MAX(dtl.gia_tt_don_vi) gia_tt_don_vi
     , MAX(dtl.thue_phai_nop_dtnt) thue_phai_nop_dtnt
     , MAX(dtl.thue_da_ke_khai) thue_da_ke_khai
     , MAX(dtl.chenh_lech) chenh_lech
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ddiem_kthac,
    	   DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) btn_id,
    	   DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) don_vi_tinh,
    	   DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) san_luong,
    	   replace(DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL),',','.') gia_don_vi,
    	   DECODE(gdien.cot_06, tkd.ky_hieu, NVL(tkd.gia_tri,0), NULL) tsuat_dtnt,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tt_don_vi,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) thue_phai_nop_dtnt,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) thue_da_ke_khai,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) chenh_lech
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '02_TAIN11')
) dtl
WHERE (gd.loai_dlieu = '02_TAIN11')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id
         , dtl.so_tt
);

--01/PHLP
CREATE OR REPLACE VIEW RCV_V_TKHAI_01_PHLP AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.loai_phi) loai_phi
     , MAX(dtl.tieu_muc) tieu_muc
     , MAX(dtl.so_tien_thu_duoc) so_tien_thu_duoc
     , MAX(dtl.ty_le) ty_le
     , MAX(dtl.so_tien_trich_cho_che_do) so_tien_trich_cho_che_do
     , MAX(dtl.so_tien_phai_nop) so_tien_phai_nop
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) loai_phi,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) tieu_muc,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_thu_duoc,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) ty_le,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_trich_cho_che_do,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_phai_nop
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_PHLP')
    and gdien.loai_dlieu =  '01_PHLP'
   and ctieu.loai_dlieu =  '01_PHLP'
) dtl
WHERE (gd.loai_dlieu = '01_PHLP')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
          dtl.so_tt;
--03/TD-TAIN
CREATE OR REPLACE VIEW RCV_V_TKHAI_03_TD_TAIN AS
(
        
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.NHA_MAY_TD) NHA_MAY_TD,
            MAX (dtl.MA_SO_THUE) MA_SO_THUE,
            MAX (dtl.SAN_LUONG)     SAN_LUONG,
            MAX (dtl.GIA_TINH_THUE)       GIA_TINH_THUE,
            MAX (dtl.THUE_PHAT_SINH)     THUE_PHAT_SINH,
            MAX (dtl.THUE_MIEN_GIAM)              THUE_MIEN_GIAM,
            MAX (dtl.THUE_PHAI_NOP)              THUE_PHAI_NOP
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) NHA_MAY_TD,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) SAN_LUONG,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) GIA_TINH_THUE,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL ) THUE_PHAT_SINH,
                    DECODE (gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_MIEN_GIAM,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_PHAI_NOP
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '03_TD_TAIN' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    );
	--PL 03-1
CREATE OR REPLACE VIEW RCV_V_PLUC_03_1_TD_TAIN AS
(
        
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.STT) STT,
            MAX (dtl.CHI_TIEU)     CHI_TIEU,
            MAX (dtl.MA_SO_THUE)       MA_SO_THUE,
            MAX (dtl.CQT_QUAN_LY)     CQT_QUAN_LY,
			MAX (dtl.CQT_PARENT_ID)     CQT_PARENT_ID,
            MAX (dtl.TY_LE_PHAN_BO)              TY_LE_PHAN_BO,
            MAX (dtl.THUE_PHAI_NOP)              THUE_PHAI_NOP
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) STT,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) CHI_TIEU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL ) CQT_QUAN_LY,
					DECODE (gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL ) CQT_PARENT_ID,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)  TY_LE_PHAN_BO,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_PHAI_NOP
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '03_1_TD_TAIN' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    );
	
--02/BVMT
CREATE OR REPLACE VIEW RCV_V_TKHAI_02_BVMT AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.don_vi_tinh) don_vi_tinh
     , MAX(dtl.so_luong) so_luong
     , MAX(dtl.muc_phi) muc_phi
     , (MAX(dtl.so_luong)*MAX(dtl.muc_phi)) phi_phai_nop
     , MAX(dtl.phi_ke_khai) phi_ke_khai
     , MAX(dtl.chenh_lech) chenh_lech
     , MAX(dtl.loai_khoang_san) loai_khoang_san
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) don_vi_tinh,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_luong,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) muc_phi,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) phi_ke_khai,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) chenh_lech,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) loai_khoang_san
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '02_BVMT11')
    and gdien.loai_dlieu =  '02_BVMT11'
   and ctieu.loai_dlieu =  '02_BVMT11'
) dtl
WHERE (gd.loai_dlieu = '02_BVMT11')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
          dtl.so_tt;
--02/NTNN
CREATE OR REPLACE VIEW RCV_V_TKHAI_02_NTNN AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , dtl.ctq_id
     , MAX(dtl.don_vi_tinh) don_vi_tinh
     , MAX(dtl.ke_khai) ke_khai
     , MAX(dtl.quyet_toan) quyet_toan
     , MAX(dtl.ghi_chu) ghi_chu
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.ma_ctieu ctq_id,
         gdien.so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) don_vi_tinh,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) ke_khai,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) quyet_toan,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) ghi_chu
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '02_NTNN')
    and gdien.loai_dlieu =  '02_NTNN'
   and ctieu.loai_dlieu =  '02_NTNN'
) dtl
WHERE (gd.loai_dlieu = '02_NTNN')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
          dtl.so_tt,
          dtl.ctq_id;
	--02-1-NTNN
CREATE OR REPLACE VIEW RCV_V_PLUC_02_1_NTNN AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.ten_nha_thau_nuoc_ngoai) ten_nha_thau_nuoc_ngoai
     , MAX(dtl.nuoc_cu_tru) nuoc_cu_tru
     , MAX(dtl.ma_so_thue_VN) ma_so_thue_VN
     , MAX(dtl.ma_so_thue_nuoc_ngoai) ma_so_thue_nuoc_ngoai
     , MAX(dtl.so_hop_dong) so_hop_dong
     , MAX(dtl.noi_dung) noi_dung
     , MAX(dtl.dia_diem) dia_diem
     , MAX(dtl.thoi_han) thoi_han
     , MAX(dtl.gia_tri_nguyen_te_hop_dong) gia_tri_nguyen_te_hop_dong
     , MAX(dtl.gia_tri_quy_doi_hop_dong) gia_tri_quy_doi_hop_dong
     , MAX(dtl.gia_tri_nguyen_te_quyet_toan) gia_tri_nguyen_te_quyet_toan
     , MAX(dtl.gia_tri_quy_doi_quyet_toan) gia_tri_quy_doi_quyet_toan
     , MAX(dtl.so_luong_LD) so_luong_LD
     , dtl.row_id
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_nha_thau_nuoc_ngoai,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) nuoc_cu_tru,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) ma_so_thue_VN,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) ma_so_thue_nuoc_ngoai,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) so_hop_dong,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) noi_dung,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) dia_diem,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) thoi_han,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_nguyen_te_hop_dong,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_quy_doi_hop_dong,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_nguyen_te_quyet_toan,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_quy_doi_quyet_toan,
         DECODE(gdien.cot_13, tkd.ky_hieu, tkd.gia_tri, NULL) so_luong_LD
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu = '02_1_NTNN')
) dtl
WHERE (gd.loai_dlieu  = '02_1_NTNN')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.row_id,
         dtl.so_tt
;
		--02-2-NTNN
CREATE OR REPLACE VIEW RCV_V_PLUC_02_2_NTNN AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.ten_nha_thau_phu_VN) ten_nha_thau_phu_VN
     , MAX(dtl.ma_so_thue) ma_so_thue
     , MAX(dtl.nha_thau_NN_ky_hd) nha_thau_NN_ky_hd
     , MAX(dtl.hop_dong_so) hop_dong_so
     , MAX(dtl.noi_dung) noi_dung
     , MAX(dtl.dia_diem) dia_diem
     , MAX(dtl.thoi_han) thoi_han
     , MAX(dtl.gia_tri_nguyen_te_hop_dong) gia_tri_nguyen_te_hop_dong
     , MAX(dtl.gia_tri_quy_doi_hop_dong) gia_tri_quy_doi_hop_dong
     , MAX(dtl.gia_tri_nguyen_te_quyet_toan) gia_tri_nguyen_te_quyet_toan
     , MAX(dtl.gia_tri_quy_doi_quyet_toan) gia_tri_quy_doi_quyet_toan
     , dtl.row_id
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_nha_thau_phu_VN,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) ma_so_thue,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) nha_thau_NN_ky_hd,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) hop_dong_so,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) noi_dung,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) dia_diem,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) thoi_han,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_nguyen_te_hop_dong,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_quy_doi_hop_dong,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_nguyen_te_quyet_toan,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_quy_doi_quyet_toan
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu = '02_2_NTNN')
) dtl
WHERE (gd.loai_dlieu  = '02_2_NTNN')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.row_id,
         dtl.so_tt
;
--02/TAIN-DK
CREATE OR REPLACE VIEW RCV_V_TKHAI_02_TAIN_DK AS
SELECT
    DTL.HDR_ID ,
    DTL.CTK_ID ,
    MAX(DTL.SO_TT)         SO_TT ,
    MAX(DTL.TEN_CTIEU)     TEN_CTIEU ,
    MAX(DTL.GIA_TRI)       GIA_TRI ,
    MAX(DTL.KY_HIEU_CTIEU) KY_HIEU_CTIEU
FROM
    (
        SELECT
            TKD.HDR_ID ,
            GDIEN.ID ,
            GDIEN.SO_TT ,
            TKD.ROW_ID ,
            GDIEN.MA_CTIEU                                                       CTK_ID ,
            GDIEN.TEN_CTIEU                                                      TEN_CTIEU ,
            REPLACE(DECODE(GDIEN.COT_01, TKD.KY_HIEU, TKD.GIA_TRI, NULL),'%','')   GIA_TRI ,
            DECODE(GDIEN.COT_01, TKD.KY_HIEU, GDIEN.KY_HIEU_CTIEU, NULL) KY_HIEU_CTIEU
        FROM
            QLT_NTK.RCV_TKHAI_DTL TKD,
            (
                SELECT
                    GD.*,
                    CT.KY_HIEU,
                    CT.KY_HIEU_CTIEU
                FROM
                    QLT_NTK.RCV_GDIEN_TKHAI GD,
                    QLT_NTK.RCV_MAP_CTIEU CT
                WHERE
                    CT.GDN_ID (+) = GD.ID
                AND GD.LOAI_DLIEU = '02_TAIN_DK' ) GDIEN
        WHERE
            GDIEN.KY_HIEU = TKD.KY_HIEU (+)
        AND TKD.LOAI_DLIEU (+)= '02_TAIN_DK' ) DTL
GROUP BY
    DTL.HDR_ID,
    DTL.CTK_ID;

	--PL 02-1/TAIN-DK
CREATE OR REPLACE VIEW RCV_V_PLUC_02_1_TAIN_DK AS
(

        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.MA_SO_THUE) MA_SO_THUE,
            MAX (dtl.TEN_NHA_THAU)     TEN_NHA_THAU,
            MAX (dtl.TY_LE_PHAN_BO)       TY_LE_PHAN_BO,
            MAX (dtl.SO_THUE_PHAT_SINH_PHAI_NOP)     SO_THUE_PHAT_SINH_PHAI_NOP,
            MAX (dtl.GHI_CHU)              GHI_CHU
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) TEN_NHA_THAU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) TY_LE_PHAN_BO,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) SO_THUE_PHAT_SINH_PHAI_NOP,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  GHI_CHU
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '02_1_TAIN_DK' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    );
	--PL 02-2/TAIN-DK		
CREATE OR REPLACE VIEW RCV_V_PLUC_02_2_TAIN_DK AS
(

        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.don_vi_tinh) don_vi_tinh,
            MAX (dtl.ngay_khai_thac)     ngay_khai_thac,
            MAX (dtl.san_luong_khai_thac)       san_luong_khai_thac,
            MAX (dtl.ngay_xuat_ban)     ngay_xuat_ban,
            MAX (dtl.san_luong_xuat_ban)              san_luong_xuat_ban,
            MAX (dtl.gia_tinh_thue)              gia_tinh_thue,
            MAX (dtl.doanh_thu)              doanh_thu,
            MAX (dtl.ghi_chu)              ghi_chu
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) don_vi_tinh,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) ngay_khai_thac,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) san_luong_khai_thac,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL ) ngay_xuat_ban,
                    DECODE (gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)  san_luong_xuat_ban,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)  gia_tinh_thue,
                    DECODE (gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL)  doanh_thu,
                    DECODE (gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL)  ghi_chu
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '02_2_TAIN_DK' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    );
--02/TNDN-DK
CREATE OR REPLACE VIEW RCV_V_TKHAI_02_TNDN_DK AS
SELECT
    DTL.HDR_ID ,
    DTL.CTK_ID ,
    MAX(DTL.SO_TT)         SO_TT ,
    MAX(DTL.TEN_CTIEU)     TEN_CTIEU ,
    MAX(DTL.GIA_TRI)       GIA_TRI ,
    MAX(DTL.KY_HIEU_CTIEU) KY_HIEU_CTIEU
FROM
    (
        SELECT
            TKD.HDR_ID ,
            GDIEN.ID ,
            GDIEN.SO_TT ,
            TKD.ROW_ID ,
            GDIEN.MA_CTIEU                                                       CTK_ID ,
            GDIEN.TEN_CTIEU                                                      TEN_CTIEU ,
            REPLACE(DECODE(GDIEN.COT_01, TKD.KY_HIEU, TKD.GIA_TRI, NULL),'%','')   GIA_TRI ,
            DECODE(GDIEN.COT_01, TKD.KY_HIEU, GDIEN.KY_HIEU_CTIEU, NULL) KY_HIEU_CTIEU
        FROM
            QLT_NTK.RCV_TKHAI_DTL TKD,
            (
                SELECT
                    GD.*,
                    CT.KY_HIEU,
                    CT.KY_HIEU_CTIEU
                FROM
                    QLT_NTK.RCV_GDIEN_TKHAI GD,
                    QLT_NTK.RCV_MAP_CTIEU CT
                WHERE
                    CT.GDN_ID (+) = GD.ID
                AND GD.LOAI_DLIEU = '02_TNDN_DK' ) GDIEN
        WHERE
            GDIEN.KY_HIEU = TKD.KY_HIEU (+)
        AND TKD.LOAI_DLIEU (+)= '02_TNDN_DK' ) DTL
GROUP BY
    DTL.HDR_ID,
    DTL.CTK_ID;

	--PL 02-1/TNDN-DK
CREATE OR REPLACE VIEW RCV_V_PLUC_02_1_TNDN_DK AS
(

        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.MA_SO_THUE) MA_SO_THUE,
            MAX (dtl.TEN_NHA_THAU)     TEN_NHA_THAU,
            MAX (dtl.TY_LE_PHAN_BO)       TY_LE_PHAN_BO,
            MAX (dtl.SO_THUE_PHAT_SINH_PHAI_NOP)     SO_THUE_PHAT_SINH_PHAI_NOP,
            MAX (dtl.GHI_CHU)              GHI_CHU
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) TEN_NHA_THAU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) TY_LE_PHAN_BO,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) SO_THUE_PHAT_SINH_PHAI_NOP,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  GHI_CHU
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '02_1_TNDN_DK' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    );
	