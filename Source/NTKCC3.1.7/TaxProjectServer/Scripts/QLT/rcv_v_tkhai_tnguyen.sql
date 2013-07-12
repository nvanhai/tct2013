-- Start of DDL Script for View QLT_OWNER.RCV_V_TKHAI_TNGUYEN
-- Generated 27-Apr-2006 14:39:01 from QLT_OWNER@QLT_93

CREATE OR REPLACE VIEW rcv_v_tkhai_tnguyen (
   hdr_id,
   ddiem_kthac,
   btn_id,
   ten_tnguyen,
   don_vi_tinh,
   san_luong,
   gia_don_vi,
   tsuat_dtnt,
   gia_tt_don_vi,
   thue_psinh_tky_dtnt,
   thue_mien_giam_dtnt,
   thue_phai_nop_dtnt,
   kieu_ddiem_kthac,
   kieu_btn_id,
   kieu_ten_tnguyen,
   kieu_don_vi_tinh,
   kieu_san_luong,
   kieu_gia_don_vi,
   kieu_tsuat_dtnt,
   kieu_gia_tt_don_vi,
   kieu_thue_psinh_tky_dtnt,
   kieu_thue_mien_giam_dtnt,
   kieu_thue_phai_nop_dtnt )
AS
SELECT dtl.hdr_id
     , MAX(dtl.ddiem_kthac) ddiem_kthac
     , MAX(dtl.btn_id) btn_id
     , MAX(dtl.ten_tnguyen) ten_tnguyen
     , MAX(dtl.don_vi_tinh) don_vi_tinh
     , MAX(dtl.san_luong) san_luong
     , MAX(dtl.gia_don_vi) gia_don_vi
     , MAX(dtl.tsuat_dtnt) tsuat_dtnt
     , MAX(dtl.gia_tt_don_vi) gia_tt_don_vi
     , MAX(dtl.thue_psinh_tky_dtnt) thue_psinh_tky_dtnt
     , MAX(dtl.thue_mien_giam_dtnt) thue_mien_giam_dtnt
     , MAX(dtl.thue_phai_nop_dtnt) thue_phai_nop_dtnt
     , MAX(dtl.kieu_ddiem_kthac) kieu_ddiem_kthac
     , MAX(dtl.kieu_btn_id) kieu_btn_id
     , MAX(dtl.kieu_ten_tnguyen) kieu_ten_tnguyen
     , MAX(dtl.kieu_don_vi_tinh) kieu_don_vi_tinh
     , MAX(dtl.kieu_san_luong) kieu_san_luong
     , MAX(dtl.kieu_gia_don_vi) kieu_gia_don_vi
     , MAX(dtl.kieu_tsuat_dtnt) kieu_tsuat_dtnt
     , MAX(dtl.kieu_gia_tt_don_vi) kieu_gia_tt_don_vi
     , MAX(dtl.kieu_thue_psinh_tky_dtnt) kieu_thue_psinh_tky_dtnt
     , MAX(dtl.kieu_thue_mien_giam_dtnt) kieu_thue_mien_giam_dtnt
     , MAX(dtl.kieu_thue_phai_nop_dtnt) kieu_thue_phai_nop_dtnt
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ddiem_kthac,
    	 DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) btn_id,
    	 DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) ten_tnguyen,
    	 DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) don_vi_tinh,
    	 DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) san_luong,
    	 DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) gia_don_vi,
    	 DECODE(gdien.cot_07, tkd.ky_hieu, NVL(tkd.gia_tri,0), NULL) tsuat_dtnt,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tt_don_vi,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) thue_psinh_tky_dtnt,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) thue_mien_giam_dtnt,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) thue_phai_nop_dtnt, 	
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_ddiem_kthac,
    	 DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_btn_id,
    	 DECODE(gdien.cot_03, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_ten_tnguyen,
    	 DECODE(gdien.cot_04, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_don_vi_tinh,
    	 DECODE(gdien.cot_05, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_san_luong,
    	 DECODE(gdien.cot_06, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_gia_don_vi,
    	 DECODE(gdien.cot_07, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_tsuat_dtnt,
    	 DECODE(gdien.cot_08, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_gia_tt_don_vi,    	
    	 DECODE(gdien.cot_09, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_thue_psinh_tky_dtnt,
    	 DECODE(gdien.cot_10, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_thue_mien_giam_dtnt,
    	 DECODE(gdien.cot_11, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_thue_phai_nop_dtnt                      	
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0601')
) dtl
WHERE (gd.loai_dlieu = '0601')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id
/

-- End of DDL Script for View QLT_OWNER.RCV_V_TKHAI_TNGUYEN

