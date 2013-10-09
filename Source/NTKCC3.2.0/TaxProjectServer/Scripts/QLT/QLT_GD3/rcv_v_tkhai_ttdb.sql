CREATE OR REPLACE VIEW rcv_v_tkhai_ttdb (
   hdr_id,
   loai_ctieu,
   id_tt,
   btt_id,
   ten_ttdb,
   dvt_don_vi_tinh,
   so_luong,
   tong_tri_gia_ban,
   tong_tri_gia_tt_dtnt,
   tsuat_dtnt,
   thue_duoc_ktru,
   dchinh_tang_giam,
   thue_pnop_tky_dtnt,
   kieu_btt_id,
   kieu_ten_ttdb,
   kieu_dvt_don_vi_tinh,
   kieu_so_luong,
   kieu_tong_tri_gia_ban,
   kieu_tong_tri_gia_tt_dtnt,
   kieu_tsuat_dtnt,
   kieu_thue_duoc_ktru,
   kieu_dchinh_tang_giam,
   kieu_thue_pnop_tky_dtnt )
AS
SELECT dtl.hdr_id
     , dtl.loai_ctieu
     , MAX(dtl.id_tt) id_tt
     , MAX(dtl.btt_id) btt_id
     , MAX(dtl.ten_ttdb) ten_ttdb
     , MAX(dtl.dvt_don_vi_tinh) dvt_don_vi_tinh
     , MAX(dtl.so_luong) so_luong
     , MAX(dtl.tong_tri_gia_ban) tong_tri_gia_ban
     , MAX(dtl.tong_tri_gia_tt_dtnt) tong_tri_gia_tt_dtnt
     , MAX(dtl.tsuat_dtnt) tsuat_dtnt
     , MAX(dtl.thue_duoc_ktru) thue_duoc_ktru
     , MAX(dtl.dchinh_tang_giam) dchinh_tang_giam
     , MAX(dtl.thue_pnop_tky_dtnt) thue_pnop_tky_dtnt
     , MAX(dtl.kieu_btt_id) kieu_btt_id
     , MAX(dtl.kieu_ten_ttdb) kieu_ten_ttdb
     , MAX(dtl.kieu_dvt_don_vi_tinh) kieu_dvt_don_vi_tinh
     , MAX(dtl.kieu_so_luong) kieu_so_luong
     , MAX(dtl.kieu_tong_tri_gia_ban) kieu_tong_tri_gia_ban
     , MAX(dtl.kieu_tong_tri_gia_tt_dtnt) kieu_tong_tri_gia_tt_dtnt
     , MAX(dtl.kieu_tsuat_dtnt) kieu_tsuat_dtnt
     , MAX(dtl.kieu_thue_duoc_ktru) kieu_thue_duoc_ktru
     , MAX(dtl.kieu_dchinh_tang_giam) kieu_dchinh_tang_giam
     , MAX(dtl.kieu_thue_pnop_tky_dtnt) kieu_thue_pnop_tky_dtnt
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.id id_tt, 
         tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         ctieu.ky_hieu_ctieu loai_ctieu,
    	 DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) btt_id,
    	 DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) ten_ttdb,
    	 DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) dvt_don_vi_tinh,
    	 DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) so_luong,
    	 DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) tong_tri_gia_ban,
    	 DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) tong_tri_gia_tt_dtnt,
         DECODE(gdien.cot_08, tkd.ky_hieu,  NVL(tkd.gia_tri,0), NULL) tsuat_dtnt,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) thue_duoc_ktru,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) dchinh_tang_giam,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) thue_pnop_tky_dtnt, 	
    	 DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_btt_id,
    	 DECODE(gdien.cot_03, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_ten_ttdb,
    	 DECODE(gdien.cot_04, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dvt_don_vi_tinh,
    	 DECODE(gdien.cot_05, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_so_luong,
    	 DECODE(gdien.cot_06, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_tong_tri_gia_ban,
    	 DECODE(gdien.cot_07, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_tong_tri_gia_tt_dtnt,
    	 DECODE(gdien.cot_08, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_tsuat_dtnt,    	
    	 DECODE(gdien.cot_09, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_thue_duoc_ktru,
    	 DECODE(gdien.cot_10, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dchinh_tang_giam,
    	 DECODE(gdien.cot_11, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_thue_pnop_tky_dtnt
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0501')
) dtl
WHERE (gd.loai_dlieu = '0501')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.loai_ctieu,
         dtl.row_id
/
