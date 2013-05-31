CREATE OR REPLACE VIEW rcv_v_pluc_ttdb_01c (
   hdr_id,
   btt_id,
   so_tt,
   id_tt,
   ten_ctieu,
   ten_hang_hoa,
   ctg_id,
   ky_ke_khai,
   so_kkhai,
   so_dchinh,
   so_clech_dtnt,
   ly_do_dchinh,
   kieu_btt_id,
   kieu_ten_hang_hoa,
   kieu_ky_ke_khai,
   kieu_so_kkhai,
   kieu_so_dchinh,
   kieu_so_clech_dtnt,
   kieu_ly_do_dchinh )
AS
SELECT dtl.hdr_id
     , MAX(dtl.btt_id) btt_id
     , gd.so_tt
     , MAX(dtl.id_tt) id_tt
     , MAX(dtl.ten_ctieu) ten_ctieu
     , MAX(dtl.ten_hang_hoa) ten_hang_hoa
     , MAX(dtl.ctg_id) ctg_id
     , MAX(dtl.ky_ke_khai) ky_ke_khai
     , MAX(dtl.so_kkhai) so_kkhai
     , MAX(dtl.so_dchinh) so_dchinh
     , MAX(dtl.so_clech_dtnt) so_clech_dtnt
     , MAX(dtl.ly_do_dchinh) ly_do_dchinh
     , MAX(dtl.kieu_btt_id) kieu_btt_id
     , MAX(dtl.kieu_ten_hang_hoa) kieu_ten_hang_hoa
     , MAX(dtl.kieu_ky_ke_khai) kieu_ky_ke_khai
     , MAX(dtl.kieu_so_kkhai) kieu_so_kkhai
     , MAX(dtl.kieu_so_dchinh) kieu_so_dchinh
     , MAX(dtl.kieu_so_clech_dtnt) kieu_so_clech_dtnt
     , MAX(dtl.kieu_ly_do_dchinh) kieu_ly_do_dchinh
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         tkd.row_id,
         tkd.id id_tt,
         gdien.id,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) btt_id,
         gdien.ten_ctieu,
         gdien.ma_ctieu ctg_id,
    	 DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) ten_hang_hoa,
    	 DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) ky_ke_khai,
    	 DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_kkhai,
    	 DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) so_dchinh,
    	 DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) so_clech_dtnt,
    	 DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) ly_do_dchinh,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_btt_id,
    	 DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_ten_hang_hoa,
    	 DECODE(gdien.cot_03, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_ky_ke_khai,
    	 DECODE(gdien.cot_04, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_so_kkhai,
    	 DECODE(gdien.cot_05, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_so_dchinh,
    	 DECODE(gdien.cot_06, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_so_clech_dtnt,
    	 DECODE(gdien.cot_07, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_ly_do_dchinh
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0502')
) dtl
WHERE (gd.loai_dlieu = '0502')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         gd.so_tt
         --dtl.ten_ctieu,
         --dtl.ctg_id
/

