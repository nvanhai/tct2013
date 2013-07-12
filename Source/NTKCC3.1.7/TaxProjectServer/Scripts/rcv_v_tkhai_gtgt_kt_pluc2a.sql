-- Start of DDL Script for View QLT_OWNER.RCV_V_TKHAI_GTGT_KT_PLUC2A
-- Generated 8-Dec-2005 18:51:01 from QLT_OWNER@QLT_93

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

-- Create synonym RCV_V_TKHAI_GTGT_KT_PLUC2A
CREATE PUBLIC SYNONYM rcv_v_tkhai_gtgt_kt_pluc2a
  FOR rcv_v_tkhai_gtgt_kt_pluc2a
/

-- Grants for View
GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2a TO qlt
/
GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2a TO qlt_read
/

-- End of DDL Script for View QLT_OWNER.RCV_V_TKHAI_GTGT_KT_PLUC2A

