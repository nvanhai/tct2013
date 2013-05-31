-- Start of DDL Script for View QLT_OWNER.RCV_V_PLUC_QTOAN_TNDN_14
-- Generated 18-Jan-2006 18:17:55 from QLT_OWNER@QLT

CREATE OR REPLACE VIEW rcv_v_pluc_qtoan_tndn_14 (
   hdr_id,
   ten_dia_chi,
   tnhap_nte,
   tnhap_vnd,
   thue_nte,
   thue_vnd,
   tnhap_tndn_nte,
   tnhap_tndn_vnd,
   tsuat_tndn,
   thue_tndn,
   thue_ktru,
   kieu_dlieu_01,
   kieu_dlieu_02,
   kieu_dlieu_03,
   kieu_dlieu_04,
   kieu_dlieu_05,
   kieu_dlieu_06,
   kieu_dlieu_07,
   kieu_dlieu_08,
   kieu_dlieu_09,
   kieu_dlieu_10 )
AS
SELECT dtl.hdr_id
     , MAX(dtl.ten_dia_chi) ten_dia_chi
     , MAX(dtl.tnhap_nte) tnhap_nte
     , MAX(dtl.tnhap_vnd) tnhap_vnd
     , MAX(dtl.thue_nte) thue_nte
     , MAX(dtl.thue_vnd) thue_vnd
     , MAX(dtl.tnhap_tndn_nte) tnhap_tndn_nte
     , MAX(dtl.tnhap_tndn_vnd) tnhap_tndn_vnd
     , MAX(dtl.tsuat_tndn) tsuat_tndn
     , MAX(dtl.thue_tndn) thue_tndn
     , MAX(dtl.thue_ktru) thue_ktru
     , MAX(dtl.kieu_dlieu_01) kieu_dlieu_01
     , MAX(dtl.kieu_dlieu_02) kieu_dlieu_02
     , MAX(dtl.kieu_dlieu_03) kieu_dlieu_03
     , MAX(dtl.kieu_dlieu_04) kieu_dlieu_04
     , MAX(dtl.kieu_dlieu_05) kieu_dlieu_05
     , MAX(dtl.kieu_dlieu_06) kieu_dlieu_06
     , MAX(dtl.kieu_dlieu_07) kieu_dlieu_07
     , MAX(dtl.kieu_dlieu_08) kieu_dlieu_08
     , MAX(dtl.kieu_dlieu_09) kieu_dlieu_09
     , MAX(dtl.kieu_dlieu_10) kieu_dlieu_10
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         tkd.row_id,
         gdien.id,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_dia_chi,
    	 DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) tnhap_nte,
    	 DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) tnhap_vnd,
    	 DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) thue_nte,
    	 DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) thue_vnd,
    	 DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) tnhap_tndn_nte,
    	 DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) tnhap_tndn_vnd,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) tsuat_tndn,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) thue_tndn,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) thue_ktru,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_01,
    	 DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_02,
    	 DECODE(gdien.cot_03, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_03,
    	 DECODE(gdien.cot_04, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_04,
    	 DECODE(gdien.cot_05, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_05,
    	 DECODE(gdien.cot_06, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_06,
    	 DECODE(gdien.cot_07, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_07,
    	 DECODE(gdien.cot_08, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_08,
    	 DECODE(gdien.cot_09, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_09,    	
    	 DECODE(gdien.cot_10, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_10
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0315')
) dtl
WHERE (gd.loai_dlieu = '0315')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id
/

-- End of DDL Script for View QLT_OWNER.RCV_V_PLUC_QTOAN_TNDN_14

