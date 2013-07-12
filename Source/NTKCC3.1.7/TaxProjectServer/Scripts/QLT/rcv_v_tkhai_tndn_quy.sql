-- Start of DDL Script for View QLT_OWNER.RCV_V_TKHAI_TNDN_QUY
-- Generated 1/13/2006 11:34:59 AM from QLT_OWNER@QLT

CREATE OR REPLACE VIEW rcv_v_tkhai_tndn_quy (
   hdr_id,
   ctk_id,
   so_tt,
   ten_ctieu,
   so_dtnt,
   kieu_dlieu )
AS
SELECT dtl.hdr_id
     , MAX(dtl.ctk_id) ctk_id
     , dtl.so_tt
     , gd.ten_ctieu
     , MAX(dtl.so_dtnt) so_dtnt
     , MAX(dtl.kieu_dlieu) kieu_dlieu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         DECODE(gdien.cot_01, tkd.ky_hieu, gdien.ma_ctieu, NULL) ctk_id,
         gdien.so_tt so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_dtnt,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu,
         gdien.id,
         tkd.row_id
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0201')
    AND (gdien.ma_ctieu IS NOT NULL)
) dtl
WHERE (gd.loai_dlieu = '0201')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.so_tt,
         dtl.row_id,
         gd.ten_ctieu
/


-- End of DDL Script for View QLT_OWNER.RCV_V_TKHAI_TNDN_QUY

