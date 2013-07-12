-- Start of DDL Script for View QLT_OWNER.RCV_V_TKHAI_GTGT_TT
-- Generated 1-Apr-2006 12:45:14 from QLT_OWNER@QLT_93

CREATE OR REPLACE VIEW rcv_v_tkhai_gtgt_tt (
   hdr_id,
   so_tt,
   ctk_id,
   ten_ctieu,
   gia_tri,
   kieu_dlieu )
AS
SELECT dtl.hdr_id
     , dtl.so_tt
     , MAX(dtl.ctk_id) ctk_id
     , gd.ten_ctieu
     , MAX(dtl.gia_tri) gia_tri
     , MAX(dtl.kieu_dlieu) kieu_dlieu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.so_tt so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu,
    	 DECODE(gdien.cot_01, tkd.ky_hieu, gdien.ma_ctieu, NULL) ctk_id,
         gdien.id,
         tkd.row_id
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0401')
    AND (gdien.ma_ctieu IS NOT NULL)
) dtl
WHERE (gd.loai_dlieu = '0401')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.so_tt,
         dtl.row_id,
         gd.ten_ctieu
/

-- End of DDL Script for View QLT_OWNER.RCV_V_TKHAI_GTGT_TT

