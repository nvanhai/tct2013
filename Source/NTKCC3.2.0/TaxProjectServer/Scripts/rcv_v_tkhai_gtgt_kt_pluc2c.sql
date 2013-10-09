-- Start of DDL Script for View QLT_OWNER.RCV_V_TKHAI_GTGT_KT_PLUC2C
-- Generated 8-Dec-2005 18:51:21 from QLT_OWNER@QLT_93

CREATE OR REPLACE VIEW rcv_v_tkhai_gtgt_kt_pluc2c (
   hdr_id,
   so_tt,
   ctg_id,
   ten_ctieu,
   gia_tri_ctieu,
   kieu_dlieu_ctieu )
AS
SELECT dtl.hdr_id
     , dtl.so_tt
     , MAX(dtl.ctg_id) ctg_id
     , gd.ten_ctieu
     , MAX(dtl.gia_tri_ctieu) gia_tri_ctieu
     , MAX(dtl.kieu_dlieu_ctieu) kieu_dlieu_ctieu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_ctieu,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ctieu,
    	 DECODE(gdien.cot_01, tkd.ky_hieu, gdien.ma_ctieu, NULL) ctg_id,
         gdien.id,
         tkd.row_id
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0104')
    AND (gdien.ma_ctieu IS NOT NULL)
) dtl
WHERE (gd.loai_dlieu = '0104')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.so_tt,
         dtl.row_id,
         gd.ten_ctieu
/

-- Create synonym RCV_V_TKHAI_GTGT_KT_PLUC2C
CREATE PUBLIC SYNONYM rcv_v_tkhai_gtgt_kt_pluc2c
  FOR rcv_v_tkhai_gtgt_kt_pluc2c
/

-- Grants for View
GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2c TO qlt
/
GRANT SELECT ON rcv_v_tkhai_gtgt_kt_pluc2c TO qlt_read
/

-- End of DDL Script for View QLT_OWNER.RCV_V_TKHAI_GTGT_KT_PLUC2C

