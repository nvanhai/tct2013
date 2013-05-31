-- Start of DDL Script for View QLT_OWNER.RCV_V_TKHAI_GTGT_KT
-- Generated 8-Dec-2005 18:50:52 from QLT_OWNER@QLT_93

CREATE OR REPLACE VIEW rcv_v_tkhai_gtgt_kt (
   hdr_id,
   ctk_id,
   so_tt,
   ten_ctieu,
   doanhso_dtnt,
   sothue_dtnt,
   kieu_dlieu_ds,
   kieu_dlieu_st,
   ky_hieu_ctieu_ds,
   ky_hieu_ctieu_st )
AS
SELECT dtl.hdr_id
     , dtl.ctk_id
     , MAX(dtl.so_tt) so_tt
     , MAX(gd.ten_ctieu) ten_ctieu
     , MAX(dtl.doanhso_dtnt) doanhso_dtnt
     , MAX(dtl.sothue_dtnt) sothue_dtnt
     , MAX(dtl.kieu_dlieu_ds) kieu_dlieu_ds
     , MAX(dtl.kieu_dlieu_st) kieu_dlieu_st
     , MAX(dtl.ky_hieu_ctieu_ds) ky_hieu_ctieu_ds
     , MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id hdr_id,
         gdien.id id,
         gdien.so_tt so_tt,
         tkd.row_id row_id,
         gdien.ma_ctieu ctk_id,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) doanhso_dtnt,
    	 DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) sothue_dtnt,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ds,
    	 DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_st,
         DECODE(gdien.cot_01, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL) ky_hieu_ctieu_ds,
         DECODE(gdien.cot_02, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL) ky_hieu_ctieu_st
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0101')
	
) dtl
WHERE (gd.loai_dlieu = '0101')	
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         dtl.ctk_id
/

-- Create synonym RCV_V_TKHAI_GTGT_KT
CREATE PUBLIC SYNONYM rcv_v_tkhai_gtgt_kt
  FOR rcv_v_tkhai_gtgt_kt
/

-- Grants for View
GRANT SELECT ON rcv_v_tkhai_gtgt_kt TO qlt
/
GRANT SELECT ON rcv_v_tkhai_gtgt_kt TO qlt_read
/

-- End of DDL Script for View QLT_OWNER.RCV_V_TKHAI_GTGT_KT

