-- Start of DDL Script for View QLT_OWNER.RCV_V_TKHAI_QTOAN_TNDN
-- Generated 1/21/2006 12:45:01 PM from QLT_OWNER@QLT_91

CREATE OR REPLACE VIEW rcv_v_tkhai_qtoan_tndn (
   hdr_id,
   ctq_id,
   ky_hieu_ctieu,
   so_tt,
   ten_ctieu,
   so_dtnt,
   so_cqt,
   kieu_dlieu )
AS
SELECT dtl.hdr_id
     , dtl.ctq_id
     , dtl.ky_hieu_ctieu
     , MAX(dtl.so_tt) so_tt
     , MAX(gd.ten_ctieu) ten_ctieu
     , MAX(dtl.so_dtnt) so_dtnt
     , MAX(dtl.so_cqt) so_cqt
     , MAX(dtl.kieu_dlieu) kieu_dlieu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id hdr_id,
         gdien.id id,
         tkd.row_id row_id,
         gdien.ma_ctieu ctq_id,
         gdien.so_tt so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_dtnt,
    	 DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_cqt,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu,
    	 ctieu.ky_hieu_ctieu
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '0301')
) dtl
WHERE (gd.loai_dlieu = '0301')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         dtl.ky_hieu_ctieu,
         dtl.ctq_id
/


-- End of DDL Script for View QLT_OWNER.RCV_V_TKHAI_QTOAN_TNDN

