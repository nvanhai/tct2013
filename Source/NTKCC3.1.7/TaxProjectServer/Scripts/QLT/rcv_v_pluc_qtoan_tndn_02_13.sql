-- Start of DDL Script for View QLT_OWNER.RCV_V_PLUC_QTOAN_TNDN_02_13
-- Generated 2/6/2006 3:17:10 PM from QLT_OWNER@QLT_91

CREATE OR REPLACE VIEW rcv_v_pluc_qtoan_tndn_02_13 (
   hdr_id,
   row_id,
   loai_dlieu,
   dcp_id,
   so_tt,
   ten_ctieu,
   so_dtnt,
   kieu_dlieu,
   ky_hieu )
AS
SELECT dtl.hdr_id
     , dtl.row_id
     , gd.loai_dlieu
     , MAX(dtl.dcp_id) dcp_id
     , dtl.so_tt
     , gd.ten_ctieu
     , MAX(dtl.so_dtnt) so_dtnt
     , MAX(dtl.kieu_dlieu) kieu_dlieu
     , dtl.ky_hieu
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         DECODE(tkd.loai_dlieu,'0303',1,'0310',1,'0311',1
                              ,'0313',1,'0314',1,tkd.row_id) row_id,
         ctieu.ky_hieu_ctieu ky_hieu,
    	 DECODE(gdien.cot_01, tkd.ky_hieu, gdien.ma_ctieu, NULL) dcp_id,
         gdien.so_tt so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_dtnt,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
	AND (ctieu.loai_dlieu = tkd.loai_dlieu)
    AND (ctieu.loai_dlieu NOT LIKE '01%')
    AND (ctieu.loai_dlieu <> '0201')
    AND (ctieu.loai_dlieu NOT IN ('0301','0302','0315','0316'))
    AND (gdien.ma_ctieu IS NOT NULL)
) dtl
WHERE (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
         gd.loai_dlieu,
         dtl.so_tt,
         gd.ten_ctieu,
         dtl.ky_hieu
/

-- End of DDL Script for View QLT_OWNER.RCV_V_PLUC_QTOAN_TNDN_02_13

