CREATE OR REPLACE VIEW RCV_V_TKHAI_03_TNDN AS
SELECT dtl.hdr_id, dtl.ctk_id, MAX(dtl.so_tt) so_tt
, MAX(dtl.so_dtnt) so_dtnt
, MAX(dtl.kieu_dlieu_ds) kieu_dlieu_ds
, MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
FROM rcv_gdien_tkhai gd,
(SELECT tkd.hdr_id hdr_id,
gdien.id id,
tkd.loai_dlieu loai_dlieu,
gdien.so_tt so_tt,
tkd.row_id row_id,
gdien.ma_ctieu ctk_id,
replace (DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL),',','.') so_dtnt,
DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ds,
DECODE(gdien.cot_01, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL) ky_hieu_ctieu_st
FROM QLT_NTK.rcv_tkhai_dtl tkd,
QLT_NTK.rcv_gdien_tkhai gdien,
QLT_NTK.rcv_map_ctieu ctieu
WHERE (ctieu.gdn_id = gdien.id)
AND (ctieu.ky_hieu = tkd.ky_hieu)
AND (tkd.loai_dlieu = '03_TNDN11')
) dtl
WHERE (gd.loai_dlieu = dtl.loai_dlieu)
--(gd.loai_dlieu = '03_TNDN11' OR gd.loai_dlieu = '01B_TNDN')
AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
dtl.ctk_id
;

CREATE OR REPLACE VIEW RCV_V_TKHAI_02_TAIN AS
SELECT hdr_id
      ,row_id
      ,so_tt
      ,ddiem_kthac
      ,btn_id
      ,don_vi_tinh
      ,san_luong
      ,gia_don_vi
      ,tsuat_dtnt
      ,gia_tt_don_vi
      ,decode(gia_don_vi,0,san_luong*gia_tt_don_vi,san_luong*gia_don_vi*(tsuat_dtnt/100)) thue_phai_nop_dtnt
      ,thue_da_ke_khai
      ,chenh_lech
FROM
(
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.ddiem_kthac) ddiem_kthac
     , MAX(dtl.btn_id) btn_id
     , MAX(dtl.don_vi_tinh) don_vi_tinh
     , MAX(dtl.san_luong) san_luong
     , MAX(dtl.gia_don_vi) gia_don_vi
     , MAX(dtl.tsuat_dtnt) tsuat_dtnt
     , MAX(dtl.gia_tt_don_vi) gia_tt_don_vi
     , MAX(dtl.thue_phai_nop_dtnt) thue_phai_nop_dtnt
     , MAX(dtl.thue_da_ke_khai) thue_da_ke_khai
     , MAX(dtl.chenh_lech) chenh_lech
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ddiem_kthac,
    	   DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) btn_id,
    	   DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) don_vi_tinh,
    	   DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) san_luong,
    	   replace(DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL),',','.') gia_don_vi,
    	   DECODE(gdien.cot_06, tkd.ky_hieu, NVL(tkd.gia_tri,0), NULL) tsuat_dtnt,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tt_don_vi,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) thue_phai_nop_dtnt,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) thue_da_ke_khai,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) chenh_lech
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '02_TAIN11')
) dtl
WHERE (gd.loai_dlieu = '02_TAIN11')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id
         , dtl.so_tt
);

