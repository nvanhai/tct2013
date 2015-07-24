CREATE OR REPLACE VIEW QLT_NTK.RCV_V_BK26_02_AC AS
select hdr_id
      , row_id so_tt
     , ten_HD
     , Mau_so
     , Ky_hieu_HD
     , So_luong
     , Tu_so
     , Den_so
     , Loai_HD
     , 'T' TT_su_dung
     FROM
(
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.ten_HD) ten_HD
     , MAX(dtl.Mau_so) Mau_so
     , MAX(dtl.Ky_hieu_HD) Ky_hieu_HD
     , MAX(dtl.Tu_so) Tu_so
     , MAX(dtl.Den_so) Den_so
     , MAX(dtl.So_luong) So_luong
     , MAX(dtl.loaiHD) Loai_HD
FROM
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         dump(DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL)) ten_HD,
         dump(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Mau_so,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) Ky_hieu_HD,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) Tu_so,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) Den_so,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) So_luong,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) loaiHD
         FROM QLT_NTK.rcv_bcao_dtl_ac tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = 'BK26_02_AC')
    and (gdien.loai_dlieu ='BK26_02_AC')
) dtl
GROUP BY dtl.hdr_id,
         dtl.row_id,
         dtl.so_tt
);
