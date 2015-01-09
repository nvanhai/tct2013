CREATE OR REPLACE VIEW QLT_NTK.RCV_V_01_AC AS
select hdr_id
      , row_id so_tt
     , DECODE(length(trim(MST_TC_dat_in)),13,substr(trim(MST_TC_dat_in),1,10)||'-'||substr(trim(MST_TC_dat_in),11,3),trim(MST_TC_dat_in)) MST_TC_dat_in
     , dump(Ten_TC_dat_in) Ten_TC_dat_in
     , dump(DC_TC_dat_in) DC_TC_dat_in
     , dump(So_HD) So_HD
     , Ngay_HD
     , dump(Ten_HD) Ten_HD
     , dump(Mau_so) Mau_so
     , dump(Ky_hieu_HD) Ky_hieu_HD
     , Tu_so
     , Den_so
     , So_luong
     , Loai_HD
     , Dtl_id
FROM
(
SELECT dtl.hdr_id
     , dtl.row_id
     , MAX(dtl.MST_TC_dat_in) MST_TC_dat_in
     , MAX(dtl.Ten_TC_dat_in) Ten_TC_dat_in
     , MAX(dtl.DC_TC_dat_in) DC_TC_dat_in
     , MAX(dtl.So_HD) So_HD
     , MAX(dtl.Ngay_HD) Ngay_HD
     , MAX(dtl.Ten_HD) Ten_HD
     , MAX(dtl.Mau_so) Mau_so
     , MAX(dtl.Ky_hieu_HD) Ky_hieu_HD
     , MAX(dtl.Tu_so) Tu_so
     , MAX(dtl.Den_so) Den_so
     , MAX(dtl.So_luong) So_luong
     , MAX(dtl.loaiHD) Loai_HD
     , MAX(dtl.dtl_id) Dtl_id
FROM
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) MST_TC_dat_in,
         (DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Ten_TC_dat_in,
         (DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL)) DC_TC_dat_in,
         (DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)) So_HD,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) Ngay_HD,
         (DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)) Ten_HD,
         (DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, ' ')) Mau_so,
         (DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, ' ')) Ky_hieu_HD,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) Tu_so,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) Den_so,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) So_luong,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) loaiHD,
         tkd.id dtl_id
         FROM QLT_NTK.rcv_bcao_dtl_ac tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_AC')
    and (gdien.loai_dlieu ='01_AC')
) dtl
GROUP BY dtl.hdr_id,
         dtl.row_id
);
