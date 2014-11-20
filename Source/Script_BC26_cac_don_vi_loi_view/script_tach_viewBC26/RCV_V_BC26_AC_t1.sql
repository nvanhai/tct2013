CREATE VIEW QLT_NTK.RCV_V_BC26_AC_t1 AS
select hdr_id
      , row_id so_tt
     , ten_HD
     , Mau_so
     , Ky_hieu_HD
     , Tong_so
     , Tu_so_ton_dk
     , Den_so_ton_dk
     , Tu_so_ps
     , Den_so_ps
     , Tu_so_SD_Mat_xoa
     , Den_so_SD_Mat_xoa
     , Tong_so_SD_Mat_xoa
     , SL_da_su_dung    
FROM
(
SELECT dtl.hdr_id
     , dtl.row_id
     , MAX(dtl.ten_HD) ten_HD
     , MAX(dtl.Mau_so) Mau_so
     , MAX(dtl.Ky_hieu_HD) Ky_hieu_HD
     , MAX(dtl.Tong_so) Tong_so
     , MAX(dtl.Tu_so_ton_dk) Tu_so_ton_dk
     , MAX(dtl.Den_so_ton_dk) Den_so_ton_dk
     , MAX(dtl.Tu_so_ps) Tu_so_ps
     , MAX(dtl.Den_so_ps) Den_so_ps
     , MAX(dtl.Tu_so_SD_Mat_xoa) Tu_so_SD_Mat_xoa
     , MAX(dtl.Den_so_SD_Mat_xoa) Den_so_SD_Mat_xoa
     , MAX(dtl.Tong_so_SD_Mat_xoa) Tong_so_SD_Mat_xoa
     , MAX(dtl.SL_da_su_dung) SL_da_su_dung
FROM
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         dump(DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL)) ten_HD,
         dump(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Mau_so,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) Ky_hieu_HD,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) Tong_so,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) Tu_so_ton_dk,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) Den_so_ton_dk,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) Tu_so_ps,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) Den_so_ps,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) Tu_so_SD_Mat_xoa,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) Den_so_SD_Mat_xoa,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) Tong_so_SD_Mat_xoa,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) SL_da_su_dung,
         DECODE(gdien.cot_13, tkd.ky_hieu, tkd.gia_tri, NULL) SL_xoa,
         DECODE(gdien.cot_14, tkd.ky_hieu, tkd.gia_tri, NULL) So_xoa,
         DECODE(gdien.cot_15, tkd.ky_hieu, tkd.gia_tri, NULL) SL_mat,
         DECODE(gdien.cot_16, tkd.ky_hieu, tkd.gia_tri, NULL) So_mat,
         DECODE(gdien.cot_17, tkd.ky_hieu, tkd.gia_tri, NULL) SL_huy,
         DECODE(gdien.cot_18, tkd.ky_hieu, tkd.gia_tri, NULL) So_huy,
         DECODE(gdien.cot_19, tkd.ky_hieu, tkd.gia_tri, NULL) Tu_so_ton_ck,
         DECODE(gdien.cot_20, tkd.ky_hieu, tkd.gia_tri, NULL) Den_so_ton_ck,
         DECODE(gdien.cot_21, tkd.ky_hieu, tkd.gia_tri, NULL) SL_ton_ck,
         DECODE(gdien.cot_22, tkd.ky_hieu, tkd.gia_tri, NULL) loaiHD
         FROM QLT_NTK.RCV_BCAO_DTL_AC tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = 'BC26_AC')
    and (gdien.loai_dlieu ='BC26_AC')
) dtl
GROUP BY dtl.hdr_id,
         dtl.row_id
);
