CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_TNCN_09A AS
SELECT dtl.hdr_id
     , dtl.id
     , dtl.ctk_id
     , MAX(dtl.so_tt) so_tt
     , MAX(gd.ten_ctieu) ten_ctieu
     , MAX(dtl.so_tien) so_tien
     , MAX(dtl.kieu_dlieu_ds) kieu_dlieu_ds
     --, MAX(dtl.kieu_dlieu_st) kieu_dlieu_st
     , MAX(dtl.ky_hieu_ctieu_ds) ky_hieu_ctieu_ds
     --, MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
  FROM QLT_NTK.rcv_gdien_tkhai gd,
       (SELECT tkd.hdr_id hdr_id,
               gdien.id id,
               tkd.loai_dlieu loai_dlieu,
               gdien.so_tt so_tt,
               tkd.row_id row_id,
               gdien.ma_ctieu ctk_id,
               DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien,
               DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ds,
               --DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_st,
               DECODE(gdien.cot_01, tkd.ky_hieu, '['||ctieu.ky_hieu_ctieu||']', NULL) ky_hieu_ctieu_ds
               --DECODE(gdien.cot_02, tkd.ky_hieu, '['||ctieu.ky_hieu_ctieu||']', NULL) ky_hieu_ctieu_st
          FROM QLT_NTK.rcv_tkhai_dtl tkd, QLT_NTK.rcv_gdien_tkhai gdien, QLT_NTK.rcv_map_ctieu ctieu
         WHERE (ctieu.gdn_id = gdien.id)
           AND (ctieu.ky_hieu = tkd.ky_hieu)
           AND (tkd.loai_dlieu = '09A_TNCN11')
         ) dtl
 WHERE (gd.loai_dlieu = dtl.loai_dlieu)
   AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.ctk_id,
         dtl.id
;

CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_TNCN_09B AS
SELECT dtl.hdr_id
     , dtl.id
     , dtl.ctk_id
     , MAX(dtl.so_tt) so_tt
     , MAX(gd.ten_ctieu) ten_ctieu
     , MAX(dtl.so_tien) so_tien
     , MAX(dtl.kieu_dlieu_ds) kieu_dlieu_ds
     --, MAX(dtl.kieu_dlieu_st) kieu_dlieu_st
     , MAX(dtl.ky_hieu_ctieu_ds) ky_hieu_ctieu_ds
     --, MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
  FROM QLT_NTK.rcv_gdien_tkhai gd,
       (SELECT tkd.hdr_id hdr_id,
               gdien.id id,
               tkd.loai_dlieu loai_dlieu,
               gdien.so_tt so_tt,
               tkd.row_id row_id,
               gdien.ma_ctieu ctk_id,
               DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien,
               DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ds,
               --DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_st,
               DECODE(gdien.cot_01, tkd.ky_hieu, '['||ctieu.ky_hieu_ctieu||']', NULL) ky_hieu_ctieu_ds
               --DECODE(gdien.cot_02, tkd.ky_hieu, '['||ctieu.ky_hieu_ctieu||']', NULL) ky_hieu_ctieu_st
          FROM QLT_NTK.rcv_tkhai_dtl tkd, QLT_NTK.rcv_gdien_tkhai gdien, QLT_NTK.rcv_map_ctieu ctieu
         WHERE (ctieu.gdn_id = gdien.id)
           AND (ctieu.ky_hieu = tkd.ky_hieu)
           AND (tkd.loai_dlieu = '09B_TNCN11')
         ) dtl
 WHERE (gd.loai_dlieu = dtl.loai_dlieu)
   AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.ctk_id,
         dtl.id
;

CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_TNCN_09C AS
SELECT
    dtl.hdr_id ,
    MAX(dtl.row_id)               row_id ,
    MAX(dtl.so_tt)                so_tt ,
    MAX(dtl.ho_ten)               ho_ten ,
    MAX(dtl.ngay_sinh)            ngay_sinh,
    MAX(dtl.mst)                  mst,
    MAX(dtl.so_cmtnd_ho_chieu)    so_cmtnd_ho_chieu,
    MAX(dtl.quan_he)              quan_he,
    MAX(dtl.so_thang_duoc_giam_tru)              so_thang_duoc_giam_tru,
    MAX(dtl.thu_nhap_giam_tru)    thu_nhap_giam_tru

FROM
    QLT_NTK.rcv_gdien_tkhai gd,
    (
        SELECT
            tkd.hdr_id,
            tkd.loai_dlieu,
            gdien.id,
            tkd.row_id,
            gdien.so_tt                                          so_tt,
            DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ho_ten,
            DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) ngay_sinh,
            DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) mst,
            DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_cmtnd_ho_chieu,
            DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) quan_he,
            DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) so_thang_duoc_giam_tru,
            DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) thu_nhap_giam_tru
        FROM
            QLT_NTK.rcv_tkhai_dtl tkd,
            QLT_NTK.rcv_gdien_tkhai gdien,
            QLT_NTK.rcv_map_ctieu ctieu
        WHERE
            (
                ctieu.gdn_id = gdien.id)
        AND (
                ctieu.ky_hieu = tkd.ky_hieu)
        AND (
                tkd.loai_dlieu = '09C_TNCN11') ) dtl
WHERE
    (
        gd.loai_dlieu = dtl.loai_dlieu)
AND (
        dtl.id = gd.id)
GROUP BY
    dtl.hdr_id,
    dtl.row_id,
    dtl.so_tt;

CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_TNCN_09MT AS
SELECT dtl.hdr_id
     , dtl.id
     , dtl.ctk_id
     , MAX(dtl.so_tt) so_tt
     , MAX(gd.ten_ctieu) ten_ctieu
     , MAX(dtl.so_tien) so_tien
     , MAX(dtl.kieu_dlieu_ds) kieu_dlieu_ds
     --, MAX(dtl.kieu_dlieu_st) kieu_dlieu_st
     , MAX(dtl.ky_hieu_ctieu_ds) ky_hieu_ctieu_ds
     --, MAX(dtl.ky_hieu_ctieu_st) ky_hieu_ctieu_st
  FROM QLT_NTK.rcv_gdien_tkhai gd,
       (SELECT tkd.hdr_id hdr_id,
               gdien.id id,
               tkd.loai_dlieu loai_dlieu,
               gdien.so_tt so_tt,
               tkd.row_id row_id,
               gdien.ma_ctieu ctk_id,
               DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien,
               DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_ds,
               --DECODE(gdien.cot_02, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu_st,
               DECODE(gdien.cot_01, tkd.ky_hieu, '['||ctieu.ky_hieu_ctieu||']', NULL) ky_hieu_ctieu_ds
               --DECODE(gdien.cot_02, tkd.ky_hieu, '['||ctieu.ky_hieu_ctieu||']', NULL) ky_hieu_ctieu_st
          FROM QLT_NTK.rcv_tkhai_dtl tkd, QLT_NTK.rcv_gdien_tkhai gdien, QLT_NTK.rcv_map_ctieu ctieu
         WHERE (ctieu.gdn_id = gdien.id)
           AND (ctieu.ky_hieu = tkd.ky_hieu)
           AND (tkd.loai_dlieu = '09MT_TNCN11')
         ) dtl
 WHERE (gd.loai_dlieu = dtl.loai_dlieu)
   AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.ctk_id,
         dtl.id
;
--view TK BC26_SL
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_BC26_AC_SL AS
select hdr_id
     , row_id so_tt
     , ten_HD
     , Ky_hieu_mau
     , Hinh_thuc_hoa_don
     , So_luong_ton_dau_ky
     , So_luong_mua_phat_hanh
     , SL_su_dung
     , SL_xoa
     , SL_mat
     , SL_huy
     , Tong_so
     , Ton_cuoi_ky
     , Ghi_chu
	 , Loai_HD
FROM
(
SELECT dtl.hdr_id
     , dtl.row_id
     , MAX(dtl.ten_HD) ten_HD
     , MAX(dtl.Ky_hieu_mau) Ky_hieu_mau
     , MAX(dtl.Hinh_thuc_hoa_don) Hinh_thuc_hoa_don
     , MAX(dtl.So_luong_ton_dau_ky) So_luong_ton_dau_ky
     , MAX(dtl.So_luong_mua_phat_hanh) So_luong_mua_phat_hanh
     , MAX(dtl.SL_su_dung) SL_su_dung
     , MAX(dtl.SL_xoa) SL_xoa
     , MAX(dtl.SL_mat) SL_mat
     , MAX(dtl.SL_huy) SL_huy
     , MAX(dtl.Tong_so) Tong_so
     , MAX(dtl.Ton_cuoi_ky) Ton_cuoi_ky
     , MAX(dtl.Ghi_chu) Ghi_chu
	 , MAX(dtl.Loai_HD) Loai_HD
FROM
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         dump(DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL)) ten_HD,
         dump(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Ky_hieu_mau,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) Hinh_thuc_hoa_don,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) So_luong_ton_dau_ky,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) So_luong_mua_phat_hanh,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) SL_su_dung,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) SL_xoa,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) SL_mat,
         DECODE(gdien.cot_19, tkd.ky_hieu, tkd.gia_tri, NULL) SL_huy,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) Tong_so,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) Ton_cuoi_ky,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) Ghi_chu,
		 DECODE(gdien.cot_13, tkd.ky_hieu, tkd.gia_tri, NULL) Loai_HD
         FROM QLT_NTK.rcv_bcao_dtl_ac tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = 'BC26_AC_SL')
    and (gdien.loai_dlieu ='BC26_AC_SL')
) dtl
GROUP BY dtl.hdr_id,
         dtl.row_id
);

--view TK BC26_SL PL 01
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_BK26_01_AC_SL AS
select hdr_id
      , row_id so_tt
     , ten_HD
     , Mau_so
     , Ky_hieu_HD
     , Tu_so
     , Den_so
     , So_luong
     , Loai_HD
     ,TT_su_dung
FROM
(
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , dtl.id
     , MAX(dtl.ten_HD) ten_HD
     , MAX(dtl.Mau_so) Mau_so
     , MAX(dtl.Ky_hieu_HD) Ky_hieu_HD
     , MAX(dtl.Tu_so) Tu_so
     , MAX(dtl.Den_so) Den_so
     , MAX(dtl.So_luong) So_luong
     , MAX(dtl.loaiHD) Loai_HD
     , MAx(dtl.TT_su_dung) TT_su_dung
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
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) loaiHD,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) TT_su_dung
         FROM QLT_NTK.rcv_bcao_dtl_ac tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = 'BK26SL_01_AC')
    and (gdien.loai_dlieu ='BK26SL_01_AC')
) dtl
GROUP BY dtl.hdr_id,
         dtl.row_id,
         dtl.so_tt,
         dtl.id
)
where ky_hieu_hd is not null;

--view TK BC26_SL PL 02
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_BK26_02_AC_SL AS
select hdr_id
      , row_id so_tt
     , ten_HD
     , Mau_so
     , Ky_hieu_HD
     , Tu_so
     , Den_so
     , So_luong
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
    AND (tkd.loai_dlieu = 'BK26SL_02_AC')
    and (gdien.loai_dlieu ='BK26SL_02_AC')
) dtl
GROUP BY dtl.hdr_id,
         dtl.row_id,
         dtl.so_tt
);

--Bang ke 3.1.0
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_BK01_AC AS
select hdr_id
      , row_id so_tt
     , ten_HD
     , Mau_so
     , Ky_hieu_HD
     , So_luong
     , Tu_so
     , Den_so     
     , Loai_HD
     , Dtl_Id
FROM
(
SELECT dtl.hdr_id
     , dtl.row_id
     , MAX(dtl.ten_HD) ten_HD
     , MAX(dtl.Mau_so) Mau_so
     , MAX(dtl.Ky_hieu_HD) Ky_hieu_HD
     , MAX(dtl.So_luong) So_luong
     , MAX(dtl.Tu_so) Tu_so
     , MAX(dtl.Den_so) Den_so
     , MAX(dtl.loaiHD) Loai_HD
     , MAX(dtl.id) dtl_id
FROM
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id id,
         gdien.so_tt,
         dump(DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL)) ten_HD,
         dump(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Mau_so,
         dump(DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL)) Ky_hieu_HD,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) So_luong,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) Tu_so,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) Den_so,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) loaiHD,
         tkd.id dtl_id
         FROM QLT_NTK.rcv_bcao_dtl_ac tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_BK_BC26_AC')
    and (gdien.loai_dlieu ='01_BK_BC26_AC')
) dtl
GROUP BY dtl.hdr_id,
         dtl.row_id
);
	