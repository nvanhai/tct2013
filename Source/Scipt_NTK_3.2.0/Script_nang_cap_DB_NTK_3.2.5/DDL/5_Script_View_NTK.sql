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
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) SL_huy,
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

--View HDR cac to hoa don
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_BC26_HDR_SL
(id, ma_cqt, ma_cqt_cden, loai_bcao, ky_tungay, ky_denngay, bcao_tungay, bcao_denngay, dtnt_tin, thu_truong, ngay_nop_bcao, hthuc_nhap, ghi_chu, tthai_nhan, ngay_nhan, ngay_bc, loai_bc26, itkhai_id, phong_xly)
AS
SELECT  id,
        ma_cqt,
        TEN_CQ_TIEP_NHAN MA_CQT_CDEN,
        '01' Loai_Bcao,
        quy_bc ky_tungay,
        last_day(add_months(quy_bc,2)) ky_denngay,
        kybc_tu_ngay BCAO_TUNGAY,
        kybc_den_ngay BCAO_DENNGAY,
        DECODE(length(trim(tin)),13,substr(trim(tin),1,10)||'-'||substr(trim(tin),11,3),trim(tin)) DTNT_TIN,
        dump(nguoi_dai_dien) thu_truong,
        ngay_nop NGAY_NOP_BCAO,
        hthuc_nop HTHUC_NHAP,
        dump(ghi_chu) ghi_chu,
        '01',
        NGAY_CAP_NHAT,
        NGAY_BC,
        Loai_Bc26,
        itkhai_id ,
        PHONG_XLY
  FROM QLT_NTK.rcv_bcao_hdr_ac
    Where Loai_Bc = 'BC26_AC_SL'
        And Da_Nhan Is Null;	
		
-- cap nhat hdr cho 01_BK_BC26_AC
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_HDR_BK01_AC
(id, tin, loai_bc, ngay_nop, kybc_tu_ngay, kybc_den_ngay, ngay_cap_nhat, nguoi_cap_nhat, so_tt_tk, da_nhan, phong_xly, phong_qly, co_bang_ke, hthuc_nop, itkhai_id, ten_dv_cq, tin_dv_cq, ngay_bc, nguoi_dai_dien, ten_cq_tiep_nhan, ly_do_mat, ngay_mat_huy, phuong_phap_huy, dung_dn_cq, ghi_chu, ma_cqt, loai_bc26, nguoi_lap_bieu, quy_bc, ngay_tb_ph)
AS
SELECT ID, tin, loai_bc, ngay_nop, kybc_tu_ngay, kybc_den_ngay,
          ngay_cap_nhat, DUMP (nguoi_cap_nhat), so_tt_tk, da_nhan, phong_xly,
          phong_qly, co_bang_ke, hthuc_nop, itkhai_id, DUMP (ten_dv_cq), tin_dv_cq,
          ngay_bc, DUMP (nguoi_dai_dien), DUMP (ten_cq_tiep_nhan), DUMP (ly_do_mat),
          ngay_mat_huy, DUMP (phuong_phap_huy), dung_dn_cq, DUMP (ghi_chu), ma_cqt,
          loai_bc26, DUMP (nguoi_lap_bieu), quy_bc, ngay_tb_ph
     FROM QLT_NTK.rcv_bcao_hdr_ac
    WHERE loai_bc = '01_BK_BC26_AC' AND da_nhan IS NULL;
	
--restore
CREATE OR REPLACE VIEW RCV_V_HDR
(id, tin, loai_bc, ngay_nop, kybc_tu_ngay, kybc_den_ngay, ngay_cap_nhat, nguoi_cap_nhat, so_tt_tk, da_nhan, phong_xly, phong_qly, co_bang_ke, hthuc_nop, itkhai_id, ten_dv_cq, tin_dv_cq, ngay_bc, nguoi_dai_dien, ten_cq_tiep_nhan, ly_do_mat, ngay_mat_huy, phuong_phap_huy, dung_dn_cq, ghi_chu, ma_cqt, loai_bc26, nguoi_lap_bieu, quy_bc, ngay_tb_ph)
AS
SELECT ID, tin, loai_bc, ngay_nop, kybc_tu_ngay, kybc_den_ngay,
          ngay_cap_nhat, DUMP (nguoi_cap_nhat), so_tt_tk, da_nhan, phong_xly,
          phong_qly, co_bang_ke, hthuc_nop, itkhai_id, DUMP (ten_dv_cq), tin_dv_cq,
          ngay_bc, DUMP (nguoi_dai_dien), DUMP (ten_cq_tiep_nhan), DUMP (ly_do_mat),
          ngay_mat_huy, DUMP (phuong_phap_huy), dung_dn_cq, DUMP (ghi_chu), ma_cqt,
          loai_bc26, DUMP (nguoi_lap_bieu), quy_bc, ngay_tb_ph
     FROM QLT_NTK.rcv_bcao_hdr_ac
    WHERE loai_bc IN ('BC21_AC','01_TBAC','01_AC','03_TBAC') AND da_nhan IS NULL;

	
--updated 01_AC cho truong hop NULL: mau so, loai hoa don
CREATE OR REPLACE VIEW RCV_V_01_AC AS
select hdr_id
      , row_id so_tt
     , DECODE(length(trim(MST_TC_dat_in)),13,substr(trim(MST_TC_dat_in),1,10)||'-'||substr(trim(MST_TC_dat_in),11,3),trim(MST_TC_dat_in)) MST_TC_dat_in
     , Ten_TC_dat_in
     , DC_TC_dat_in
     , So_HD
     , Ngay_HD
     , Ten_HD
     , Mau_so
     , Ky_hieu_HD
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
         dump(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Ten_TC_dat_in,
         dump(DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL)) DC_TC_dat_in,
         dump(DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)) So_HD,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) Ngay_HD,
         dump(DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)) Ten_HD,
         dump(DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, ' ')) Mau_so,
         dump(DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, ' ')) Ky_hieu_HD,
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
