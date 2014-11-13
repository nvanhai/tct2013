--------------------------------
--03/TNDN
--------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_QTOAN_TNDN_14 AS
SELECT dtl.hdr_id, dtl.ctk_id, MAX(dtl.so_tt) so_tt
, MAX(dtl.so_tien) so_tien
, MAX(dtl.kieu_dlieu) kieu_dlieu
, MAX(dtl.ky_hieu_ctieu) ky_hieu_ctieu
FROM QLT_NTK.rcv_gdien_tkhai gd,
(SELECT tkd.hdr_id hdr_id,
gdien.id id,
tkd.loai_dlieu loai_dlieu,
gdien.so_tt so_tt,
tkd.row_id row_id,
gdien.ma_ctieu ctk_id,
replace (DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL),',','.') so_tien,
DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu,
DECODE(gdien.cot_01, tkd.ky_hieu, ctieu.ky_hieu_ctieu, NULL) ky_hieu_ctieu
FROM QLT_NTK.rcv_tkhai_dtl tkd,
QLT_NTK.rcv_gdien_tkhai gdien,
QLT_NTK.rcv_map_ctieu ctieu
WHERE (ctieu.gdn_id = gdien.id)
AND (ctieu.ky_hieu = tkd.ky_hieu)
AND (tkd.loai_dlieu = '03_TNDN14')
) dtl
WHERE (gd.loai_dlieu = dtl.loai_dlieu)
--(gd.loai_dlieu = '03_TNDN11' OR gd.loai_dlieu = '01B_TNDN')
AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
dtl.ctk_id
order by so_tt;

	--PL 03 - 1x/TNDN
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_01ABC_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.so_tien) so_tien
     , dtl.ten_ctieu
     , dtl.ma_ctieu
     , dtl.so_tt
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu in ('03_1A_TNDN14','03_1B_TNDN14','03_1C_TNDN14'))
) dtl
WHERE (gd.loai_dlieu in ('03_1A_TNDN14','03_1B_TNDN14','03_1C_TNDN14'))
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
       --  dtl.row_id,
         dtl.so_tt;
	--PL03-2x/TNDN
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_2AB_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.nam_phat_sinh_lo) nam_phat_sinh_lo
     , MAX(dtl.so_lo_phat_sinh) so_lo_phat_sinh
     , MAX(dtl.so_lo_da_chuyen) so_lo_da_chuyen
     , MAX(dtl.so_lo_duoc_chuyen) so_lo_duoc_chuyen
     , MAX(dtl.so_lo_con_duoc_chuyen) so_lo_con_duoc_chuyen
     , dtl.ten_ctieu
     , dtl.ma_ctieu
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) nam_phat_sinh_lo,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) so_lo_phat_sinh,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) so_lo_da_chuyen,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_lo_duoc_chuyen,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) so_lo_con_duoc_chuyen
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu ='03_2A_TNDN14' or gdien.loai_dlieu ='03_2B_TNDN14')
) dtl
WHERE (gd.loai_dlieu = '03_2A_TNDN14' or gd.loai_dlieu = '03_2B_TNDN14')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
       --  dtl.row_id,
         dtl.so_tt;
	--PL 03-3A
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_3A_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.dieu_kien_uu_dai_hoac_so_tien) dieu_kien_uu_dai_hoac_so_tien
     , dtl.ten_ctieu
     , dtl.ma_ctieu
     , dtl.row_id
     , dtl.so_tt
     , dtl.kieu_dlieu
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         tkd.row_id,
         gdien.so_tt,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         ctieu.kieu_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) dieu_kien_uu_dai_hoac_so_tien
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu ='03_3A_TNDN14')
) dtl
WHERE (gd.loai_dlieu = '03_3A_TNDN14')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
         dtl.row_id,
         dtl.so_tt,
         dtl.kieu_dlieu;
	--PL 03
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_3B_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , dtl.ten_ctieu
     , dtl.row_id
     , MAX(dtl.dieu_kien_uu_dai_hoac_so_tien) dieu_kien_uu_dai_hoac_so_tien
     , dtl.ma_ctieu
     , dtl.so_tt
     , dtl.kieu_dlieu
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt ,
         tkd.row_id,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         ctieu.kieu_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) dieu_kien_uu_dai_hoac_so_tien
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu ='03_3B_TNDN14')
    order by gdien.so_tt
) dtl
WHERE (gd.loai_dlieu = '03_3B_TNDN14')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
         dtl.row_id,
         dtl.so_tt,
         dtl.kieu_dlieu;	
	--PL 03
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_3C_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.dieu_kien_uu_dai_hoac_so_tien) dieu_kien_uu_dai_hoac_so_tien
     , dtl.ten_ctieu
     , dtl.row_id
     , dtl.ma_ctieu
     , dtl.so_tt
     , dtl.kieu_dlieu
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         ctieu.kieu_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) dieu_kien_uu_dai_hoac_so_tien
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu ='03_3C_TNDN14')
    order by gdien.so_tt
) dtl
WHERE (gd.loai_dlieu = '03_3C_TNDN14')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
         dtl.row_id,
         dtl.so_tt,
         dtl.kieu_dlieu;	
	--PL 03
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_4_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.ten_dia_chi_NNT) ten_dia_chi_NNT
     , MAX(dtl.thu_nhap_nuoc_ngoai_ngoai_te) thu_nhap_nuoc_ngoai_ngoai_te
     , MAX(dtl.nuoc_ngoai_ten_ngoai_te) nuoc_ngoai_ten_ngoai_te
     , MAX(dtl.thu_nhap_nuoc_ngoai_dong_VN) thu_nhap_nuoc_ngoai_dong_VN
     , MAX(dtl.thue_thu_nhap_ngoai_te) thue_thu_nhap_ngoai_te
     , MAX(dtl.thue_thu_nhap_VND) thue_thu_nhap_VND
     , MAX(dtl.thu_nhap_chiu_thue_ngoai_te) thu_nhap_chiu_thue_ngoai_te
     , MAX(dtl.thu_nhap_chiu_thue_VND) thu_nhap_chiu_thue_VND
     , MAX(dtl.thue_suat_thue_TNDN) thue_suat_thue_TNDN
     , MAX(dtl.so_thue_phai_nop) so_thue_phai_nop
     , MAX(dtl.so_thue_da_khau_tru) so_thue_da_khau_tru
     , dtl.ten_ctieu
     , dtl.ma_ctieu
     , dtl.row_id
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_dia_chi_NNT,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) thu_nhap_nuoc_ngoai_ngoai_te,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) nuoc_ngoai_ten_ngoai_te,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) thu_nhap_nuoc_ngoai_dong_VN,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) thue_thu_nhap_ngoai_te,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) thue_thu_nhap_VND,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) thu_nhap_chiu_thue_ngoai_te,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) thu_nhap_chiu_thue_VND,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) thue_suat_thue_TNDN,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue_phai_nop,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue_da_khau_tru
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu ='03_4_TNDN14')
) dtl
WHERE (gd.loai_dlieu = '03_4_TNDN14')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
         dtl.row_id,
         dtl.so_tt;	
	--PL 03-5
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_5_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.so_tien) so_tien
     , dtl.ten_ctieu
     , dtl.row_id
     , dtl.ma_ctieu
     , dtl.so_tt
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu ='03_5_TNDN14')
    order by gdien.so_tt
) dtl
WHERE (gd.loai_dlieu = '03_5_TNDN14')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
         dtl.row_id,
         dtl.so_tt
order by dtl.so_tt;
	--PL 03
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_6_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , dtl.ten_ctieu
     , dtl.row_id
     , MAX(dtl.xac_dinh_so_trich_lap_quy_pt) xac_dinh_so_trich_lap_quy_pt
     , MAX(dtl.nam_trich_lap) nam_trich_lap
     , MAX(dtl.muc_trich_lap_trong_ky) muc_trich_lap_trong_ky
     , MAX(dtl.so_tien_trich_lap_trong_ky) so_tien_trich_lap_trong_ky
     , MAX(dtl.so_tien_da_su_dung) so_tien_da_su_dung
     , MAX(dtl.so_tien_chuyen_tu_cac_ky_truoc) so_tien_chuyen_tu_cac_ky_truoc
     , MAX(dtl.so_tien_chuyen_cac_ky_sau) so_tien_chuyen_cac_ky_sau
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) xac_dinh_so_trich_lap_quy_pt,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) nam_trich_lap,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) muc_trich_lap_trong_ky,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_trich_lap_trong_ky,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_da_su_dung,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_chuyen_tu_cac_ky_truoc,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_chuyen_cac_ky_sau
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu ='03_6_TNDN14')
) dtl
WHERE (gd.loai_dlieu = '03_6_TNDN14')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
         dtl.row_id,
         dtl.so_tt
;
	--PL 03-7
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_7_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , dtl.ten_ctieu
     , dtl.so_tt
     , dtl.row_id
	 , dtl.ma_ctieu
     , MAX(dtl.noi_dung) noi_dung
     , MAX(dtl.dthu_gia_tri_theo_ke_toan) dthu_gia_tri_theo_ke_toan
     , MAX(dtl.dthu_gia_tri_theo_thi_truong) dthu_gia_tri_theo_thi_truong
     , MAX(dtl.dthu_chenh_lech) dthu_chenh_lech
     , MAX(dtl.dthu_ppxd_gia_chi_phi) dthu_ppxd_gia_chi_phi
     , MAX(dtl.cphi_gia_tri_theo_ke_toan) cphi_gia_tri_theo_ke_toan
     , MAX(dtl.cphi_gia_tri_theo_thi_truong) cphi_gia_tri_theo_thi_truong
     , MAX(dtl.cphi_chenh_lech) cphi_chenh_lech
     , MAX(dtl.cphi_ppxd_gia_chi_phi) cphi_ppxd_gia_chi_phi
     , MAX(dtl.loi_nhuan_tang) loi_nhuan_tang
     , MAX(dtl.ten_ben_lk) ten_ben_lk
     , MAX(dtl.dia_chi) dia_chi
	 , Decode(Length(MAX(dtl.ma_so_thue)),13,SUBSTR(MAX(dtl.ma_so_thue),1,10) || '-' || SUBSTR(MAX(dtl.ma_so_thue),11,3),MAX(dtl.ma_so_thue)) ma_so_thue
     , MAX(dtl.hinh_thuc_lien_ket_A) hinh_thuc_lien_ket_A
     , MAX(dtl.hinh_thuc_lien_ket_B) hinh_thuc_lien_ket_B
     , MAX(dtl.hinh_thuc_lien_ket_C) hinh_thuc_lien_ket_C
     , MAX(dtl.hinh_thuc_lien_ket_D) hinh_thuc_lien_ket_D
     , MAX(dtl.hinh_thuc_lien_ket_E) hinh_thuc_lien_ket_E
     , MAX(dtl.hinh_thuc_lien_ket_F) hinh_thuc_lien_ket_F
     , MAX(dtl.hinh_thuc_lien_ket_G) hinh_thuc_lien_ket_G
     , MAX(dtl.hinh_thuc_lien_ket_H) hinh_thuc_lien_ket_H
     , MAX(dtl.hinh_thuc_lien_ket_I) hinh_thuc_lien_ket_I
     , MAX(dtl.hinh_thuc_lien_ket_J) hinh_thuc_lien_ket_J
     , MAX(dtl.hinh_thuc_lien_ket_K) hinh_thuc_lien_ket_K
     , MAX(dtl.hinh_thuc_lien_ket_L) hinh_thuc_lien_ket_L
     , MAX(dtl.hinh_thuc_lien_ket_M) hinh_thuc_lien_ket_M
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         DECODE(gdien.cot_01, DECODE(tkd.ky_hieu,1,0,tkd.ky_hieu), tkd.gia_tri, NULL) noi_dung,
         DECODE(gdien.cot_02, DECODE(tkd.ky_hieu,2,0,tkd.ky_hieu), tkd.gia_tri, NULL) dthu_gia_tri_theo_ke_toan,
         DECODE(gdien.cot_03, DECODE(tkd.ky_hieu,3,0,tkd.ky_hieu), tkd.gia_tri, NULL) dthu_gia_tri_theo_thi_truong,
         DECODE(gdien.cot_04, DECODE(tkd.ky_hieu,4,0,tkd.ky_hieu), tkd.gia_tri, NULL) dthu_chenh_lech,
         DECODE(gdien.cot_05, DECODE(tkd.ky_hieu,5,0,tkd.ky_hieu), tkd.gia_tri, NULL) cphi_gia_tri_theo_ke_toan,
         DECODE(gdien.cot_06, DECODE(tkd.ky_hieu,6,0,tkd.ky_hieu), tkd.gia_tri, NULL) cphi_gia_tri_theo_thi_truong,
         DECODE(gdien.cot_07, DECODE(tkd.ky_hieu,7,0,tkd.ky_hieu), tkd.gia_tri, NULL) cphi_chenh_lech,
         DECODE(gdien.cot_08, DECODE(tkd.ky_hieu,8,0,tkd.ky_hieu), tkd.gia_tri, NULL) loi_nhuan_tang,
         DECODE(gdien.cot_09, DECODE(tkd.ky_hieu,9,0,tkd.ky_hieu), tkd.gia_tri, NULL) dthu_ppxd_gia_chi_phi,
         DECODE(gdien.cot_10, DECODE(tkd.ky_hieu,10,0,tkd.ky_hieu), tkd.gia_tri, NULL) cphi_ppxd_gia_chi_phi,
         DECODE(gdien.cot_01, DECODE(tkd.ky_hieu,1,tkd.ky_hieu,0), tkd.gia_tri, NULL) ten_ben_lk,
         DECODE(gdien.cot_02, DECODE(tkd.ky_hieu,2,tkd.ky_hieu,0), tkd.gia_tri, NULL) dia_chi,
         DECODE(gdien.cot_03, DECODE(tkd.ky_hieu,3,tkd.ky_hieu,0), tkd.gia_tri, NULL) ma_so_thue,
         DECODE(gdien.cot_04, DECODE(tkd.ky_hieu,4,tkd.ky_hieu,0), tkd.gia_tri, NULL) hinh_thuc_lien_ket_A,
         DECODE(gdien.cot_05, DECODE(tkd.ky_hieu,5,tkd.ky_hieu,0), tkd.gia_tri, NULL) hinh_thuc_lien_ket_B,
         DECODE(gdien.cot_06, DECODE(tkd.ky_hieu,6,tkd.ky_hieu,0), tkd.gia_tri, NULL) hinh_thuc_lien_ket_C,
         DECODE(gdien.cot_07, DECODE(tkd.ky_hieu,7,tkd.ky_hieu,0), tkd.gia_tri, NULL) hinh_thuc_lien_ket_D,
         DECODE(gdien.cot_08, DECODE(tkd.ky_hieu,8,tkd.ky_hieu,0), tkd.gia_tri, NULL) hinh_thuc_lien_ket_E,
         DECODE(gdien.cot_09, DECODE(tkd.ky_hieu,9,tkd.ky_hieu,0), tkd.gia_tri, NULL) hinh_thuc_lien_ket_F,
         DECODE(gdien.cot_10, DECODE(tkd.ky_hieu,10,tkd.ky_hieu,0), tkd.gia_tri, NULL) hinh_thuc_lien_ket_G,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) hinh_thuc_lien_ket_H,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) hinh_thuc_lien_ket_I,
         DECODE(gdien.cot_13, tkd.ky_hieu, tkd.gia_tri, NULL) hinh_thuc_lien_ket_J,
         DECODE(gdien.cot_14, tkd.ky_hieu, tkd.gia_tri, NULL) hinh_thuc_lien_ket_K,
         DECODE(gdien.cot_15, tkd.ky_hieu, tkd.gia_tri, NULL) hinh_thuc_lien_ket_L,
         DECODE(gdien.cot_16, tkd.ky_hieu, tkd.gia_tri, NULL) hinh_thuc_lien_ket_M
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu ='03_7_TNDN14')
) dtl
WHERE (gd.loai_dlieu = '03_7_TNDN14')
  AND (dtl.id = gd.id)-- and row_id='1'
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
         dtl.row_id,
         dtl.so_tt
;	
	--PL 03-8
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_8_14 AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , dtl.row_id
     , dtl.so_tt
     , dtl.ma_ctieu
     , dtl.ten_ctieu
     , MAX(dtl.ten_doanh_nghiep) ten_doanh_nghiep
	 , Decode(Length(MAX(dtl.mst)),13,SUBSTR(MAX(dtl.mst),1,10) || '-' || SUBSTR(MAX(dtl.mst),11,3),MAX(dtl.mst)) mst
     , MAX(dtl.ty_le_phan_bo) ty_le_phan_bo
     , MAX(dtl.so_thue_tam_phan_bo_quy_I) so_thue_tam_phan_bo_quy_I
     , MAX(dtl.so_thue_tam_phan_bo_quy_II) so_thue_tam_phan_bo_quy_II
     , MAX(dtl.so_thue_tam_phan_bo_quy_III) so_thue_tam_phan_bo_quy_III
     , MAX(dtl.so_thue_tam_phan_bo_quy_IV) so_thue_tam_phan_bo_quy_IV
     , MAX(dtl.tong_so_thue_tam_phan_bo) tong_so_thue_tam_phan_bo
     , MAX(dtl.phan_bo_tong_so_thue_phai_nop) phan_bo_tong_so_thue_phai_nop
     , MAX(dtl.phan_bo_so_thue_phai_nop) phan_bo_so_thue_phai_nop
     , MAX(dtl.co_quan_thue_quan_ly) co_quan_thue_quan_ly
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         gdien.ten_ctieu,
         gdien.ma_ctieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_doanh_nghiep,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) mst,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) ty_le_phan_bo,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue_tam_phan_bo_quy_I,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue_tam_phan_bo_quy_II,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue_tam_phan_bo_quy_III,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) so_thue_tam_phan_bo_quy_IV,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) tong_so_thue_tam_phan_bo,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) phan_bo_tong_so_thue_phai_nop,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) phan_bo_so_thue_phai_nop,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) co_quan_thue_quan_ly
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu ='03_8_TNDN14')
) dtl
WHERE (gd.loai_dlieu = '03_8_TNDN14')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.ten_ctieu,
         dtl.ma_ctieu,
         dtl.row_id,
         dtl.so_tt;	
	--PL 03-9
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_QTOAN_TNDN_9_14 AS
(
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            dtl.ten_ctieu,
            MAX (dtl.so_thu_tu) so_thu_tu,
            MAX (dtl.chi_tieu)     chi_tieu,
			Decode(Length(MAX(dtl.mst)),13,SUBSTR(MAX(dtl.mst),1,10) || '-' || SUBSTR(MAX(dtl.mst),11,3),MAX(dtl.mst)) mst,
            MAX (dtl.ty_le_phan_bo)     ty_le_phan_bo,
            MAX (dtl.so_thue_phai_nop)              so_thue_phai_nop,
            MAX (dtl.co_quan_thue_quan_ly)              co_quan_thue_quan_ly,
            MAX (dtl.CQT_PARENT_ID)              CQT_PARENT_ID
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    gdien.ten_ctieu,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) so_thu_tu,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) chi_tieu,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) mst,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) ty_le_phan_bo,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  so_thue_phai_nop,
                    DECODE (gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)  co_quan_thue_quan_ly,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL ) CQT_PARENT_ID
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '03_9_TD_TNDN14' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id,
            dtl.ten_ctieu
    );
--------------------------------
--02/TAIN
-------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_02_TAIN_14 AS
SELECT hdr_id
      ,row_id
      ,so_tt
      ,ddiem_kthac
      ,btn_id
      ,don_vi_tinh
      ,san_luong
      ,gia_tinh_thue
      ,tsuat_dtnt
      ,muc_thue_an_dinh
      ,thue_phai_nop_dtnt     
      ,thue_du_kien_duoc_mien_giam
      ,round((to_number(thue_phai_nop_dtnt) - to_number(thue_du_kien_duoc_mien_giam))) thue_tai_nguyen_phat_sinh
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
     , MAX(dtl.gia_tinh_thue) gia_tinh_thue
     , MAX(dtl.tsuat_dtnt) tsuat_dtnt
     , MAX(dtl.muc_thue_an_dinh) muc_thue_an_dinh
     , MAX(dtl.thue_phai_nop_dtnt) thue_phai_nop_dtnt
     , MAX(dtl.thue_du_kien_duoc_mien_giam) thue_du_kien_duoc_mien_giam
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
         replace(DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL),',','.') gia_tinh_thue,
         DECODE(gdien.cot_06, tkd.ky_hieu, NVL(tkd.gia_tri,0), NULL) tsuat_dtnt,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) muc_thue_an_dinh,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) thue_phai_nop_dtnt,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) thue_du_kien_duoc_mien_giam,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) thue_da_ke_khai,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) chenh_lech
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '02_TAIN14')
) dtl
WHERE (gd.loai_dlieu = '02_TAIN14')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id
         , dtl.so_tt
)
order by so_tt;
--------------------------------
--02/BVMT
-------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_02_BVMT_14 AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.don_vi_tinh) don_vi_tinh
     , MAX(dtl.so_luong) so_luong
     , MAX(dtl.muc_phi) muc_phi
     , (MAX(dtl.so_luong)*MAX(dtl.muc_phi)) phi_phai_nop
     , MAX(dtl.phi_ke_khai) phi_ke_khai
     , MAX(dtl.chenh_lech) chenh_lech
     , MAX(dtl.loai_khoang_san) loai_khoang_san
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) don_vi_tinh,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_luong,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) muc_phi,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) phi_ke_khai,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) chenh_lech,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) loai_khoang_san
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '02_BVMT14')
    and gdien.loai_dlieu =  '02_BVMT14'
   and ctieu.loai_dlieu =  '02_BVMT14'
) dtl
WHERE (gd.loai_dlieu = '02_BVMT14')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
          dtl.so_tt;
--------------------------------
--02/PHLP
--------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_02_PHLP AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.loai_phi) loai_phi
     , MAX(dtl.tieu_muc) tieu_muc
     , MAX(dtl.so_tien_thu_duoc) so_tien_thu_duoc
     , MAX(dtl.ty_le) ty_le
     , MAX(dtl.so_tien_trich_cho_che_do) so_tien_trich_cho_che_do
     , MAX(dtl.so_tien_phai_nop) so_tien_phai_nop
     , MAX(dtl.so_tien_da_ke_khai) so_tien_da_ke_khai
     , MAX(dtl.chenh_lech) chenh_lech
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) loai_phi,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) tieu_muc,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_thu_duoc,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) ty_le,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_trich_cho_che_do,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_phai_nop,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_da_ke_khai,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) chenh_lech
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '02_PHLP')
    and gdien.loai_dlieu =  '02_PHLP'
   and ctieu.loai_dlieu =  '02_PHLP'
) dtl
WHERE (gd.loai_dlieu = '02_PHLP')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
          dtl.so_tt;
--------------------------------
--03A_TD_TAIN
--------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_03A_TD_TAIN AS
(
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            dtl.ten_ctieu,
            MAX (dtl.NHA_MAY_TD) NHA_MAY_TD,
            Decode(Length(MAX(dtl.MA_SO_THUE)),13,SUBSTR(MAX(dtl.MA_SO_THUE),1,10) || '-' || SUBSTR(MAX(dtl.MA_SO_THUE),11,3),MAX(dtl.MA_SO_THUE)) MA_SO_THUE,
            MAX (dtl.SAN_LUONG)     SAN_LUONG,
            MAX (dtl.GIA_TINH_THUE)       GIA_TINH_THUE,
            MAX (dtl.THUE_PHAT_SINH)     THUE_PHAT_SINH,
            MAX (dtl.THUE_MIEN_GIAM)              THUE_MIEN_GIAM,
            MAX (dtl.THUE_PHAI_NOP)              THUE_PHAI_NOP,
            MAX (dtl.THUE_DA_KHAI)              THUE_DA_KHAI,
            MAX (dtl.THUE_CHENH_LECH)              THUE_CHENH_LECH
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    gdien.ten_ctieu,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) NHA_MAY_TD,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) SAN_LUONG,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) GIA_TINH_THUE,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL ) THUE_PHAT_SINH,
                    DECODE (gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_MIEN_GIAM,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_PHAI_NOP,
                    DECODE (gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_DA_KHAI,
                    DECODE (gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_CHENH_LECH
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '03A_TD_TAIN' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id,
            dtl.ten_ctieu
    );
	--PL 03A-1
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_03A_1_TD_TAIN AS
(        
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.STT) STT,
            MAX (dtl.CHI_TIEU)     CHI_TIEU,
            Decode(Length(MAX(dtl.MA_SO_THUE)),13,SUBSTR(MAX(dtl.MA_SO_THUE),1,10) || '-' || SUBSTR(MAX(dtl.MA_SO_THUE),11,3),MAX(dtl.MA_SO_THUE)) MA_SO_THUE,
            MAX (dtl.CQT_QUAN_LY)     CQT_QUAN_LY,
			MAX (dtl.CQT_PARENT_ID)     CQT_PARENT_ID,
            MAX (dtl.TY_LE_PHAN_BO)              TY_LE_PHAN_BO,
            MAX (dtl.THUE_PHAI_NOP)              THUE_PHAI_NOP
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) STT,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) CHI_TIEU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL ) CQT_QUAN_LY,
					          DECODE (gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL ) CQT_PARENT_ID,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)  TY_LE_PHAN_BO,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_PHAI_NOP
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '03A_1_TD_TAIN' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    );
--------------------------------
--01/PHLP
--------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_01_PHLP AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , MAX(dtl.loai_phi) loai_phi
     , MAX(dtl.tieu_muc) tieu_muc
     , MAX(dtl.so_tien_thu_duoc) so_tien_thu_duoc
     , MAX(dtl.ty_le) ty_le
     , MAX(dtl.so_tien_trich_cho_che_do) so_tien_trich_cho_che_do
     , MAX(dtl.so_tien_phai_nop) so_tien_phai_nop
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.so_tt,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) loai_phi,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) tieu_muc,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_thu_duoc,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) ty_le,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_trich_cho_che_do,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) so_tien_phai_nop
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '01_PHLP')
    and gdien.loai_dlieu =  '01_PHLP'
   and ctieu.loai_dlieu =  '01_PHLP'
) dtl
WHERE (gd.loai_dlieu = '01_PHLP')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
          dtl.so_tt;
--------------------------------
--02/NTNN
--------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_02_NTNN AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , dtl.ctq_id
     , MAX(dtl.ke_khai) ke_khai     
     , MAX(dtl.quyet_toan) quyet_toan
     , MAX(dtl.chenh_lech) chenh_lech
     , MAX(dtl.ghi_chu) ghi_chu
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.ma_ctieu ctq_id,
         gdien.so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ke_khai,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) quyet_toan,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) chenh_lech,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) ghi_chu
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '02_NTNN14')
    and gdien.loai_dlieu =  '02_NTNN14'
   and ctieu.loai_dlieu =  '02_NTNN14'
) dtl
WHERE (gd.loai_dlieu = '02_NTNN14')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
          dtl.so_tt,
          dtl.ctq_id;
	-- PL 02-1
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_02_1_NTNN AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.ten_nha_thau_nuoc_ngoai) ten_nha_thau_nuoc_ngoai
     , MAX(dtl.nuoc_cu_tru) nuoc_cu_tru
     , Decode(Length(MAX(dtl.ma_so_thue_VN)),13,SUBSTR(MAX(dtl.ma_so_thue_VN),1,10) || '-' || SUBSTR(MAX(dtl.ma_so_thue_VN),11,3),MAX(dtl.ma_so_thue_VN)) ma_so_thue_VN
     , Decode(Length(MAX(dtl.ma_so_thue_nuoc_ngoai)),13,SUBSTR(MAX(dtl.ma_so_thue_nuoc_ngoai),1,10) || '-' || SUBSTR(MAX(dtl.ma_so_thue_nuoc_ngoai),11,3),MAX(dtl.ma_so_thue_nuoc_ngoai)) ma_so_thue_nuoc_ngoai
     , MAX(dtl.so_hop_dong) so_hop_dong
     , MAX(dtl.noi_dung) noi_dung
     , MAX(dtl.dia_diem) dia_diem
     , MAX(dtl.thoi_han) thoi_han
     , MAX(dtl.gia_tri_nguyen_te_hop_dong) gia_tri_nguyen_te_hop_dong
     , MAX(dtl.gia_tri_quy_doi_hop_dong) gia_tri_quy_doi_hop_dong
     , MAX(dtl.gia_tri_nguyen_te_quyet_toan) gia_tri_nguyen_te_quyet_toan
     , MAX(dtl.gia_tri_quy_doi_quyet_toan) gia_tri_quy_doi_quyet_toan
     , MAX(dtl.so_luong_LD) so_luong_LD
     , dtl.row_id
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_nha_thau_nuoc_ngoai,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) nuoc_cu_tru,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) ma_so_thue_VN,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) ma_so_thue_nuoc_ngoai,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) so_hop_dong,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) noi_dung,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) dia_diem,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) thoi_han,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_nguyen_te_hop_dong,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_quy_doi_hop_dong,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_nguyen_te_quyet_toan,
         DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_quy_doi_quyet_toan,
         DECODE(gdien.cot_13, tkd.ky_hieu, tkd.gia_tri, NULL) so_luong_LD
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu = '02_1_NTNN14')
) dtl
WHERE (gd.loai_dlieu  = '02_1_NTNN14')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.row_id,
         dtl.so_tt
;
	-- PL 02-2
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_02_2_NTNN AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.ten_nha_thau_phu_VN) ten_nha_thau_phu_VN
     , Decode(Length(MAX(dtl.ma_so_thue)),13,SUBSTR(MAX(dtl.ma_so_thue),1,10) || '-' || SUBSTR(MAX(dtl.ma_so_thue),11,3),MAX(dtl.ma_so_thue)) ma_so_thue
     , MAX(dtl.nha_thau_NN_ky_hd) nha_thau_NN_ky_hd
     , MAX(dtl.hop_dong_so) hop_dong_so
     , MAX(dtl.noi_dung) noi_dung
     , MAX(dtl.dia_diem) dia_diem
     , MAX(dtl.thoi_han) thoi_han
     , MAX(dtl.gia_tri_nguyen_te_hop_dong) gia_tri_nguyen_te_hop_dong
     , MAX(dtl.gia_tri_quy_doi_hop_dong) gia_tri_quy_doi_hop_dong
     , MAX(dtl.gia_tri_nguyen_te_quyet_toan) gia_tri_nguyen_te_quyet_toan
     , MAX(dtl.gia_tri_quy_doi_quyet_toan) gia_tri_quy_doi_quyet_toan
     , dtl.row_id
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_nha_thau_phu_VN,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) ma_so_thue,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) nha_thau_NN_ky_hd,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) hop_dong_so,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) noi_dung,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) dia_diem,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) thoi_han,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_nguyen_te_hop_dong,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_quy_doi_hop_dong,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_nguyen_te_quyet_toan,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_quy_doi_quyet_toan
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu = '02_2_NTNN14')
) dtl
WHERE (gd.loai_dlieu  = '02_2_NTNN14')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.row_id,
         dtl.so_tt;
--------------------------------
--04/NTNN
--------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_04_NTNN AS
SELECT dtl.hdr_id
     , dtl.row_id
     , dtl.so_tt
     , dtl.ctq_id
     , MAX(dtl.ke_khai) ke_khai
     , MAX(dtl.quyet_toan) quyet_toan
     , MAX(dtl.chenh_lech) chenh_lech
     , MAX(dtl.ghi_chu) ghi_chu
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,0) row_id,
         gdien.id,
         gdien.ma_ctieu ctq_id,
         gdien.so_tt,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ke_khai,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) quyet_toan,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) chenh_lech,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) ghi_chu
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (tkd.loai_dlieu = '04_NTNN14')
    and gdien.loai_dlieu =  '04_NTNN14'
   and ctieu.loai_dlieu =  '04_NTNN14'
) dtl
WHERE (gd.loai_dlieu = '04_NTNN14')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.row_id,
          dtl.so_tt,
          dtl.ctq_id;
	--PL 04-1
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_04_1_NTNN AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , MAX(dtl.ten_nha_thau_phu_VN) ten_nha_thau_phu_VN
     , Decode(Length(MAX(dtl.ma_so_thue)),13,SUBSTR(MAX(dtl.ma_so_thue),1,10) || '-' || SUBSTR(MAX(dtl.ma_so_thue),11,3),MAX(dtl.ma_so_thue)) ma_so_thue
     , MAX(dtl.nha_thau_nuoc_ngoai_ky_hd) nha_thau_nuoc_ngoai_ky_hd
     , MAX(dtl.so_hd) so_hd
     , MAX(dtl.noi_dung_hop_dong) noi_dung_hop_dong
     , MAX(dtl.dia_diem) dia_diem
     , MAX(dtl.thoi_han) thoi_han
     , MAX(dtl.gia_tri_hd_nguyen_te) gia_tri_hd_nguyen_te
     , MAX(dtl.gia_tri_hd_gia_tri_tien_VN) gia_tri_hd_gia_tri_tien_VN
     , MAX(dtl.gia_tri_quyet_toan_nguyen_te) gia_tri_quyet_toan_nguyen_te
     , MAX(dtl.gia_tri_qt_gia_tri_tien_VN) gia_tri_qt_gia_tri_tien_VN
     , dtl.row_id
FROM QLT_NTK.rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         gdien.id,
         gdien.so_tt,
         tkd.row_id,
         gdien.loai_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) ten_nha_thau_phu_VN,
         DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) ma_so_thue,
         DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) nha_thau_nuoc_ngoai_ky_hd,
         DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) so_hd,
         DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) noi_dung_hop_dong,
         DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) dia_diem,
         DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) thoi_han,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_hd_nguyen_te,
         DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_hd_gia_tri_tien_VN,
         DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_quyet_toan_nguyen_te,
         DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL) gia_tri_qt_gia_tri_tien_VN
  FROM QLT_NTK.rcv_tkhai_dtl tkd,
       QLT_NTK.rcv_gdien_tkhai gdien,
       QLT_NTK.rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
  AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu = '04_1_NTNN14')
) dtl
WHERE (gd.loai_dlieu  = '04_1_NTNN14')
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.row_id,
         dtl.so_tt;
--------------------------------
--02/TNDN-DK
--------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_02_TNDN_DK AS
SELECT
    DTL.HDR_ID ,
    DTL.CTK_ID ,
    MAX(DTL.SO_TT)         SO_TT ,
    MAX(DTL.TEN_CTIEU)     TEN_CTIEU ,
    MAX(DTL.GIA_TRI)       GIA_TRI ,
    MAX(DTL.KY_HIEU_CTIEU) KY_HIEU_CTIEU
FROM
    (
        SELECT
            TKD.HDR_ID ,
            GDIEN.ID ,
            GDIEN.SO_TT ,
            TKD.ROW_ID ,
            GDIEN.MA_CTIEU                                                       CTK_ID ,
            GDIEN.TEN_CTIEU                                                      TEN_CTIEU ,
            REPLACE(DECODE(GDIEN.COT_01, TKD.KY_HIEU, TKD.GIA_TRI, NULL),'%','')   GIA_TRI ,
            DECODE(GDIEN.COT_01, TKD.KY_HIEU, GDIEN.KY_HIEU_CTIEU, NULL) KY_HIEU_CTIEU
        FROM
            QLT_NTK.RCV_TKHAI_DTL TKD,
            (
                SELECT
                    GD.*,
                    CT.KY_HIEU,
                    CT.KY_HIEU_CTIEU
                FROM
                    QLT_NTK.RCV_GDIEN_TKHAI GD,
                    QLT_NTK.RCV_MAP_CTIEU CT
                WHERE
                    CT.GDN_ID (+) = GD.ID
                AND GD.LOAI_DLIEU = '02_TNDN_DK' ) GDIEN
        WHERE
            GDIEN.KY_HIEU = TKD.KY_HIEU (+)
        AND TKD.LOAI_DLIEU (+)= '02_TNDN_DK' ) DTL
GROUP BY
    DTL.HDR_ID,
    DTL.CTK_ID;
	--PL 02-1
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_02_1_TNDN_DK AS
(
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
			Decode(Length(MAX(dtl.MA_SO_THUE)),13,SUBSTR(MAX(dtl.MA_SO_THUE),1,10) || '-' || SUBSTR(MAX(dtl.MA_SO_THUE),11,3),MAX(dtl.MA_SO_THUE)) MA_SO_THUE,
            MAX (dtl.TEN_NHA_THAU)     TEN_NHA_THAU,
            MAX (dtl.TY_LE_PHAN_BO)       TY_LE_PHAN_BO,
            MAX (dtl.SO_THUE_PHAT_SINH_PHAI_NOP)     SO_THUE_PHAT_SINH_PHAI_NOP,
            MAX (dtl.GHI_CHU)              GHI_CHU
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) TEN_NHA_THAU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) TY_LE_PHAN_BO,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) SO_THUE_PHAT_SINH_PHAI_NOP,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  GHI_CHU
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '02_1_TNDN_DK' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    );	
--------------------------------
--02/TAIN-DK
--------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_02_TAIN_DK AS
SELECT
    DTL.HDR_ID ,
    DTL.CTK_ID ,
    MAX(DTL.SO_TT)         SO_TT ,
    MAX(DTL.TEN_CTIEU)     TEN_CTIEU ,
    MAX(DTL.GIA_TRI)       GIA_TRI ,
    MAX(DTL.KY_HIEU_CTIEU) KY_HIEU_CTIEU
FROM
    (
        SELECT
            TKD.HDR_ID ,
            GDIEN.ID ,
            GDIEN.SO_TT ,
            TKD.ROW_ID ,
            GDIEN.MA_CTIEU                                                       CTK_ID ,
            GDIEN.TEN_CTIEU                                                      TEN_CTIEU ,
            REPLACE(DECODE(GDIEN.COT_01, TKD.KY_HIEU, TKD.GIA_TRI, NULL),'%','')   GIA_TRI ,
            DECODE(GDIEN.COT_01, TKD.KY_HIEU, GDIEN.KY_HIEU_CTIEU, NULL) KY_HIEU_CTIEU
        FROM
            QLT_NTK.RCV_TKHAI_DTL TKD,
            (
                SELECT
                    GD.*,
                    CT.KY_HIEU,
                    CT.KY_HIEU_CTIEU
                FROM
                    QLT_NTK.RCV_GDIEN_TKHAI GD,
                    QLT_NTK.RCV_MAP_CTIEU CT
                WHERE
                    CT.GDN_ID (+) = GD.ID
                AND GD.LOAI_DLIEU = '02_TAIN_DK' ) GDIEN
        WHERE
            GDIEN.KY_HIEU = TKD.KY_HIEU (+)
        AND TKD.LOAI_DLIEU (+)= '02_TAIN_DK' ) DTL
GROUP BY
    DTL.HDR_ID,
    DTL.CTK_ID;
	--PL 02-1
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_02_1_TAIN_DK AS
(
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
			Decode(Length(MAX(dtl.MA_SO_THUE)),13,SUBSTR(MAX(dtl.MA_SO_THUE),1,10) || '-' || SUBSTR(MAX(dtl.MA_SO_THUE),11,3),MAX(dtl.MA_SO_THUE)) MA_SO_THUE,
            MAX (dtl.TEN_NHA_THAU)     TEN_NHA_THAU,
            MAX (dtl.TY_LE_PHAN_BO)       TY_LE_PHAN_BO,
            MAX (dtl.SO_THUE_PHAT_SINH_PHAI_NOP)     SO_THUE_PHAT_SINH_PHAI_NOP,
            MAX (dtl.GHI_CHU)              GHI_CHU
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) TEN_NHA_THAU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) TY_LE_PHAN_BO,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) SO_THUE_PHAT_SINH_PHAI_NOP,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  GHI_CHU
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '02_1_TAIN_DK' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    );	
	--PL 02-2
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_02_2_TAIN_DK AS
(
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.don_vi_tinh) don_vi_tinh,
            MAX (dtl.ngay_khai_thac)     ngay_khai_thac,
            MAX (dtl.san_luong_khai_thac)       san_luong_khai_thac,
            MAX (dtl.ngay_xuat_ban)     ngay_xuat_ban,
            MAX (dtl.san_luong_xuat_ban)              san_luong_xuat_ban,
            MAX (dtl.gia_tinh_thue)              gia_tinh_thue,
            MAX (dtl.doanh_thu)              doanh_thu,
            MAX (dtl.ghi_chu)              ghi_chu
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) don_vi_tinh,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) ngay_khai_thac,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) san_luong_khai_thac,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL ) ngay_xuat_ban,
                    DECODE (gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)  san_luong_xuat_ban,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)  gia_tinh_thue,
                    DECODE (gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL)  doanh_thu,
                    DECODE (gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL)  ghi_chu
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '02_2_TAIN_DK' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    );
--------------------------------
--KHBS 2014
--------------------------------
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_KHBS14 AS
SELECT   dtl.hdr_id, dtl.so_tt so_tt, dtl.row_id row_id,
            MAX (dtl.ten_ctieu_dc) ten_ctieu_dc, MAX (dtl.ma_ctieu) ma_ctieu,
            MAX (dtl.so_da_kk) so_da_kk, MAX (dtl.so_dieu_chinh)
                                                                so_dieu_chinh,
            MAX (dtl.so_chenh_lech) so_chenh_lech
       FROM QLT_NTK.rcv_gdien_tkhai gd,
            (SELECT tkd.hdr_id, tkd.row_id row_id, gdien.ID,
                    gdien.so_tt so_tt,
                    DECODE (gdien.cot_01,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) ten_ctieu_dc,
                    DECODE (gdien.cot_02,
                            tkd.ky_hieu, '[' || tkd.gia_tri || ']',
                            NULL
                           ) ma_ctieu,
                    DECODE (gdien.cot_03,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) so_da_kk,
                    DECODE (gdien.cot_04,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) so_dieu_chinh,
                    DECODE (gdien.cot_05,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) so_chenh_lech
               FROM QLT_NTK.RCV_TKHAI_DTL tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
              WHERE (ctieu.gdn_id = gdien.ID)
                AND (ctieu.ky_hieu = tkd.ky_hieu)
                AND (   tkd.loai_dlieu = 'KHBS_01A_TNDN13'
                     OR tkd.loai_dlieu = 'KHBS_01B_TNDN13'
                     OR tkd.loai_dlieu = 'KHBS_03_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_02_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_01_TAIN13'
                     OR tkd.loai_dlieu = 'KHBS_01_TTDB13'
                     OR tkd.loai_dlieu = 'KHBS_03_NTNN13'
                     OR tkd.loai_dlieu = 'KHBS_05_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_01_BVMT13'
                     OR tkd.loai_dlieu = 'KHBS_04_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_02_TNDN13'
                     OR tkd.loai_dlieu = 'KHBS_01_PHXD13'
                     OR tkd.loai_dlieu = 'KHBS_01_NTNN13'
                     OR tkd.loai_dlieu = 'KHBS_01_TBVMT13'
                     OR tkd.loai_dlieu = 'KHBS_01A_TNDN_DK'
                     OR tkd.loai_dlieu = 'KHBS_01B_TNDN_DK'
                     OR tkd.loai_dlieu = 'KHBS_01_TD_GTGT'
                     OR tkd.loai_dlieu = 'KHBS_03_TD_TAIN'
                     --Bo sung to QT 2014
                     OR tkd.loai_dlieu = 'KHBS_03_TNDN14'
                     OR tkd.loai_dlieu = 'KHBS_02_TAIN14'
                     OR tkd.loai_dlieu = 'KHBS_02_BVMT14'
                     OR tkd.loai_dlieu = 'KHBS_02_PHLP'
                     OR tkd.loai_dlieu = 'KHBS_03A_TD_TAIN'
                     OR tkd.loai_dlieu = 'KHBS_01_PHLP'
                     OR tkd.loai_dlieu = 'KHBS_02_NTNN14'
                     OR tkd.loai_dlieu = 'KHBS_04_NTNN14'
                     OR tkd.loai_dlieu = 'KHBS_02_TNDN_DK'
                     OR tkd.loai_dlieu = 'KHBS_02_TAIN_DK'
                     --End QT
					 OR tkd.loai_dlieu = 'KHBS_01_TAIN_DK'                       
                    )) dtl
      WHERE (   gd.loai_dlieu = 'KHBS_01A_TNDN13'
             OR gd.loai_dlieu = 'KHBS_01B_TNDN13'
             OR gd.loai_dlieu = 'KHBS_03_GTGT13'
             OR gd.loai_dlieu = 'KHBS_02_GTGT13'
             OR gd.loai_dlieu = 'KHBS_01_TAIN13'
             OR gd.loai_dlieu = 'KHBS_01_TTDB13'
             OR gd.loai_dlieu = 'KHBS_03_NTNN13'
             OR gd.loai_dlieu = 'KHBS_05_GTGT13'
             OR gd.loai_dlieu = 'KHBS_01_BVMT13'
             OR gd.loai_dlieu = 'KHBS_04_GTGT13'
             OR gd.loai_dlieu = 'KHBS_02_TNDN13'
             OR gd.loai_dlieu = 'KHBS_01_PHXD13'
			 OR gd.loai_dlieu = 'KHBS_01_NTNN13'
             OR gd.loai_dlieu = 'KHBS_01_TBVMT13'
             OR gd.loai_dlieu = 'KHBS_01A_TNDN_DK'
             OR gd.loai_dlieu = 'KHBS_01B_TNDN_DK'
             OR gd.loai_dlieu = 'KHBS_01_TD_GTGT'
             OR gd.loai_dlieu = 'KHBS_03_TD_TAIN'
                     --Bo sung to QT 2014
                     OR gd.loai_dlieu = 'KHBS_03_TNDN14'
                     OR gd.loai_dlieu = 'KHBS_02_TAIN14'
                     OR gd.loai_dlieu = 'KHBS_02_BVMT14'
                     OR gd.loai_dlieu = 'KHBS_02_PHLP'
                     OR gd.loai_dlieu = 'KHBS_03A_TD_TAIN'
                     OR gd.loai_dlieu = 'KHBS_01_PHLP'
                     OR gd.loai_dlieu = 'KHBS_02_NTNN14'
                     OR gd.loai_dlieu = 'KHBS_04_NTNN14'
                     OR gd.loai_dlieu = 'KHBS_02_TNDN_DK'
                     OR gd.loai_dlieu = 'KHBS_02_TAIN_DK'
                     --End QT             
			 OR gd.loai_dlieu = 'KHBS_01_TAIN_DK'
            )
        AND (dtl.ID = gd.ID)
   GROUP BY dtl.hdr_id, dtl.so_tt, dtl.row_id;
commit;	