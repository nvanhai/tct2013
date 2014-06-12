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
	