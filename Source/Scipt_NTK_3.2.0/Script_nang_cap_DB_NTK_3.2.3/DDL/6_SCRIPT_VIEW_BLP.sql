--01_AC
CREATE OR REPLACE VIEW
    QLT_NTK.RCV_V_01_AC_BLP
    (
        HDR_ID,
        SO_TT,
        MST_TC_DAT_IN,
        TEN_TC_DAT_IN,
        DC_TC_DAT_IN,
        SO_BLP,
        NGAY_BLP,
        TEN_BLP,
        MAU_SO,
        KY_HIEU_BLP,
        TU_SO,
        DEN_SO,
        SO_LUONG,
        --LOAI_BLP,
        DTL_ID
    ) AS
SELECT
    hdr_id ,
    row_id so_tt ,
    DECODE(LENGTH(trim(MST_TC_dat_in)),13,SUBSTR(trim(MST_TC_dat_in),1,10)||'-'||SUBSTR(trim
    (MST_TC_dat_in),11,3),trim(MST_TC_dat_in)) MST_TC_dat_in ,
    Ten_TC_dat_in ,
    DC_TC_dat_in ,
    So_BLP ,
    Ngay_BLP ,
    Ten_BLP ,
    Mau_so ,
    Ky_hieu_BLP ,
    Tu_so ,
    Den_so ,
    So_luong ,
    --Loai_BLP ,
    Dtl_id
FROM
    (
        SELECT
            dtl.hdr_id ,
            dtl.row_id ,
            MAX(dtl.MST_TC_dat_in) MST_TC_dat_in ,
            MAX(dtl.Ten_TC_dat_in) Ten_TC_dat_in ,
            MAX(dtl.DC_TC_dat_in)  DC_TC_dat_in ,
            MAX(dtl.So_BLP)         So_BLP ,
            MAX(dtl.Ngay_BLP)       Ngay_BLP ,
            MAX(dtl.Ten_BLP)        Ten_BLP ,
            MAX(dtl.Mau_so)        Mau_so ,
            MAX(dtl.Ky_hieu_BLP)    Ky_hieu_BLP ,
            MAX(dtl.Tu_so)         Tu_so ,
            MAX(dtl.Den_so)        Den_so ,
            MAX(dtl.So_luong)      So_luong ,
            --MAX(dtl.loaiHD)        Loai_BLP ,
            MAX(dtl.dtl_id)        Dtl_id
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    NVL(tkd.row_id,0) row_id,
                    gdien.id,
                    gdien.so_tt,
                    DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL)       MST_TC_dat_in,
                    dump(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Ten_TC_dat_in,
                    dump(DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL)) DC_TC_dat_in,
                    dump(DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)) So_BLP,
                    DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)       Ngay_BLP,
                    dump(DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)) Ten_BLP,
                    dump(DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)) Mau_so,
                    dump(DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL)) Ky_hieu_BLP,
                    DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL)       Tu_so,
                    DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL)       Den_so,
                    DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL)       So_luong,
                    --DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL)       loaiHD,
                    tkd.id                                                     dtl_id
                FROM
                    QLT_NTK.rcv_bcao_dtl_ac tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.id)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '01_AC_BLP')
                AND (
                        gdien.loai_dlieu ='01_AC_BLP') ) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.row_id ) ;

--01_TBAC            
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_01_TBAC_BLP AS
SELECT
    hdr_id ,
    row_id so_tt ,
    ten_BLP ,
    Mau_so ,
    Ky_hieu_BLP ,
    Tu_so ,
    Den_so ,
    So_luong ,
    Ngay_BD_SD ,
    Ten_DN_in ,
    DECODE(LENGTH(trim(MST_DN_in)),13,SUBSTR(trim(MST_DN_in),1,10)||'-'||SUBSTR(trim(MST_DN_in),11,
    3),trim(MST_DN_in)) MST_DN_in ,
    So_BLP_in ,
    Ngay_BLP_in ,
    --Loai_BLP ,
    Dtl_Id
FROM
    (
        SELECT
            dtl.hdr_id ,
            dtl.row_id ,
            MAX(dtl.ten_BLP)     ten_BLP ,
            MAX(dtl.Mau_so)     Mau_so ,
            MAX(dtl.Ky_hieu_BLP) Ky_hieu_BLP ,
            MAX(dtl.So_luong)   So_luong ,
            MAX(dtl.Tu_so)      Tu_so ,
            MAX(dtl.Den_so)     Den_so ,
            MAX(dtl.Ngay_BD_SD) Ngay_BD_SD ,
            MAX(dtl.Ten_DN_in)  Ten_DN_in ,
            MAX(dtl.MST_DN_in)  MST_DN_in ,
            MAX(dtl.So_BLP_in)   So_BLP_in ,
            MAX(dtl.Ngay_BLP_in) Ngay_BLP_in ,
            --MAX(dtl.loaiHD)     Loai_BLP ,
            MAX(dtl.id)         dtl_id
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    NVL(tkd.row_id,0) row_id,
                    gdien.id          id,
                    gdien.so_tt,
                    dump(DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL)) ten_BLP,
                    dump(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Mau_so,
                    dump(DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL)) Ky_hieu_BLP,
                    DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)       So_luong,
                    DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)       Tu_so,
                    DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)       Den_so,
                    DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)       Ngay_BD_SD,
                    dump(DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL)) So_BLP_in,
                    dump(DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL)) Ten_DN_in,
                    DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL)       MST_DN_in,
                    DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL)       Ngay_BLP_in,
                    --DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL)       loaiHD,
                    tkd.id                                                     dtl_id
                FROM
                    QLT_NTK.rcv_bcao_dtl_ac tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.id)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '01_TBAC_BLP')
                AND (
                        gdien.loai_dlieu ='01_TBAC_BLP') ) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.row_id );
--03_TBAC            
CREATE OR REPLACE VIEW
    QLT_NTK.RCV_V_03_TBAC_BLP
    (
        HDR_ID,
        SO_TT,
        TEN_BLP,
        MAU_SO,
        KY_HIEU_BLP,
        TU_SO,
        DEN_SO,
        SO_LUONG,
        --LOAI_BLP,
        DTL_ID
    ) AS
SELECT
    hdr_id ,
    row_id so_tt ,
    ten_BLP ,
    Mau_so ,
    Ky_hieu_BLP ,
    Tu_so ,
    Den_so ,
    So_luong ,
    --Loai_BLP ,
    dtl_Id
FROM
    (
        SELECT
            dtl.hdr_id ,
            dtl.row_id ,
            MAX(dtl.ten_BLP)     ten_BLP ,
            MAX(dtl.Mau_so)     Mau_so ,
            MAX(dtl.Ky_hieu_BLP) Ky_hieu_BLP ,
            MAX(dtl.Tu_so)      Tu_so ,
            MAX(dtl.Den_so)     Den_so ,
            MAX(dtl.So_luong)   So_luong ,
            --MAX(dtl.loaiHD)     Loai_BLP ,
            MAX(dtl.Id)         dtl_Id
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    NVL(tkd.row_id,0) row_id,
                    gdien.id          id,
                    gdien.so_tt,
                    dump(DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL)) ten_BLP,
                    dump(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Mau_so,
                    dump(DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL)) Ky_hieu_BLP,
                    DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)       Tu_so,
                    DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)       Den_so,
                    DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)       So_luong,
                    --DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)       loaiHD,
                    tkd.id                                                     dtl_id
                FROM
                    QLT_NTK.rcv_bcao_dtl_ac tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.id)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '03_TBAC_BLP')
                AND (
                        gdien.loai_dlieu ='03_TBAC_BLP') ) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.row_id ) ;
--BC21_AC            
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_BC21_AC_BLP AS
SELECT
    hdr_id ,
    row_id so_tt ,
    ten_BLP ,
    Mau_so ,
    Ky_hieu_BLP ,
    Tu_so ,
    Den_so ,
    So_luong ,
    Lien_BLP ,
    Ghi_chu ,
    DTL_Id
FROM
    (
        SELECT
            dtl.hdr_id ,
            dtl.row_id ,
            MAX(dtl.ten_BLP)     ten_BLP ,
            MAX(dtl.Mau_so)     Mau_so ,
            MAX(dtl.Ky_hieu_BLP) Ky_hieu_BLP ,
            MAX(dtl.Tu_so)      Tu_so ,
            MAX(dtl.Den_so)     Den_so ,
            MAX(dtl.So_luong)   So_luong ,
            MAX(dtl.Lien_BLP)    Lien_BLP ,
            MAX(dtl.Ghi_chu)    Ghi_chu ,
            MAX(dtl.Id)         dtl_Id
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    NVL(tkd.row_id,0) row_id,
                    gdien.id          id,
                    gdien.so_tt,
                    dump(DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) ten_BLP,
                    dump(DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL)) Mau_so,
                    dump(DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL)) Ky_hieu_BLP,
                    DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)       Tu_so,
                    DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)       Den_so,
                    DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)       So_luong,
                    dump(DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)) Lien_BLP,
                    dump(DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL)) Ghi_chu,
                    tkd.id                                                     dtl_id
                FROM
                    QLT_NTK.rcv_bcao_dtl_ac tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.id)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = 'BC21_AC_BLP')
                AND (
                        gdien.loai_dlieu ='BC21_AC_BLP') ) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.row_id );

--BC26_AC
create or replace view QLT_NTK.rcv_v_bc26_ac_blp as
select
    hdr_id ,
    row_id so_tt ,
    ten_BLP ,
    Mau_so ,
    Ky_hieu_BLP ,
    Tong_so ,
    Tu_so_ton_dk ,
    Den_so_ton_dk ,
    Tu_so_ps ,
    Den_so_ps ,
    Tu_so_SD_Mat_xoa ,
    Den_so_SD_Mat_xoa ,
    Tong_so_SD_Mat_xoa ,
    SL_da_su_dung ,
    SL_xoa ,
    So_xoa ,
    SL_mat ,
    So_mat ,
    SL_huy ,
    So_huy ,
    Tu_so_ton_ck ,
    Den_so_ton_ck ,
    SL_ton_ck 
from
    (
        select
            dtl.hdr_id ,
            dtl.row_id ,
            max(dtl.ten_BLP)             ten_BLP ,
            max(dtl.Mau_so)             Mau_so ,
            max(dtl.Ky_hieu_BLP)         Ky_hieu_BLP ,
            max(dtl.Tong_so)            Tong_so ,
            max(dtl.Tu_so_ton_dk)       Tu_so_ton_dk ,
            max(dtl.Den_so_ton_dk)      Den_so_ton_dk ,
            max(dtl.Tu_so_ps)           Tu_so_ps ,
            max(dtl.Den_so_ps)          Den_so_ps ,
            max(dtl.Tu_so_SD_Mat_xoa)   Tu_so_SD_Mat_xoa ,
            max(dtl.Den_so_SD_Mat_xoa)  Den_so_SD_Mat_xoa ,
            max(dtl.Tong_so_SD_Mat_xoa) Tong_so_SD_Mat_xoa ,
            max(dtl.SL_da_su_dung)      SL_da_su_dung ,
            max(dtl.SL_xoa)             SL_xoa ,
            max(dtl.So_xoa)             So_xoa ,
            max(dtl.SL_mat)             SL_mat ,
            max(dtl.So_mat)             So_mat ,
            max(dtl.SL_huy)             SL_huy ,
            max(dtl.So_huy)             So_huy ,
            max(dtl.Tu_so_ton_ck)       Tu_so_ton_ck ,
            max(dtl.Den_so_ton_ck)      Den_so_ton_ck ,
            max(dtl.SL_ton_ck)          SL_ton_ck 
        from
            (
                select
                    tkd.hdr_id,
                    nvl(tkd.row_id,0) row_id,
                    gdien.id,
                    gdien.so_tt,
                    dump(decode(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, null)) ten_BLP,
                    dump(decode(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, null)) Mau_so,
                    dump(decode(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, null)) Ky_hieu_BLP,
                    decode(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, null)       Tong_so,
                    decode(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, null)       Tu_so_ton_dk,
                    decode(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, null)       Den_so_ton_dk,
                    decode(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, null)       Tu_so_ps,
                    decode(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, null)       Den_so_ps,
                    decode(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, null)       Tu_so_SD_Mat_xoa,
                    decode(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, null)       Den_so_SD_Mat_xoa,
                    decode(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, null)       Tong_so_SD_Mat_xoa,
                    decode(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, null)       SL_da_su_dung,
                    decode(gdien.cot_13, tkd.ky_hieu, tkd.gia_tri, null)       SL_xoa,
                    decode(gdien.cot_14, tkd.ky_hieu, tkd.gia_tri, null)       So_xoa,
                    decode(gdien.cot_15, tkd.ky_hieu, tkd.gia_tri, null)       SL_mat,
                    decode(gdien.cot_16, tkd.ky_hieu, tkd.gia_tri, null)       So_mat,
                    decode(gdien.cot_17, tkd.ky_hieu, tkd.gia_tri, null)       SL_huy,
                    decode(gdien.cot_18, tkd.ky_hieu, tkd.gia_tri, null)       So_huy,
                    decode(gdien.cot_19, tkd.ky_hieu, tkd.gia_tri, null)       Tu_so_ton_ck,
                    decode(gdien.cot_20, tkd.ky_hieu, tkd.gia_tri, null)       Den_so_ton_ck,
                    decode(gdien.cot_21, tkd.ky_hieu, tkd.gia_tri, null)       SL_ton_ck
                from
                    QLT_NTK.rcv_bcao_dtl_ac tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                where
                    (
                        ctieu.gdn_id = gdien.id)
                and (
                        ctieu.ky_hieu = tkd.ky_hieu)
                and (
                        tkd.loai_dlieu = 'BC26_AC_BLP')
                and (
                        gdien.loai_dlieu ='BC26_AC_BLP') ) dtl
        group by
            dtl.hdr_id,
            dtl.row_id );

--VIEW HDR BC26_AC_BLP
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_BC26_HDR_BLP
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
    Where Loai_Bc = 'BC26_AC_BLP'
        And Da_Nhan Is Null;
--VIEW HDR 
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_HDR_BLP
(id, tin, loai_bc, ngay_nop, kybc_tu_ngay, kybc_den_ngay, ngay_cap_nhat, nguoi_cap_nhat, so_tt_tk, da_nhan, phong_xly, phong_qly, co_bang_ke, hthuc_nop, itkhai_id, ten_dv_cq, tin_dv_cq, ngay_bc, nguoi_dai_dien, ten_cq_tiep_nhan, ly_do_mat, ngay_mat_huy, phuong_phap_huy, dung_dn_cq, ghi_chu, ma_cqt, loai_bc26, nguoi_lap_bieu, quy_bc, ngay_tb_ph)
AS
SELECT ID, tin, loai_bc, ngay_nop, kybc_tu_ngay, kybc_den_ngay,
          ngay_cap_nhat, DUMP (nguoi_cap_nhat), so_tt_tk, da_nhan, phong_xly,
          phong_qly, co_bang_ke, hthuc_nop, itkhai_id, DUMP (ten_dv_cq), tin_dv_cq,
          ngay_bc, DUMP (nguoi_dai_dien), DUMP (ten_cq_tiep_nhan), DUMP (ly_do_mat),
          ngay_mat_huy, DUMP (phuong_phap_huy), dung_dn_cq, DUMP (ghi_chu), ma_cqt,
          loai_bc26, DUMP (nguoi_lap_bieu), quy_bc, ngay_tb_ph
     FROM QLT_NTK.rcv_bcao_hdr_ac
    WHERE loai_bc IN ('BC21_AC_BLP','01_TBAC_BLP','01_AC_BLP','03_TBAC_BLP') AND da_nhan IS NULL;
			