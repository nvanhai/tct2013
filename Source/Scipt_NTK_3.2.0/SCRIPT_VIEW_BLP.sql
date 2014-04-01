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
        LOAI_BLP,
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
    Loai_BLP ,
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
            MAX(dtl.loaiHD)        Loai_BLP ,
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
                    DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL)       loaiHD,
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
CREATE OR REPLACE VIEW RCV_V_01_TBAC_BLP AS
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
    Loai_BLP ,
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
            MAX(dtl.loaiHD)     Loai_BLP ,
            MAX(dtl.id)         dtl_id
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    NVL(tkd.row_id,0) row_id,
                    gdien.id          id,
                    gdien.so_tt,
                    (DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL)) ten_BLP,
                    (DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL)) Mau_so,
                    (DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL)) Ky_hieu_BLP,
                    DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)       So_luong,
                    DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)       Tu_so,
                    DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)       Den_so,
                    DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)       Ngay_BD_SD,
                    DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) So_BLP_in,
                    DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL) Ten_DN_in,
                    DECODE(gdien.cot_10, tkd.ky_hieu, tkd.gia_tri, NULL)       MST_DN_in,                   
                    DECODE(gdien.cot_11, tkd.ky_hieu, tkd.gia_tri, NULL)       Ngay_BLP_in,
                    DECODE(gdien.cot_12, tkd.ky_hieu, tkd.gia_tri, NULL)       loaiHD,
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
        LOAI_BLP,
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
    Loai_BLP ,
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
            MAX(dtl.loaiHD)     Loai_BLP ,
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
                    DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)       loaiHD,
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
CREATE OR REPLACE VIEW
    QLT_NTK.RCV_V_BC21_AC_BLP
    (
        HDR_ID,
        SO_TT,
        TEN_BLP,
        MAU_SO,
        KY_HIEU_BLP,
        TU_SO,
        DEN_SO,
        SO_LUONG,
        LIEN_BLP,
        GHI_CHU,
        LOAI_BLP,
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
    Lien_BLP ,
    Ghi_chu ,
    Loai_BLP ,
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
            MAX(dtl.Loai_BLP)    Loai_BLP ,
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
                    dump(DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)) Lien_BLP,
                    dump(DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL)) Ghi_chu,
                    DECODE(gdien.cot_09, tkd.ky_hieu, tkd.gia_tri, NULL)       Loai_BLP,
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
            dtl.row_id ) ;