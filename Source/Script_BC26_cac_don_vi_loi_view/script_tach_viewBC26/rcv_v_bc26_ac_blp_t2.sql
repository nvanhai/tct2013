create or replace view qlt_ntk.rcv_v_bc26_ac_blp_t2 as
select
    hdr_id ,
    row_id so_tt ,
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
