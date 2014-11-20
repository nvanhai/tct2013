create or replace view qlt_ntk.rcv_v_bc26_ac_blp_t1 as
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
    SL_da_su_dung 
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
            max(dtl.SL_da_su_dung)      SL_da_su_dung 
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
