CREATE OR REPLACE VIEW QLT_NTK.RCV_V_BC26_AC AS
select 
     t1.hdr_id
     , t1.so_tt
     , t1.ten_HD
     , t1.Mau_so
     , t1.Ky_hieu_HD
     , t1.Tong_so
     , t1.Tu_so_ton_dk
     , t1.Den_so_ton_dk
     , t1.Tu_so_ps
     , t1.Den_so_ps
     , t1.Tu_so_SD_Mat_xoa
     , t1.Den_so_SD_Mat_xoa
     , t1.Tong_so_SD_Mat_xoa
     , t1.SL_da_su_dung
     , t2.SL_xoa
     , t2.So_xoa
     , t2.SL_mat
     , t2.So_mat
     , t2.SL_huy
     , t2.So_huy
     , t2.Tu_so_ton_ck
     , t2.Den_so_ton_ck
     , t2.SL_ton_ck
     , t2.Loai_HD
 from 
QLT_NTK.RCV_V_BC26_AC_T1 t1,
QLT_NTK.RCV_V_BC26_AC_T2 t2
where t1.hdr_id = t2.hdr_id
and t1.so_tt = t2.so_tt;
