update qlt_ntk.rcv_bcao_hdr_ac set da_nhan= null 
where id in (
    select distinct hdr_id from qlt_ntk.rcv_bcao_dtl_ac d where d.loai_dlieu in ('BK26_02_AC','BK26_01_AC')) and ngay_cap_nhat >'22-jul-2014' and da_nhan='Y';
