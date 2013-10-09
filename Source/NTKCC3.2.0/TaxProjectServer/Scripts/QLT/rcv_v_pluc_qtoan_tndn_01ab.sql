CREATE OR REPLACE VIEW rcv_v_pluc_qtoan_tndn_01ab (
   hdr_id,
   loai_dlieu,
   row_id,
   nam_psinh,
   so_psinh,
   lo_chuyen_01,
   lo_chuyen_02,
   lo_chuyen_03,
   lo_chuyen_04,
   lo_chuyen_05,
   lo_chuyen_06,
   kieu_dlieu,
   so_tt )
AS
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , dtl.row_id
     , MAX(dtl.nam_psinh) nam_psinh
     , MAX(dtl.so_psinh) so_psinh
     , MAX(dtl.lo_chuyen_01) lo_chuyen_01
     , MAX(dtl.lo_chuyen_02) lo_chuyen_02
     , MAX(dtl.lo_chuyen_03) lo_chuyen_03
     , MAX(dtl.lo_chuyen_04) lo_chuyen_04
     , MAX(dtl.lo_chuyen_05) lo_chuyen_05
     , MAX(dtl.lo_chuyen_06) lo_chuyen_06
     , MAX(dtl.kieu_dlieu) kieu_dlieu
     , dtl.so_tt
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,1) row_id,
         gdien.id,
         gdien.so_tt,
         gdien.loai_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) nam_psinh,
    	 DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) so_psinh,
    	 DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) lo_chuyen_01,
    	 DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) lo_chuyen_02,
    	 DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) lo_chuyen_03,
    	 DECODE(gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL) lo_chuyen_04,
    	 DECODE(gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL) lo_chuyen_05,
         DECODE(gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL) lo_chuyen_06,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
    AND (gdien.loai_dlieu = tkd.loai_dlieu)	
    AND (gdien.loai_dlieu IN ('0302','0316'))
    AND (gdien.so_tt < 8)
) dtl
WHERE (gd.loai_dlieu IN ('0302','0316'))
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.row_id,
         dtl.so_tt
UNION ALL
SELECT dtl.hdr_id
     , dtl.loai_dlieu
     , dtl.row_id
     , MAX(dtl.nam_psinh) nam_psinh
     , MAX(dtl.so_psinh) so_psinh
     , MAX(dtl.lo_chuyen_01) lo_chuyen_01
     , MAX(dtl.lo_chuyen_02) lo_chuyen_02
     , MAX(dtl.lo_chuyen_03) lo_chuyen_03
     , NULL lo_chuyen_04
     , NULL lo_chuyen_05
     , NULL lo_chuyen_06
     , MAX(dtl.kieu_dlieu) kieu_dlieu
     , dtl.so_tt
FROM rcv_gdien_tkhai gd,
(
  SELECT tkd.hdr_id,
         NVL(tkd.row_id,1) row_id,
         gdien.id,
         gdien.so_tt,
         gdien.loai_dlieu,
         DECODE(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL) nam_psinh,
    	 DECODE(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL) so_psinh,
    	 DECODE(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL) lo_chuyen_01,
    	 DECODE(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL) lo_chuyen_02,
    	 DECODE(gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL) lo_chuyen_03,
    	 DECODE(gdien.cot_01, ctieu.ky_hieu, ctieu.kieu_dlieu, NULL) kieu_dlieu
  FROM rcv_tkhai_dtl tkd,
       rcv_gdien_tkhai gdien,
       rcv_map_ctieu ctieu
  WHERE (ctieu.gdn_id = gdien.id)
	AND (ctieu.ky_hieu = tkd.ky_hieu)
	AND (gdien.loai_dlieu = tkd.loai_dlieu)
    AND (gdien.loai_dlieu IN ('0302','0316'))
    AND (gdien.so_tt > 7)
) dtl
WHERE (gd.loai_dlieu IN ('0302','0316'))
  AND (dtl.id = gd.id)
GROUP BY dtl.hdr_id,
         dtl.loai_dlieu,
         dtl.row_id,
         dtl.so_tt
/

