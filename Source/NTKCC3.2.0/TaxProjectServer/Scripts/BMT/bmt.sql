--Connect bang user BMT_OWNER
grant select on bmt_nsd_nhom to login_user
/
grant select on BMT_NHOM_CHUC_NANG to login_user
/
grant select on bmt_chuc_nang to login_user
/
DELETE FROM bmt_dm_ud
WHERE MA_UD = 'HTKK'
/
DELETE FROM bmt_chuc_nang
WHERE MA_UD = 'HTKK'
/
INSERT INTO bmt_dm_ud
(MA_UD,MO_TA)
VALUES
('HTKK','H� th�ng nh�n d� li�u t� khai')
/
INSERT INTO bmt_chuc_nang
(MA_CHUC_NANG,MUC_MENU,TEN_FILE,TEN_CHUC_NANG,MUC_MENU_CHA,LOAI_CHUC_NANG,DUOC_CN,MA_UD)
VALUES
(1,'NHAN_DU_LIEU','NHAN_DU_LIEU','Nh�n d� li�u t� khai','NHAN_DU_LIEU','M',NULL,'HTKK')
/
DELETE FROM BMT_THAM_SO
WHERE MA_UD = 'HTKK'
/
INSERT INTO BMT_THAM_SO
(TEN,GIA_TRI,GHI_CHU,MA_UD)
VALUES
('OWNER_PASS','QLT_RECV','Owner Pass Of App HTKK.','HTKK')
/
INSERT INTO BMT_THAM_SO
(TEN,GIA_TRI,GHI_CHU,MA_UD)
VALUES
('OWNER','QLT_RECV','Owner Of App HTKK.','HTKK')
/
BEGIN
    BMT_PCK_BMHT.prc_get_key;
    UPDATE BMT_THAM_SO SET GIA_TRI = BMT_PCK_BMHT.fnc_encrypt_data('QLT_RECV') where MA_UD ='HTKK' AND TEN ='OWNER_PASS';
END;
/
COMMIT
/
EXIT
/
