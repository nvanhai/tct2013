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
('HTKK','HÖ thèng nhËn d÷ liÖu tê khai')
/
INSERT INTO bmt_chuc_nang
(MA_CHUC_NANG,MUC_MENU,TEN_FILE,TEN_CHUC_NANG,MUC_MENU_CHA,LOAI_CHUC_NANG,DUOC_CN,MA_UD)
VALUES
(1,'NHAN_DU_LIEU','NHAN_DU_LIEU','NhËn d÷ liÖu tê khai','NHAN_DU_LIEU','M',NULL,'HTKK')
/


