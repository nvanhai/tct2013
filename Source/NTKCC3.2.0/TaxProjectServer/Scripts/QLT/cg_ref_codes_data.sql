--Connect user QLT_OWNER
DELETE FROM cg_ref_codes
WHERE (rv_domain = 'HTKK_ABOUT.VERSION')
/
INSERT INTO cg_ref_codes
VALUES
('HTKK_ABOUT.VERSION','1.2.0',NULL,NULL,'Phiªn b¶n')
/
COMMIT
/
