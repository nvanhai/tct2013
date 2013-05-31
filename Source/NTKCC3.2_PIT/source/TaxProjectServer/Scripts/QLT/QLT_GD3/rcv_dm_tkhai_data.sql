--Update kieu ky cho nhung to khai da trien khai truoc
UPDATE rcv_dm_tkhai SET kieu_ky = 'M' WHERE ma = '01'
/
UPDATE rcv_dm_tkhai SET kieu_ky = 'Q' WHERE ma = '02'
/
UPDATE rcv_dm_tkhai SET kieu_ky = 'Y' WHERE ma = '03'
/
--Insert nhung to khai moi cho GD3
INSERT INTO rcv_dm_tkhai
VALUES
('04','07A/GTGT Tê khai thuÕ gi¸ trÞ gia t¨ng trùc tiÕp','M')
/
INSERT INTO rcv_dm_tkhai
VALUES
('05','Tê khai thuÕ tiªu thô ®Æc biÖt','M')
/
INSERT INTO rcv_dm_tkhai
VALUES
('06','Tê khai thuÕ tµi nguyªn','M')
/
INSERT INTO rcv_dm_tkhai
VALUES
('07','Tê khai quyÕt to¸n n¨m thuÕ GTGT trùc tiÕp (mÉu sè 12A/GTGT)','Y')
/
INSERT INTO rcv_dm_tkhai
VALUES
('08','Tê khai quyÕt to¸n thuÕ tµi nguyªn','Y')
/
