DELETE FROM rcv_map_tkhai
WHERE (ma_tkhai IN ('04','05','06','07','08'))
/
UPDATE rcv_map_tkhai
SET nhom_hso = '02'
WHERE (ma_tkhai IN ('01','02'))
/
UPDATE rcv_map_tkhai
SET nhom_hso = '03'
WHERE (ma_tkhai = '03')
/
INSERT INTO rcv_map_tkhai
VALUES
('04','02','07A/GTGT Tê khai thuÕ gi¸ trÞ gia t¨ng trùc tiÕp','TK','02')
/
INSERT INTO rcv_map_tkhai
VALUES
('05','25','Tê khai thuÕ tiªu thô ®Æc biÖt','TK','02')
/
INSERT INTO rcv_map_tkhai
VALUES
('06','24','Tê khai thuÕ tµi nguyªn','TK','02')
/
INSERT INTO rcv_map_tkhai
VALUES
('07','02','Tê khai quyÕt to¸n n¨m thuÕ GTGT trùc tiÕp (mÉu sè 12A/GTGT)','QT','03')
/
INSERT INTO rcv_map_tkhai
VALUES
('08','08','Tê khai quyÕt to¸n thuÕ tµi nguyªn','QT','03')
/
