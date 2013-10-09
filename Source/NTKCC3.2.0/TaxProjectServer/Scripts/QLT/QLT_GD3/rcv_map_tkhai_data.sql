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
('04','02','07A/GTGT T� khai thu� gi� tr� gia t�ng tr�c ti�p','TK','02')
/
INSERT INTO rcv_map_tkhai
VALUES
('05','25','T� khai thu� ti�u th� ��c bi�t','TK','02')
/
INSERT INTO rcv_map_tkhai
VALUES
('06','24','T� khai thu� t�i nguy�n','TK','02')
/
INSERT INTO rcv_map_tkhai
VALUES
('07','02','T� khai quy�t to�n n�m thu� GTGT tr�c ti�p (m�u s� 12A/GTGT)','QT','03')
/
INSERT INTO rcv_map_tkhai
VALUES
('08','08','T� khai quy�t to�n thu� t�i nguy�n','QT','03')
/
