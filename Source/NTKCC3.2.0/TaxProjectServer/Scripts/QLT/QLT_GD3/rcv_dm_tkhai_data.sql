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
('04','07A/GTGT T� khai thu� gi� tr� gia t�ng tr�c ti�p','M')
/
INSERT INTO rcv_dm_tkhai
VALUES
('05','T� khai thu� ti�u th� ��c bi�t','M')
/
INSERT INTO rcv_dm_tkhai
VALUES
('06','T� khai thu� t�i nguy�n','M')
/
INSERT INTO rcv_dm_tkhai
VALUES
('07','T� khai quy�t to�n n�m thu� GTGT tr�c ti�p (m�u s� 12A/GTGT)','Y')
/
INSERT INTO rcv_dm_tkhai
VALUES
('08','T� khai quy�t to�n thu� t�i nguy�n','Y')
/
