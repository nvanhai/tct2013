DELETE FROM rcv_map_ctieu
WHERE (loai_dlieu NOT LIKE '01%')
AND (loai_dlieu NOT LIKE '02%')
AND (loai_dlieu NOT LIKE '03%')
/
DELETE FROM rcv_gdien_tkhai
WHERE (loai_dlieu NOT LIKE '01%')
AND (loai_dlieu NOT LIKE '02%')
AND (loai_dlieu NOT LIKE '03%')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(411,'Kh«ng ph¸t sinh thuÕ TT§B trong kú','1','2','3','4',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0501',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(412,'Tæng céng',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'12','13','14',2,'0801',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(410,'Hµng ho¸ chÞu thuÕ tµi nguyªn trªn quyÕt to¸n','1','2','3','4','5','6','7','8','9','10','11',1,'0801',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(413,'Tæng céng',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'12','13','14',2,'0601',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(414,'Kª khai thuÕ TT§B ph¶i nép','5','6','7',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0501',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(415,'Hµng ho¸ b¸n trong n­íc chÞu thuÕ TT§B','8','9','10','11','12','13','14','15','16','17','18',NULL,'0501',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(416,'DÞch vô chÞu thuÕ TT§B','19','20','21','22','23','24','25','26','27','28','29',NULL,'0501',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(417,'Hµng ho¸ thuéc tr­êng hîp kh«ng ph¶i chÞu thuÕ TT§B','30','31','32',NULL,NULL,'33',NULL,NULL,NULL,NULL,NULL,NULL,'0501',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(418,'Tªn HHDV ®iÒu chØnh','1','2','3','4','5','6','7',NULL,NULL,NULL,NULL,1,'0502',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(419,'§¬n vÞ tÝnh','8','9','10','11','12','13','14',NULL,NULL,NULL,NULL,2,'0502','13')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(420,'S¶n l­îng tiªu thô','15','16','17','18','19','20','21',NULL,NULL,NULL,NULL,3,'0502','14')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(421,'Doanh sè b¸n (ch­a cã thuÕ GTGT)','22','23','24','25','26','27','28',NULL,NULL,NULL,NULL,4,'0502','15')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(422,'Gi¸ tÝnh thuÕ TT§B','29','30','31','32','33','34','35',NULL,NULL,NULL,NULL,5,'0502','16')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(423,'ThuÕ suÊt','36','37','38','39','40','41','42',NULL,NULL,NULL,NULL,6,'0502','17')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(424,'ThuÕ TT§B ®­îc khÊu trõ','43','44','45','46','47','48','49',NULL,NULL,NULL,NULL,7,'0502','18')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(425,'ThuÕ TT§B ph¶i nép','50','51','52','53','54','55','56',NULL,NULL,NULL,NULL,8,'0502','19')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(426,'Tæng céng','52','53','54',NULL,NULL,'55','56','57','58','59','60',NULL,'0501',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(427,'Tæng céng','57','58','59','60','61','62','63',NULL,NULL,NULL,NULL,9,'0502',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(429,'Hµng ho¸ xuÊt khÈu','34','35','36','37','38','39',NULL,NULL,NULL,NULL,NULL,NULL,'0501',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(430,'Hµng ho¸ b¸n ®Ó xuÊt khÈu','40','41','42','43','44','45',NULL,NULL,NULL,NULL,NULL,NULL,'0501',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(431,'Hµng ho¸ gia c«ng ®Ó xuÊt khÈu','46','47','48','49','50','51',NULL,NULL,NULL,NULL,NULL,NULL,'0501',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(409,'Hµng ho¸ chÞu thuÕ tµi nguyªn trªn tê khai','1','2','3','4','5','6','7','8','9','10','11',1,'0601',NULL)
/
INSERT INTO rcv_gdien_tkhai
VALUES
(396,'Gi¸ trÞ gia t¨ng ©m kú tr­íc chuyÓn qua','1',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,'0401','149')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(397,'Hµng ho¸ tån kho ®Çu kú','2',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'0401','17')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(398,'Hµng ho¸, dÞch vô mua trong kú','3',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,2,'0401','18')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(399,'Gi¸ vèn cña hµng ho¸, dÞch vô mua ngoµi sö dông cho SXKD hµng ho¸, dÞch vô chÞu thuÕ GTGT b¸n ra','4',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,3,'0401','19')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(400,'Tæng doanh sè HHDV chÞu thuÕ GTGT b¸n ra','5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,4,'0401','20')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(401,'Gi¸ trÞ gia t¨ng ph¸t sinh trong kú','6',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,5,'0401','21')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(402,'ThuÕ GTGT ph¶i nép','7',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,6,'0401','22')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(403,'Hµng ho¸ tån kho ®Çu n¨m','1',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'0701','1')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(404,'Hµng ho¸, dÞch vô mua trong n¨m','2',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,2,'0701','2')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(405,'Gi¸ vèn cña hµng ho¸, dÞch vô mua ngoµi sö dông cho SXKD hµng ho¸, dÞch vô chÞu thuÕ GTGT b¸n ra','3',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,3,'0701','3')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(406,'Tæng doanh sè hµng ho¸, dÞch vô chÞu thuÕ GTGT b¸n ra','4',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,4,'0701','4')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(407,'Gi¸ trÞ gia t¨ng ph¸t sinh trong n¨m (5=4-3)','5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,5,'0701','5')
/
INSERT INTO rcv_gdien_tkhai
VALUES
(408,'ThuÕ GTGT ph¶i nép (6=5 x thuÕ suÊt (%))','6',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,6,'0701','6')
/
