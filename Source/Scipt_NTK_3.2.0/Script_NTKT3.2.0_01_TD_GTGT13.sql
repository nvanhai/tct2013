-------------------------
---01/TD-GTGT------------
-------------------------
insert into RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('01_TD_GTGT13', 'T¿ KHAI THU¿ GI¿ TR¿ GIA TNG 01/T_GTGT', 'M', '320', to_date('14-06-2008', 'dd-mm-yyyy'), null);
	--RCV_GDIEN_TKHAI
insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5620, 'S¿n l¿¿ng i¿n (Kw/h)', '21', null, null, null, null, null, null, null, null, null, null, 1, '01_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5621, 'Gi¿ t¿nh thu¿ ', '22', null, null, null, null, null, null, null, null, null, null, 2, '01_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5622, 'T¿ng tr¿ gi¿ t¿nh thu¿ [23]=[21]x[22]', '23', null, null, null, null, null, null, null, null, null, null, 3, '01_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5623, 'Thu¿ su¿t (%)', '24', null, null, null, null, null, null, null, null, null, null, 4, '01_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5624, 'Thu¿ GTGT ¿u ra [25]=[23]x[24]', '25', null, null, null, null, null, null, null, null, null, null, 5, '01_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5625, 'S¿ thu¿ GTGT ¿u v¿o ¿¿c kh¿u tr¿ c¿a ho¿t ¿ng s¿n xu¿t i¿n', '26', null, null, null, null, null, null, null, null, null, null, 6, '01_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5626, 'S¿ thu¿ GTGT ph¿i n¿p [27]= [25]-[26]', '27', null, null, null, null, null, null, null, null, null, null, 7, '01_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);

	--RCV_MAP_CTIEU	
insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TD_GTGT13', '21', 'N', 5620, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TD_GTGT13', '22', 'N', 5621, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TD_GTGT13', '23', 'N', 5622, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TD_GTGT13', '24', 'N', 5623, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TD_GTGT13', '25', 'N', 5624, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TD_GTGT13', '26', 'N', 5625, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TD_GTGT13', '27', 'N', 5626, null);

-------------------------
---01-1/TD-GTGT----------
-------------------------
insert into RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('01_1_TD_GTGT13', '01-1/T-GTGT', 'M', '320', to_date('14-06-2008', 'dd-mm-yyyy'), null);

	--RCV_GDIEN_TKHAI
insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5810, 'PL 01_1_TD_GTGT13', null, '2', '3', '4', '5', '6', null, null, null, null, null, 1, '01_1_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5811, 'T࠮g c࠮g (PL 01_1_TD_GTGT13)', null, null, null, '7', '8', '9', null, null, null, null, null, 2, '01_1_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);
	
	--RCV_MAP_CTIEU	
insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_1_TD_GTGT13', '2', 'C', 5810, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_1_TD_GTGT13', '3', 'C', 5810, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_1_TD_GTGT13', '4', 'N', 5810, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_1_TD_GTGT13', '5', 'N', 5810, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_1_TD_GTGT13', '6', 'N', 5810, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_1_TD_GTGT13', '7', 'N', 5811, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_1_TD_GTGT13', '8', 'N', 5811, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_1_TD_GTGT13', '9', 'N', 5811, null);

-------------------------
---PL01-2/TD-GTGT--------
-------------------------
insert into RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('01_2_TD_GTGT13', 'BࠎG PHࠎ BߠSߠTHUߠGIߠTRߠGIA TNG PHࠉ N�? C�? CߠSߠSࠎ XUࠔ TH࠙ Iࠎ CHO', 'M', '320', to_date('14-06-2008', 'dd-mm-yyyy'), null);

	--RCV_GDIEN_TKHAI
insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5627, 'PL 01_2_TD_GTGT13', '1', '2', '3', '4', '5', '6', '7', null, null, null, null, 1, '01_2_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5628, 'T࠮g c࠮g (PL 01_2_TD_GTGT13)', null, null, null, null, null, '8', null, null, null, null, null, 2, '01_2_TD_GTGT13', null, null, null, null, null, null, null, null, null, null, null, null);
	
	--RCV_MAP_CTIEU
insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_2_TD_GTGT13', '1', 'C', 5627, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_2_TD_GTGT13', '2', 'C', 5627, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_2_TD_GTGT13', '3', 'C', 5627, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_2_TD_GTGT13', '4', 'C', 5627, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_2_TD_GTGT13', '5', 'N', 5627, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_2_TD_GTGT13', '6', 'N', 5627, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_2_TD_GTGT13', '7', 'C', 5627, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_2_TD_GTGT13', '8', 'N', 5628, null);
