-------------------------
---03/TD-TAIN------------
-------------------------
insert into RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('03_TD_TAIN13', 'T¿ khai thu¿ t¿i nguy¿n(D¿nh cho c¿ s¿ s¿n xu¿t Th¿y i¿n)', 'M', '320', to_date('14-06-2008', 'dd-mm-yyyy'), null);

	--RCV_GDIEN_TKHAI
insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5629, 'T¿i nguy¿n', '1', null, null, null, null, null, null, null, null, null, null, 1, '03_TD_TAIN13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5630, 'Thu¿ su¿t', '2', null, null, null, null, null, null, null, null, null, null, 2, '03_TD_TAIN13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5631, 'T¿ khai thu¿ t¿i nguy¿n', '23', '24', '25', '26', '27', '28', '29', null, null, null, null, 3, '03_TD_TAIN13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5809, 'T¿ khai thu¿ t¿i nguy¿n', null, null, null, null, '30', '31', '32', null, null, null, null, 4, '03_TD_TAIN13', null, null, null, null, null, null, null, null, null, null, null, null);

	--RCV_MAP_CTIEU	
insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '1', 'C', 5629, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '2', 'N', 5630, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '23', 'C', 5631, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '24', 'C', 5631, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '25', 'N', 5631, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '26', 'N', 5631, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '27', 'N', 5631, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '28', 'N', 5631, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '29', 'N', 5631, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '30', 'N', 5809, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '31', 'N', 5809, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_TD_TAIN13', '32', 'N', 5809, null);

-------------------------
---PL01-2/TD-GTGT--------
-------------------------
insert into RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('03_1_TD_TAIN13', 'PL 03-1/TD-TAIN', 'M', '320', to_date('14-06-2008', 'dd-mm-yyyy'), null);

	--RCV_GDIEN_TKHAI
insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5632, 'PL 03-1/TD-TAIN', '1', '2', '3', '4', '5', '6', '7', null, null, null, null, 1, '03_1_TD_TAIN13', null, null, null, null, null, null, null, null, null, null, null, null);

insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5633, 'T࠮g c࠮g', null, null, null, null, null, '8', null, null, null, null, null, 2, '03_1_TD_TAIN13', null, null, null, null, null, null, null, null, null, null, null, null);

	
	--RCV_MAP_CTIEU
insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_1_TD_TAIN13', '1', 'C', 5632, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_1_TD_TAIN13', '2', 'C', 5632, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_1_TD_TAIN13', '3', 'C', 5632, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_1_TD_TAIN13', '4', 'C', 5632, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_1_TD_TAIN13', '5', 'N', 5632, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_1_TD_TAIN13', '6', 'N', 5632, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_1_TD_TAIN13', '7', 'C', 5633, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('03_1_TD_TAIN13', '8', 'N', 5633, null);

