-------------------------
---01/TBVMT--------------
-------------------------
insert into RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('01_TBVMT13', 'T¿ khai thu¿ b¿o v¿ m¿i tr¿¿ng', 'M', '320', to_date('01-01-1900', 'dd-mm-yyyy'), null);

	--RCV_GDIEN_TKHAI
insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5634, 'T¿ khai thu¿ b¿o v¿ m¿ tr¿¿ng', '1', '2', '3', '4', '5', null, null, null, null, null, null, 1, '01_TBVMT13', null, null, null, null, null, null, null, null, null, null, null, null);
	
	--RCV_MAP_CTIEU
insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TBVMT13', '1', 'C', 5634, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TBVMT13', '2', 'C', 5634, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TBVMT13', '3', 'N', 5634, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TBVMT13', '4', 'N', 5634, null);

insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('01_TBVMT13', '5', 'N', 5634, null);
	